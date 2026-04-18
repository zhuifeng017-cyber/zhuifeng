#!/usr/bin/env python3
"""
维运网船舶信息爬虫
从 weiyun001.com 根据英文船名批量提取：船长、船宽、船型、吃水

用法:
    python weiyun_scraper.py input.xlsx              # 读取 Excel，自动识别船名列
    python weiyun_scraper.py input.xlsx -c 船名列名  # 指定列名
    python weiyun_scraper.py -s KOWLOON MSC_OSCAR   # 直接指定船名
    python weiyun_scraper.py --gen-sample            # 生成示例输入文件
"""

import argparse
import asyncio
import json
import logging
import re
import time
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

SITE = "https://www.weiyun001.com"
# 礼貌间隔，避免被限速
REQUEST_DELAY = 2.5

# 目标字段及其在页面上的中文标签
TARGET_FIELDS = ["船长", "船宽", "船型", "吃水"]


# ---------------------------------------------------------------------------
# 核心提取逻辑
# ---------------------------------------------------------------------------

async def _wait_for_card(page) -> bool:
    """等待船舶信息卡片出现，返回是否成功."""
    card_selectors = [
        ".ship-info",
        ".vessel-info",
        ".ship-detail",
        "[class*='shipInfo']",
        "[class*='vesselInfo']",
        # 直接根据截图中出现的字段判断
        "text=船长",
    ]
    for sel in card_selectors:
        try:
            await page.locator(sel).first.wait_for(state="visible", timeout=8000)
            return True
        except PWTimeout:
            continue
    return False


async def _extract_by_label(page, field: str) -> str | None:
    """通过标签文本定位相邻数值."""
    try:
        # 策略1：找到包含该标签的元素，取其 nextSibling 文字
        val = await page.evaluate(
            """(field) => {
                const walker = document.createTreeWalker(
                    document.body, NodeFilter.SHOW_TEXT, null, false);
                let node;
                while ((node = walker.nextNode())) {
                    if (node.textContent.trim() === field) {
                        // 尝试同级下一节点
                        let sib = node.parentElement.nextElementSibling;
                        if (sib) return sib.innerText.trim();
                        // 尝试父元素下一节点
                        sib = node.parentElement.parentElement
                                  ?.nextElementSibling;
                        if (sib) return sib.innerText.trim();
                    }
                }
                return null;
            }""",
            field,
        )
        if val:
            return val.split("\n")[0].strip()
    except Exception:
        pass

    # 策略2：XPath 取包含标签的行的第二个 td / span
    xpath_patterns = [
        f"//td[normalize-space()='{field}']/following-sibling::td[1]",
        f"//span[normalize-space()='{field}']/following-sibling::span[1]",
        f"//*[normalize-space()='{field}']/../following-sibling::*[1]",
    ]
    for xpath in xpath_patterns:
        try:
            el = page.locator(f"xpath={xpath}").first
            if await el.count() > 0:
                txt = await el.inner_text()
                return txt.split("\n")[0].strip()
        except Exception:
            continue

    return None


async def _intercept_api(page) -> dict:
    """拦截页面发出的 JSON 响应，尝试从中提取字段."""
    captured: list[dict] = []

    async def on_response(resp):
        try:
            ct = resp.headers.get("content-type", "")
            if "json" in ct and resp.status == 200:
                data = await resp.json()
                captured.append(data)
        except Exception:
            pass

    page.on("response", on_response)
    await asyncio.sleep(3)
    page.remove_listener("response", on_response)

    FIELD_KEYS = {
        "船长": ["length", "shipLength", "chuanChang", "loa"],
        "船宽": ["width", "beam", "shipWidth", "chuanKuan"],
        "船型": ["type", "shipType", "vesselType", "chuanXing"],
        "吃水": ["draft", "draught", "chiShui"],
    }

    result = {}
    for payload in captured:
        flat = _flatten(payload)
        for field, keys in FIELD_KEYS.items():
            if field in result:
                continue
            for k, v in flat.items():
                if any(key.lower() in k.lower() for key in keys) and v:
                    result[field] = str(v)
                    break
    return result


def _flatten(obj, prefix="", sep=".") -> dict:
    items = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            items.update(_flatten(v, f"{prefix}{sep}{k}" if prefix else k, sep))
    elif isinstance(obj, list):
        for i, v in enumerate(obj[:5]):
            items.update(_flatten(v, f"{prefix}[{i}]", sep))
    else:
        items[prefix] = obj
    return items


# ---------------------------------------------------------------------------
# 搜索 + 提取一条记录
# ---------------------------------------------------------------------------

async def scrape_one(page, ship_name: str) -> dict:
    record = {
        "ship_name": ship_name,
        "船长": None,
        "船宽": None,
        "船型": None,
        "吃水": None,
        "IMO": None,
        "status": "pending",
    }

    try:
        # ---- 打开首页 ----
        await page.goto(SITE, wait_until="domcontentloaded", timeout=30_000)
        await asyncio.sleep(1)

        # ---- 找搜索框 ----
        input_sel_candidates = [
            "input[placeholder*='船']",
            "input[placeholder*='ship' i]",
            "input[placeholder*='vessel' i]",
            "input[placeholder*='IMO']",
            ".search-input",
            "header input",
            "nav input",
            "input[type='text']",
        ]
        search_box = None
        for sel in input_sel_candidates:
            try:
                el = page.locator(sel).first
                if await el.count() > 0 and await el.is_visible():
                    search_box = el
                    break
            except Exception:
                continue

        if search_box is None:
            log.warning(f"[{ship_name}] 未找到搜索框")
            record["status"] = "no_search_box"
            return record

        # ---- 输入船名并搜索 ----
        await search_box.click()
        await search_box.triple_click()
        await search_box.fill(ship_name.upper())
        await asyncio.sleep(0.5)

        # 尝试点击搜索按钮
        search_btn_candidates = [
            "button[type='submit']",
            "button.search-btn",
            ".search-icon",
            ".icon-search",
            "button:has-text('搜索')",
            "button:has-text('查询')",
        ]
        clicked_btn = False
        for sel in search_btn_candidates:
            try:
                btn = page.locator(sel).first
                if await btn.count() > 0 and await btn.is_visible():
                    await btn.click()
                    clicked_btn = True
                    break
            except Exception:
                continue
        if not clicked_btn:
            await search_box.press("Enter")

        # ---- 等待结果 ----
        await asyncio.sleep(2)

        # 若出现下拉建议列表，点击第一条
        dropdown_candidates = [
            ".autocomplete-item",
            ".suggestion-item",
            ".search-dropdown li",
            ".dropdown-menu li",
            f".result-item:has-text('{ship_name.upper()}')",
            f"li:has-text('{ship_name.upper()}')",
        ]
        for sel in dropdown_candidates:
            try:
                el = page.locator(sel).first
                if await el.count() > 0 and await el.is_visible():
                    await el.click()
                    await asyncio.sleep(2)
                    break
            except Exception:
                continue

        # ---- 拦截 API 数据 ----
        api_data = await _intercept_api(page)

        # ---- 等待信息卡片 ----
        found_card = await _wait_for_card(page)
        if not found_card:
            log.warning(f"[{ship_name}] 未检测到信息卡片")

        # ---- 提取字段 ----
        for field in TARGET_FIELDS:
            # 优先使用 API 数据
            if api_data.get(field):
                record[field] = api_data[field]
                continue
            val = await _extract_by_label(page, field)
            if val:
                record[field] = val

        # 额外提取 IMO（用于核对）
        imo_val = await _extract_by_label(page, "IMO")
        if not imo_val:
            imo_val = await _extract_by_label(page, "imo")
        record["IMO"] = imo_val

        filled = sum(1 for f in TARGET_FIELDS if record[f])
        record["status"] = "ok" if filled > 0 else "no_data"

    except PWTimeout:
        record["status"] = "timeout"
        log.error(f"[{ship_name}] 超时")
    except Exception as exc:
        record["status"] = f"error:{exc!s:.80}"
        log.error(f"[{ship_name}] 异常: {exc}")

    return record


# ---------------------------------------------------------------------------
# 批量运行
# ---------------------------------------------------------------------------

async def run(ship_names: list[str], output: str) -> pd.DataFrame:
    log.info(f"共 {len(ship_names)} 艘船，输出文件: {output}")

    results = []
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"],
        )
        ctx = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="zh-CN",
        )
        page = await ctx.new_page()

        for idx, name in enumerate(ship_names, 1):
            name = str(name).strip()
            if not name:
                continue
            log.info(f"[{idx}/{len(ship_names)}] {name}")
            rec = await scrape_one(page, name)
            results.append(rec)
            log.info(
                f"  船长={rec['船长']}  船宽={rec['船宽']}  "
                f"船型={rec['船型']}  吃水={rec['吃水']}  [{rec['status']}]"
            )
            if idx < len(ship_names):
                await asyncio.sleep(REQUEST_DELAY)

        await browser.close()

    df = pd.DataFrame(results)
    df.to_excel(output, index=False)
    log.info(f"已保存 {len(df)} 条记录 → {output}")

    ok = (df["status"] == "ok").sum()
    log.info(f"成功率: {ok}/{len(df)} ({ok/max(len(df),1)*100:.1f}%)")
    return df


# ---------------------------------------------------------------------------
# CLI 入口
# ---------------------------------------------------------------------------

def _detect_ship_col(df: pd.DataFrame, hint: str | None) -> str:
    if hint and hint in df.columns:
        return hint
    for col in df.columns:
        up = str(col).upper()
        if any(kw in up for kw in ["SHIP", "VESSEL", "NAME", "船名", "英文", "VN"]):
            return col
    return df.columns[0]


def main():
    parser = argparse.ArgumentParser(description="维运网船舶信息批量爬虫")
    parser.add_argument("input", nargs="?", default="input.xlsx", help="输入 Excel/CSV 文件")
    parser.add_argument("-c", "--column", default=None, help="船名所在列名")
    parser.add_argument("-o", "--output", default="ship_data_output.xlsx", help="输出 Excel 文件")
    parser.add_argument("-s", "--ships", nargs="+", help="直接指定船名（跳过输入文件）")
    parser.add_argument("--gen-sample", action="store_true", help="生成示例输入文件后退出")
    args = parser.parse_args()

    if args.gen_sample:
        sample = pd.DataFrame({"船名(英文)": ["KOWLOON", "EVER GIVEN", "MSC OSCAR"]})
        sample.to_excel("input.xlsx", index=False)
        print("已生成 input.xlsx，请填入实际船名后运行爬虫。")
        return

    if args.ships:
        ship_names = args.ships
    else:
        p = Path(args.input)
        if not p.exists():
            log.error(f"输入文件不存在: {p}  (可用 --gen-sample 生成示例)")
            return
        df_in = pd.read_excel(p) if p.suffix == ".xlsx" else pd.read_csv(p)
        col = _detect_ship_col(df_in, args.column)
        log.info(f"使用列 '{col}' 作为船名")
        ship_names = df_in[col].dropna().astype(str).unique().tolist()

    asyncio.run(run(ship_names, args.output))


if __name__ == "__main__":
    main()

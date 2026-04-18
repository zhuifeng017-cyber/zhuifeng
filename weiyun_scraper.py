#!/usr/bin/env python3
"""
维运网船舶信息爬虫 (DrissionPage 版)
驱动本机真实 Chrome，自动通过 JS/Cookie 防爬挑战
提取：船长、船宽、船型、吃水

用法:
    python weiyun_scraper.py input.xlsx              # 读取 Excel，自动识别船名列
    python weiyun_scraper.py input.xlsx -c 英文船名  # 指定列名
    python weiyun_scraper.py -s KOWLOON EVER_GIVEN  # 直接指定船名
    python weiyun_scraper.py --gen-sample            # 生成示例输入文件
    python weiyun_scraper.py --headless              # 无头模式（需 Xvfb）
"""

import argparse
import logging
import random
import time
from pathlib import Path

import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions
from DrissionPage.errors import ElementNotFoundError

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)

SITE = "https://www.weiyun001.com"
TARGET_FIELDS = ["船长", "船宽", "船型", "吃水"]

# 监听关键字：匹配 shipLocate 请求（含加密参数的详情 API）
LISTEN_KEYWORD = "shipLocate"

# API 字段映射（JSON key → 中文字段）
FIELD_MAP = {
    "船长": ["length", "shipLength", "loa", "chuanChang", "船长"],
    "船宽": ["width", "beam", "shipWidth", "chuanKuan", "船宽"],
    "船型": ["type", "shipType", "vesselType", "chuanXing", "船型"],
    "吃水": ["draft", "draught", "chiShui", "吃水"],
    "IMO":  ["imo", "imoNo", "imoNumber"],
}


# ---------------------------------------------------------------------------
# 辅助函数
# ---------------------------------------------------------------------------

def _jitter(lo: float = 1.5, hi: float = 3.5):
    """随机等待，模拟人工操作节奏."""
    time.sleep(random.uniform(lo, hi))


def _human_type(ele, text: str):
    """逐字符输入，规避批量填充检测."""
    ele.clear()
    for ch in text:
        ele.input(ch)
        time.sleep(random.uniform(0.04, 0.12))


def _flatten(obj, prefix: str = "") -> dict:
    """将嵌套 JSON 展平为 {dotted.key: value} 字典."""
    items: dict = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            items.update(_flatten(v, f"{prefix}.{k}" if prefix else k))
    elif isinstance(obj, list):
        for i, v in enumerate(obj[:20]):
            items.update(_flatten(v, f"{prefix}[{i}]"))
    else:
        items[prefix] = obj
    return items


def _parse_api_body(body) -> dict:
    """从 API JSON 响应中提取目标字段."""
    if not isinstance(body, (dict, list)):
        return {}
    flat = _flatten(body)
    result: dict = {}
    for field, keys in FIELD_MAP.items():
        for k, v in flat.items():
            if v and any(key.lower() in k.lower() for key in keys):
                result[field] = str(v)
                break
    return result


def _extract_from_dom(page: ChromiumPage, field: str) -> str | None:
    """
    DOM 兜底提取：两种策略
    1. TreeWalker 找到标签文字节点 → 取相邻 sibling 的文字
    2. XPath 取标签后的第一个兄弟元素
    """
    # 策略 1：TreeWalker
    try:
        val = page.run_js(
            """(field) => {
                const walker = document.createTreeWalker(
                    document.body, NodeFilter.SHOW_TEXT);
                let node;
                while ((node = walker.nextNode())) {
                    if (node.textContent.trim() === field) {
                        const sib = node.parentElement.nextElementSibling;
                        if (sib) return sib.innerText.trim();
                        const par = node.parentElement.parentElement;
                        if (par?.nextElementSibling)
                            return par.nextElementSibling.innerText.trim();
                    }
                }
                return null;
            }""",
            field,
        )
        if val:
            return str(val).split("\n")[0].strip()
    except Exception:
        pass

    # 策略 2：XPath
    for xpath in [
        f"xpath://td[normalize-space()='{field}']/following-sibling::td[1]",
        f"xpath://span[normalize-space()='{field}']/following-sibling::span[1]",
        f"xpath://*[normalize-space()='{field}']/../following-sibling::*[1]",
    ]:
        try:
            el = page.ele(xpath, timeout=2)
            if el:
                txt = el.text.strip().split("\n")[0].strip()
                if txt:
                    return txt
        except Exception:
            continue

    return None


# ---------------------------------------------------------------------------
# 单条船舶搜索
# ---------------------------------------------------------------------------

def scrape_one(page: ChromiumPage, ship_name: str) -> dict:
    record: dict = {
        "ship_name": ship_name,
        "船长": None, "船宽": None, "船型": None, "吃水": None,
        "IMO": None, "status": "pending",
    }

    try:
        # 开始监听含 shipLocate 的请求（包括加密 vn/imo 参数的详情 API）
        page.listen.start(LISTEN_KEYWORD, method="GET")

        page.get(SITE, timeout=30, retry=2)
        _jitter(1.2, 2.5)

        # ---- 找搜索框 ----
        input_candidates = [
            "tag:input@placeholder:船",
            "tag:input@placeholder:ship",
            "tag:input@placeholder:vessel",
            "tag:input@placeholder:IMO",
            ".search-input",
            "tag:input@type=text",
        ]
        search_box = None
        for sel in input_candidates:
            try:
                el = page.ele(sel, timeout=3)
                if el:
                    search_box = el
                    break
            except ElementNotFoundError:
                continue

        if not search_box:
            log.warning(f"[{ship_name}] 未找到搜索框")
            record["status"] = "no_search_box"
            page.listen.stop()
            return record

        # ---- 人工输入船名 ----
        _human_type(search_box, ship_name.upper())
        _jitter(0.4, 1.0)

        # ---- 点击搜索按钮 / 回车 ----
        btn_candidates = [
            "tag:button@type=submit",
            ".search-btn", ".icon-search", ".search-icon",
            "tag:button@text():搜索",
            "tag:button@text():查询",
        ]
        clicked = False
        for sel in btn_candidates:
            try:
                btn = page.ele(sel, timeout=2)
                if btn:
                    btn.click()
                    clicked = True
                    break
            except ElementNotFoundError:
                continue
        if not clicked:
            search_box.input("\n")  # 回车兜底

        _jitter(2.0, 3.5)

        # ---- 若出现下拉建议，点击第一条 ----
        dropdown_candidates = [
            ".autocomplete-item", ".suggestion-item",
            ".search-dropdown li", ".dropdown-menu li",
            f"tag:li@@text():{ship_name.upper()}",
        ]
        for sel in dropdown_candidates:
            try:
                el = page.ele(sel, timeout=2)
                if el:
                    el.click()
                    _jitter(1.5, 2.5)
                    break
            except ElementNotFoundError:
                continue

        # ---- 等待 shipLocate API 响应（最多 15 秒）----
        # 网站通过加密参数请求此接口，DrissionPage listen 直接捕获响应体
        packet = page.listen.wait(timeout=15)
        page.listen.stop()

        if packet and packet.response:
            try:
                api_result = _parse_api_body(packet.response.body)
                if api_result:
                    record.update(api_result)
                    log.debug(f"[{ship_name}] API 命中: {api_result}")
            except Exception as e:
                log.debug(f"[{ship_name}] 解析 API 响应失败: {e}")

        # ---- DOM 兜底提取（API 未返回的字段）----
        for field in TARGET_FIELDS + ["IMO"]:
            if not record.get(field):
                val = _extract_from_dom(page, field)
                if val:
                    record[field] = val

        filled = sum(1 for f in TARGET_FIELDS if record.get(f))
        record["status"] = "ok" if filled > 0 else "no_data"

    except Exception as exc:
        record["status"] = f"error:{str(exc)[:80]}"
        log.error(f"[{ship_name}] 异常: {exc}")
        try:
            page.listen.stop()
        except Exception:
            pass

    return record


# ---------------------------------------------------------------------------
# 浏览器初始化
# ---------------------------------------------------------------------------

def _build_page(headless: bool = False) -> ChromiumPage:
    opt = ChromiumOptions()

    # 关键反检测参数
    opt.set_argument("--disable-blink-features=AutomationControlled")
    opt.set_argument("--no-sandbox")
    opt.set_argument("--disable-gpu")
    opt.set_argument("--disable-dev-shm-usage")

    # 使用本地持久化 Profile，复用真实 Cookie / localStorage
    # 首次运行后 Cookie 将保留，后续请求更像真实用户
    profile_dir = Path("./chrome_profile").resolve()
    profile_dir.mkdir(exist_ok=True)
    opt.set_user_data_path(str(profile_dir))

    opt.headless(headless)

    page = ChromiumPage(opt)

    # 隐藏 webdriver 特征
    page.run_js(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )

    return page


# ---------------------------------------------------------------------------
# 批量执行
# ---------------------------------------------------------------------------

def run(ship_names: list[str], output: str, headless: bool = False) -> pd.DataFrame:
    log.info(f"共 {len(ship_names)} 艘船 → {output}  (headless={headless})")
    page = _build_page(headless)
    results: list[dict] = []

    try:
        for idx, name in enumerate(ship_names, 1):
            name = str(name).strip()
            if not name:
                continue

            log.info(f"[{idx}/{len(ship_names)}] {name}")
            rec = scrape_one(page, name)
            results.append(rec)
            log.info(
                f"  船长={rec['船长']}  船宽={rec['船宽']}  "
                f"船型={rec['船型']}  吃水={rec['吃水']}  [{rec['status']}]"
            )

            if idx < len(ship_names):
                _jitter(2.5, 5.0)  # 批量间隔，降低被限速风险
    finally:
        page.quit()

    df = pd.DataFrame(results)
    df.to_excel(output, index=False)

    ok = (df["status"] == "ok").sum()
    log.info(
        f"完成。成功率 {ok}/{len(df)} "
        f"({ok / max(len(df), 1) * 100:.1f}%)，已保存 → {output}"
    )
    return df


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _detect_ship_col(df: pd.DataFrame, hint: str | None) -> str:
    if hint and hint in df.columns:
        return hint
    for col in df.columns:
        if any(kw in str(col).upper() for kw in ["SHIP", "VESSEL", "NAME", "船名", "英文", "VN"]):
            return col
    return df.columns[0]


def main():
    parser = argparse.ArgumentParser(description="维运网船舶信息批量爬虫 (DrissionPage)")
    parser.add_argument("input", nargs="?", default="input.xlsx", help="输入 Excel/CSV")
    parser.add_argument("-c", "--column", default=None, help="船名列名")
    parser.add_argument("-o", "--output", default="ship_data_output.xlsx", help="输出 Excel")
    parser.add_argument("-s", "--ships", nargs="+", help="直接指定船名")
    parser.add_argument("--headless", action="store_true", help="无头模式（需配置 Xvfb）")
    parser.add_argument("--gen-sample", action="store_true", help="生成示例 input.xlsx")
    args = parser.parse_args()

    if args.gen_sample:
        pd.DataFrame({"船名(英文)": ["KOWLOON", "EVER GIVEN", "MSC OSCAR"]}).to_excel(
            "input.xlsx", index=False
        )
        print("已生成 input.xlsx，请填入实际船名后重新运行。")
        return

    if args.ships:
        ship_names = args.ships
    else:
        p = Path(args.input)
        if not p.exists():
            log.error(f"文件不存在: {p}  (可用 --gen-sample 生成示例)")
            return
        df_in = pd.read_excel(p) if p.suffix == ".xlsx" else pd.read_csv(p)
        col = _detect_ship_col(df_in, args.column)
        log.info(f"使用列 '{col}' 作为船名")
        ship_names = df_in[col].dropna().astype(str).unique().tolist()

    run(ship_names, args.output, headless=args.headless)


if __name__ == "__main__":
    main()

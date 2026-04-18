#!/usr/bin/env python3
"""
维运网船舶信息爬虫 (DrissionPage 版)
提取：船长、船宽、船型、吃水

两种启动方式（推荐方式一）：

【方式一】接管已登录的 Chrome（最稳定）
  1. 用调试端口启动 Chrome（只需做一次，或加到 Chrome 快捷方式）：
       chrome.exe --remote-debugging-port=9222 --user-data-dir=C:/chrome_debug
  2. 在打开的 Chrome 里登录 weiyun001.com
  3. 运行脚本：
       python weiyun_scraper.py -s RABAUL CHIEF --port 9222

【方式二】脚本自己开 Chrome，自动检测登录状态后暂停让你手动登录
       python weiyun_scraper.py -s RABAUL CHIEF
       # 浏览器打开后，手动登录，登录完按终端回车继续

其他用法:
    python weiyun_scraper.py input.xlsx -c 英文船名  # 从 Excel 批量跑
    python weiyun_scraper.py --gen-sample            # 生成示例输入
"""

import argparse
import logging
import random
import time
from pathlib import Path

import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)

SITE = "https://www.weiyun001.com"
TARGET_FIELDS = ["船长", "船宽", "船型", "吃水"]
LISTEN_KEYWORD = "shipLocate"

FIELD_MAP = {
    "船长": ["length", "shipLength", "loa", "chuanChang", "船长"],
    "船宽": ["width", "beam", "shipWidth", "chuanKuan", "船宽"],
    "船型": ["type", "shipType", "vesselType", "chuanXing", "船型"],
    "吃水": ["draft", "draught", "chiShui", "吃水"],
    "IMO":  ["imo", "imoNo", "imoNumber"],
}


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def _jitter(lo: float = 1.5, hi: float = 3.5):
    time.sleep(random.uniform(lo, hi))


def _human_type(ele, text: str):
    ele.clear()
    for ch in text:
        ele.input(ch)
        time.sleep(random.uniform(0.05, 0.13))


def _flatten(obj, prefix: str = "") -> dict:
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
    # TreeWalker
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
# 登录检测
# ---------------------------------------------------------------------------

def _is_logged_in(page: ChromiumPage) -> bool:
    """检测当前页面是否已登录（找登出/用户名元素）."""
    login_indicators = [
        "tag:a@text():退出", "tag:a@text():登出", "tag:a@text():注销",
        ".user-info", ".user-avatar", ".logout", ".sign-out",
        "tag:span@text():退出登录",
    ]
    for sel in login_indicators:
        try:
            el = page.ele(sel, timeout=2)
            if el:
                return True
        except Exception:
            continue

    # 反向判断：如果找到登录按钮，说明未登录
    not_logged_in = [
        "tag:a@text():登录", "tag:a@text():注册", "tag:button@text():登录",
        ".login-btn", "#loginBtn",
    ]
    for sel in not_logged_in:
        try:
            el = page.ele(sel, timeout=2)
            if el:
                return False
        except Exception:
            continue

    # 无法判断，乐观认为已登录
    return True


def _ensure_logged_in(page: ChromiumPage):
    """如未登录则暂停，等待用户手动在浏览器里完成登录."""
    page.get(SITE, timeout=30)
    _jitter(1.5, 2.5)

    if not _is_logged_in(page):
        log.warning("=" * 60)
        log.warning("检测到未登录！请在已打开的浏览器窗口中手动登录维运网。")
        log.warning("登录完成后，回到此终端按 Enter 键继续...")
        log.warning("=" * 60)
        input()
        # 登录后刷新一次确认
        page.refresh()
        _jitter(1.0, 2.0)
        if not _is_logged_in(page):
            log.warning("仍未检测到登录状态，将继续尝试（如已登录可忽略此警告）。")
    else:
        log.info("已检测到登录状态，继续执行。")


# ---------------------------------------------------------------------------
# 单条船舶抓取
# ---------------------------------------------------------------------------

def scrape_one(page: ChromiumPage, ship_name: str) -> dict:
    record: dict = {
        "ship_name": ship_name,
        "船长": None, "船宽": None, "船型": None, "吃水": None,
        "IMO": None, "status": "pending",
    }

    try:
        page.listen.start(LISTEN_KEYWORD, method="GET")

        # 回到首页搜索（不重新登录）
        page.get(SITE, timeout=30)
        _jitter(1.0, 2.0)

        # 找搜索框
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
            except Exception:
                continue

        if not search_box:
            log.warning(f"[{ship_name}] 未找到搜索框，保存截图 debug_{ship_name}.png")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_search_box"
            page.listen.stop()
            return record

        # 人工输入
        _human_type(search_box, ship_name.upper())
        _jitter(0.5, 1.2)

        # 搜索
        btn_candidates = [
            "tag:button@type=submit",
            ".search-btn", ".icon-search", ".search-icon",
            "tag:button@text():搜索", "tag:button@text():查询",
        ]
        clicked = False
        for sel in btn_candidates:
            try:
                btn = page.ele(sel, timeout=2)
                if btn:
                    btn.click()
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            search_box.input("\n")

        _jitter(2.0, 3.5)

        # 下拉建议，点第一条
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
            except Exception:
                continue

        # 等待 shipLocate API（加密参数由浏览器自己处理）
        log.info(f"  等待 shipLocate API 响应（最多 15s）...")
        packet = page.listen.wait(timeout=15)
        page.listen.stop()

        if packet and packet.response:
            try:
                api_result = _parse_api_body(packet.response.body)
                if api_result:
                    record.update(api_result)
                    log.debug(f"  API 命中: {api_result}")
            except Exception as e:
                log.debug(f"  解析 API 响应失败: {e}")
        else:
            log.warning(f"  未捕获到 API 响应，尝试 DOM 提取...")
            # 超时时截图，便于排查
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)

        # DOM 兜底
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

def _build_page(port: int | None = None) -> ChromiumPage:
    opt = ChromiumOptions()
    opt.set_argument("--disable-blink-features=AutomationControlled")
    opt.set_argument("--no-sandbox")
    opt.set_argument("--disable-gpu")
    opt.set_argument("--disable-dev-shm-usage")

    if port:
        # 接管已开启调试端口的 Chrome（最稳定，直接用已登录 session）
        log.info(f"连接到已有 Chrome（端口 {port}）...")
        opt.set_address(f"127.0.0.1:{port}")
    else:
        # 自己开 Chrome，持久化 profile
        profile_dir = Path("./chrome_profile").resolve()
        profile_dir.mkdir(exist_ok=True)
        opt.set_user_data_path(str(profile_dir))
        opt.headless(False)

    page = ChromiumPage(addr_or_opts=opt)
    # 隐藏 webdriver 标记
    page.run_js(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return page


# ---------------------------------------------------------------------------
# 批量执行
# ---------------------------------------------------------------------------

def run(ship_names: list[str], output: str, port: int | None = None) -> pd.DataFrame:
    log.info(f"共 {len(ship_names)} 艘船 → {output}")
    page = _build_page(port)
    results: list[dict] = []

    try:
        # 登录检测（仅在自己开 Chrome 时检测，接管模式假定已登录）
        if not port:
            _ensure_logged_in(page)

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

            # 增量保存，防止中途崩溃丢数据
            pd.DataFrame(results).to_excel(output, index=False)

            if idx < len(ship_names):
                _jitter(2.5, 5.0)
    finally:
        if not port:
            page.quit()
        # 有 port 时不关闭，用户的 Chrome 继续开着

    df = pd.DataFrame(results)
    ok = (df["status"] == "ok").sum()
    log.info(f"完成。成功率 {ok}/{len(df)} ({ok/max(len(df),1)*100:.1f}%) → {output}")
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
    parser = argparse.ArgumentParser(description="维运网船舶信息批量爬虫")
    parser.add_argument("input", nargs="?", default="input.xlsx")
    parser.add_argument("-c", "--column", default=None, help="船名列名")
    parser.add_argument("-o", "--output", default="ship_data_output.xlsx")
    parser.add_argument("-s", "--ships", nargs="+", help="直接指定船名")
    parser.add_argument(
        "--port", type=int, default=None,
        help="接管已开启 --remote-debugging-port 的 Chrome，例如 --port 9222"
    )
    parser.add_argument("--gen-sample", action="store_true")
    args = parser.parse_args()

    if args.gen_sample:
        pd.DataFrame({"船名(英文)": ["KOWLOON", "EVER GIVEN", "MSC OSCAR"]}).to_excel(
            "input.xlsx", index=False
        )
        print("已生成 input.xlsx")
        return

    if args.ships:
        ship_names = args.ships
    else:
        p = Path(args.input)
        if not p.exists():
            log.error(f"文件不存在: {p}")
            return
        df_in = pd.read_excel(p) if p.suffix == ".xlsx" else pd.read_csv(p)
        col = _detect_ship_col(df_in, args.column)
        log.info(f"使用列 '{col}' 作为船名")
        ship_names = df_in[col].dropna().astype(str).unique().tolist()

    run(ship_names, args.output, port=args.port)


if __name__ == "__main__":
    main()

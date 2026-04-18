#!/usr/bin/env python3
"""
维运网船舶信息爬虫
流程：搜索船名 → 点击结果 → 跳转 /shipLocate 页面 → 提取船长/船宽/船型/吃水

启动方式（推荐方式一）：
  1. 用调试端口打开 Chrome，登录 weiyun001.com：
       chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\chrome_debug
  2. 运行：
       python weiyun_scraper.py -s RABAUL CHIEF --port 9222

  方式二（脚本自己开 Chrome，自动暂停让你登录）：
       python weiyun_scraper.py -s RABAUL CHIEF

  从 Excel 批量跑：
       python weiyun_scraper.py input.xlsx -c 英文船名 --port 9222
"""

import argparse
import logging
import random
import time
from pathlib import Path

import pandas as pd
import urllib.request
from DrissionPage import ChromiumPage, ChromiumOptions

log = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)

SITE = "https://www.weiyun001.com"
TARGET_FIELDS = ["船长", "船宽", "船型", "吃水"]


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------

def _jitter(lo: float = 1.0, hi: float = 2.5):
    time.sleep(random.uniform(lo, hi))


def _human_type(ele, text: str):
    ele.clear()
    for ch in text:
        ele.input(ch)
        time.sleep(random.uniform(0.05, 0.13))


def _wait_for_url(page: ChromiumPage, keyword: str, timeout: float = 20.0) -> bool:
    """等待页面 URL 中出现指定关键字，返回是否成功。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if keyword in page.url:
            return True
        time.sleep(0.4)
    return False


# ---------------------------------------------------------------------------
# 从 shipLocate 页面提取字段
# ---------------------------------------------------------------------------

def _extract_fields(page: ChromiumPage) -> dict:
    """
    从已加载的 shipLocate 页面提取目标字段。
    策略1：JS TreeWalker 找标签文字，取其相邻节点值。
    策略2：XPath 找标签后的兄弟元素。
    """
    result: dict = {}

    for field in TARGET_FIELDS + ["IMO"]:
        # 策略 1：TreeWalker
        try:
            val = page.run_js(
                """(field) => {
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT);
                    let node;
                    while ((node = walker.nextNode())) {
                        const txt = node.textContent.trim();
                        if (txt === field || txt === field + ':' || txt === field + '：') {
                            // 尝试父元素的下一兄弟
                            const par = node.parentElement;
                            if (par.nextElementSibling)
                                return par.nextElementSibling.innerText.trim();
                            // 尝试祖父元素的下一兄弟
                            if (par.parentElement?.nextElementSibling)
                                return par.parentElement.nextElementSibling.innerText.trim();
                        }
                    }
                    return null;
                }""",
                field,
            )
            if val:
                result[field] = str(val).split("\n")[0].strip()
                continue
        except Exception:
            pass

        # 策略 2：XPath
        for xpath in [
            f"xpath://td[normalize-space(.)='{field}']/following-sibling::td[1]",
            f"xpath://span[normalize-space(.)='{field}']/following-sibling::span[1]",
            f"xpath://div[normalize-space(.)='{field}']/following-sibling::div[1]",
            f"xpath://*[normalize-space(.)='{field}']/../following-sibling::*[1]",
        ]:
            try:
                el = page.ele(xpath, timeout=1)
                if el:
                    txt = el.text.strip().split("\n")[0].strip()
                    if txt:
                        result[field] = txt
                        break
            except Exception:
                continue

    return result


# ---------------------------------------------------------------------------
# 单条抓取（核心流程）
# ---------------------------------------------------------------------------

def scrape_one(page: ChromiumPage, ship_name: str) -> dict:
    record: dict = {
        "ship_name": ship_name,
        "船长": None, "船宽": None, "船型": None, "吃水": None,
        "IMO": None, "status": "pending",
    }

    try:
        # ── 1. 回到首页 ──────────────────────────────────────────────
        page.get(SITE, timeout=30)
        _jitter(1.2, 2.2)

        # ── 2. 找搜索框 ───────────────────────────────────────────────
        for sel in [
            "tag:input@placeholder:船",
            "tag:input@placeholder:ship",
            "tag:input@placeholder:vessel",
            "tag:input@placeholder:IMO",
            ".search-input",
            "tag:input@type=text",
        ]:
            try:
                el = page.ele(sel, timeout=3)
                if el:
                    search_box = el
                    break
            except Exception:
                continue
        else:
            log.warning(f"[{ship_name}] 未找到搜索框 → 截图 debug_{ship_name}.png")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_search_box"
            return record

        # ── 3. 输入船名 ───────────────────────────────────────────────
        _human_type(search_box, ship_name.upper())
        _jitter(1.0, 2.0)   # 等待自动补全出现

        # ── 4. 触发自动补全并选择第一条 ──────────────────────────────
        # 截图：看输入后页面上出现了什么
        page.get_screenshot(path=f"debug_{ship_name}_after_type.png")

        # 策略A：用键盘 ↓ 键选中补全第一项，再回车确认
        # 适用于大多数基于 input 的自动补全组件
        try:
            page.actions.key_down("ArrowDown").key_up("ArrowDown")
            time.sleep(0.4)
            page.actions.key_down("Return").key_up("Return")
            log.info("  已用键盘 ↓+Enter 选中补全项")
        except Exception as e:
            log.debug(f"  键盘导航失败: {e}")

        _jitter(1.5, 2.5)

        # 如果键盘没触发跳转，再尝试点击具体的补全条目
        if "shipLocate" not in page.url:
            clicked = False
            for sel in [
                ".autocomplete-item",
                ".suggestion-item",
                ".search-result-item",
                ".search-dropdown li",
                ".dropdown-menu li",
                f"tag:li@@text():{ship_name.upper()}",
            ]:
                try:
                    el = page.ele(sel, timeout=2)
                    if el and el.text.strip():
                        el.click()
                        clicked = True
                        log.info(f"  点击补全项: {el.text[:40]!r}")
                        break
                except Exception:
                    continue

            if not clicked:
                log.info("  未找到补全项，回车提交搜索...")
                search_box.input("\n")
                _jitter(2.0, 3.0)
                # 搜索结果页点第一条
                for sel in [
                    ".search-result-item", ".result-item",
                    ".ship-item", ".vessel-item", ".list-item",
                ]:
                    try:
                        el = page.ele(sel, timeout=4)
                        if el and el.text.strip():
                            el.click()
                            log.info(f"  点击搜索结果: {el.text[:40]!r}")
                            break
                    except Exception:
                        continue

        # ── 5. 等待跳转到 shipLocate 页面 ────────────────────────────
        log.info(f"  等待跳转 shipLocate 页面（最多 20s）...")
        arrived = _wait_for_url(page, "shipLocate", timeout=20)

        if not arrived:
            log.warning(f"  未到达 shipLocate，当前 URL: {page.url}")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_navigate"
            return record

        log.info(f"  已到达: {page.url}")
        page.wait.load()        # 等待页面完全渲染
        _jitter(1.0, 2.0)

        # ── 6. 提取字段 ───────────────────────────────────────────────
        extracted = _extract_fields(page)
        record.update(extracted)

        filled = sum(1 for f in TARGET_FIELDS if record.get(f))
        if filled == 0:
            log.warning(f"  字段均为空 → 截图 debug_{ship_name}.png")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_data"
        else:
            record["status"] = "ok"

    except Exception as exc:
        record["status"] = f"error:{str(exc)[:80]}"
        log.error(f"[{ship_name}] 异常: {exc}")

    return record


# ---------------------------------------------------------------------------
# 登录检测
# ---------------------------------------------------------------------------

def _ensure_logged_in(page: ChromiumPage):
    page.get(SITE, timeout=30)
    _jitter(1.5, 2.5)

    not_logged_in_signs = [
        "tag:a@text():登录", "tag:a@text():注册",
        "tag:button@text():登录", ".login-btn", "#loginBtn",
    ]
    needs_login = False
    for sel in not_logged_in_signs:
        try:
            if page.ele(sel, timeout=2):
                needs_login = True
                break
        except Exception:
            continue

    if needs_login:
        log.warning("=" * 60)
        log.warning("检测到未登录！请在已打开的浏览器窗口中手动登录维运网。")
        log.warning("登录完成后，回到此终端按 Enter 键继续...")
        log.warning("=" * 60)
        input()
        page.refresh()
        _jitter(1.0, 2.0)
    else:
        log.info("已登录，开始抓取。")


# ---------------------------------------------------------------------------
# 浏览器初始化
# ---------------------------------------------------------------------------

def _find_chrome_port() -> int | None:
    """自动扫描 9222-9230，找到正在运行的 Chrome 调试端口。"""
    for p in range(9222, 9231):
        try:
            with urllib.request.urlopen(
                f"http://127.0.0.1:{p}/json/version", timeout=1
            ) as r:
                if r.status == 200:
                    log.info(f"自动检测到 Chrome 调试端口: {p}")
                    return p
        except Exception:
            continue
    return None


def _build_page(port: int | None = None) -> ChromiumPage:
    opt = ChromiumOptions()
    opt.set_argument("--disable-blink-features=AutomationControlled")
    opt.set_argument("--no-sandbox")
    opt.set_argument("--disable-gpu")
    opt.set_argument("--disable-dev-shm-usage")

    if port:
        log.info(f"接管已有 Chrome（端口 {port}）...")
        opt.set_address(f"127.0.0.1:{port}")
    else:
        profile_dir = Path("./chrome_profile").resolve()
        profile_dir.mkdir(exist_ok=True)
        opt.set_user_data_path(str(profile_dir))
        opt.headless(False)

    page = ChromiumPage(addr_or_opts=opt)
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
            # 每条都增量保存
            pd.DataFrame(results).to_excel(output, index=False)

            if idx < len(ship_names):
                _jitter(3.0, 5.0)
    finally:
        if not port:
            page.quit()

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
    parser.add_argument("--port", default=None,
                        help="接管已登录 Chrome 的调试端口，例如 --port 9222；"
                             "填 auto 则自动扫描")
    parser.add_argument("--gen-sample", action="store_true")
    args = parser.parse_args()

    if args.gen_sample:
        pd.DataFrame({"船名(英文)": ["KOWLOON", "EVER GIVEN", "MSC OSCAR"]}).to_excel(
            "input.xlsx", index=False)
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

    port = None
    if args.port:
        if str(args.port).lower() == "auto":
            port = _find_chrome_port()
            if port is None:
                log.error(
                    "未找到 Chrome 调试端口（9222-9230）。\n"
                    "请用以下命令重新打开 Chrome：\n"
                    "  chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\\chrome_debug\n"
                    "然后在 Chrome 里登录网站，再运行脚本。"
                )
                return
        else:
            port = int(args.port)

    run(ship_names, args.output, port=port)


if __name__ == "__main__":
    main()

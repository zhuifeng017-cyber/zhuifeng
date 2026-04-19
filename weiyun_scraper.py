#!/usr/bin/env python3
"""
维运网船舶信息爬虫
流程：搜索船名 → 点击结果 → 跳转 /shipLocate 页面 → 提取船长/船宽/船型/吃水

启动方式（推荐方式一）：
  1. 用调试端口打开 Chrome，登录 weiyun001.com：
       chrome.exe --remote-debugging-port=9222 --user-data-dir=C:/chrome_debug
  2. 运行：
       python weiyun_scraper.py -s RABAUL CHIEF --port 9222

  方式二（脚本自己开 Chrome，自动暂停让你登录）：
       python weiyun_scraper.py -s RABAUL CHIEF

  从 Excel 批量跑：
       python weiyun_scraper.py input.xlsx -c 英文船名 --port 9222
"""

import argparse
import logging
import os
import random
import subprocess
import sys
import time
import urllib.request
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


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------

def _jitter(lo: float = 1.0, hi: float = 2.5):
    time.sleep(random.uniform(lo, hi))


def _set_input(page: ChromiumPage, ele, text: str):
    """
    向 Ant Design / Vue / React 受控 input 写入值。
    用 ele.run_js()（this = 元素本身）调用 native setter，
    再 dispatch input/change 事件触发框架响应式更新。
    """
    ele.click()
    time.sleep(0.3)
    # ele.run_js 里 this 就是该元素，比 page.run_js(script, ele) 更稳定
    ele.run_js(f"""
        const setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(this, {repr(text)});
        this.dispatchEvent(new Event('input',  {{bubbles: true}}));
        this.dispatchEvent(new Event('change', {{bubbles: true}}));
    """)
    time.sleep(0.4)


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
        # ── 1. 直接打开船舶定位搜索页 ────────────────────────────────
        # 点击首页「船舶定位」标签会跳转到此页面，直接导航更可靠
        page.get(f"{SITE}/shipLocate", timeout=30)
        _jitter(1.5, 2.5)

        # 若被重定向回首页（未登录），截图并退出
        if "shipLocate" not in page.url:
            log.warning(f"[{ship_name}] 跳转船舶定位页失败，当前 URL: {page.url} → 截图")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_navigate"
            return record

        log.info(f"  已到达船舶定位页: {page.url}")

        # ── 2. 关闭可能存在的弹窗 ────────────────────────────────────
        for btn_sel in [
            "tag:button@@text():确认",
            "tag:button@@text():确定",
            "tag:button@@text():关闭",
        ]:
            try:
                btn = page.ele(btn_sel, timeout=1)
                if btn:
                    btn.click()
                    _jitter(0.3, 0.8)
                    break
            except Exception:
                pass

        # ── 3. 找输入框并填入船名 ─────────────────────────────────────
        _jitter(1.5, 2.5)

        # 先列出页面上所有输入框，帮助定位正确元素
        all_inputs = page.run_js("""
            return JSON.stringify(Array.from(document.querySelectorAll('input')).map(el => ({
                ph:   el.placeholder,
                type: el.type,
                id:   el.id,
                cls:  el.className.slice(0, 60)
            })));
        """)
        log.info(f"  页面输入框列表: {all_inputs}")

        # 船舶定位页的搜索框是 Ant Design Input，class = ant-input
        # （通过日志确认 placeholder = '请输入搜索内容'，type=text，cls=ant-input）
        search_box = None
        for sel in [
            ".ant-input",                                          # Ant Design Input 组件
            "xpath://input[contains(@placeholder,'英文船名')]",
            "xpath://input[contains(@placeholder,'MMSI')]",
            "xpath://input[contains(@placeholder,'船名')]",
            "xpath://input[@type='text' and not(@id)]",           # 无 id 的文本框
        ]:
            try:
                el = page.ele(sel, timeout=3)
                if el:
                    search_box = el
                    log.info(f"  找到输入框: placeholder={el.attr('placeholder')!r} cls={el.attr('class')!r}")
                    break
            except Exception:
                continue

        if not search_box:
            log.warning(f"[{ship_name}] 未找到搜索输入框 → 截图")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_search_box"
            return record

        _set_input(page, search_box, ship_name.upper())

        # 确认输入框有值，若为空则退出
        cur_val = search_box.run_js("return this.value")
        log.info(f"  输入框当前值: {cur_val!r}")
        if not cur_val or not cur_val.strip():
            log.warning(f"[{ship_name}] 输入框仍为空，跳过搜索 → 截图")
            page.get_screenshot(path=f"debug_{ship_name}_empty.png", full_page=True)
            record["status"] = "input_failed"
            return record

        # ── 4. 等待自动补全下拉候选出现，点击第一条匹配结果 ────────────
        log.info("  等待搜索候选下拉...")
        _jitter(0.8, 1.5)

        dropdown_item = None
        for attempt in range(10):
            for sel in [
                "xpath://div[contains(@class,'ant-select-item') and not(contains(@class,'empty'))]",
                "xpath://li[contains(@class,'ant-select-dropdown-menu-item')]",
                ".ant-select-item-option",
                "xpath://div[contains(@class,'suggest')]//div[1]",
                "xpath://div[contains(@class,'result-item')]//div[1]",
            ]:
                try:
                    el = page.ele(sel, timeout=1)
                    if el and el.text.strip():
                        dropdown_item = el
                        log.info(f"  找到候选项: {el.text.strip()[:60]!r}")
                        break
                except Exception:
                    continue
            if dropdown_item:
                break
            _jitter(0.4, 0.8)

        if dropdown_item:
            dropdown_item.click()
            log.info("  已点击候选项")
        else:
            # 兜底：发送 Enter 键触发搜索
            log.warning("  未找到下拉候选，改用 Enter 键")
            search_box.run_js("""
                this.dispatchEvent(new KeyboardEvent('keydown',
                    {key:'Enter', keyCode:13, bubbles:true, cancelable:true}));
                this.dispatchEvent(new KeyboardEvent('keyup',
                    {key:'Enter', keyCode:13, bubbles:true, cancelable:true}));
            """)

        # ── 5. 等待 URL 更新出现 vn= 参数（页面本身就在 /shipLocate，
        #        需要等参数出现以确认已加载到具体船舶）────────────────────
        log.info("  等待船舶详情 URL（vn= 参数）...")
        arrived = _wait_for_url(page, "vn=", timeout=20)

        if not arrived:
            log.warning(f"  URL 未出现 vn= 参数，当前 URL: {page.url}")
            page.get_screenshot(path=f"debug_{ship_name}.png", full_page=True)
            record["status"] = "no_navigate"
            return record

        log.info(f"  已到达: {page.url}")
        _jitter(1.5, 2.5)

        # ── 6. 点击「展开」按钮，展开船舶详细参数 ────────────────────────
        log.info("  查找「展开」按钮...")
        expand_clicked = False
        for sel in [
            "xpath://span[normalize-space(text())='展开']",
            "xpath://a[normalize-space(text())='展开']",
            "xpath://div[normalize-space(text())='展开']",
            "xpath://button[contains(.,'展开')]",
            "tag:span@@text():展开",
            "tag:a@@text():展开",
        ]:
            try:
                btn = page.ele(sel, timeout=3)
                if btn:
                    btn.click()
                    log.info(f"  已点击「展开」按钮 (sel={sel!r})")
                    expand_clicked = True
                    _jitter(0.8, 1.5)
                    break
            except Exception:
                continue
        if not expand_clicked:
            log.warning("  未找到「展开」按钮，直接提取当前页面数据")

        # ── 7. 提取字段 ───────────────────────────────────────────────
        _jitter(1.0, 2.0)
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
    _jitter(2.0, 3.0)

    def _needs_login() -> str | None:
        """返回未登录原因字符串，None 表示已登录。"""
        url = page.url.lower()
        # 1. URL 被重定向（被踢到注册/登录/绑定页）
        for kw in ["bind", "login", "register", "signin", "auth", "regist"]:
            if kw in url:
                return f"URL 重定向: {page.url}"
        # 2. 页面标题含注册/登录字样
        try:
            title = page.title
            for kw in ["绑定", "注册", "登录", "Login", "Register"]:
                if kw in title:
                    return f"页面标题: {title}"
        except Exception:
            pass
        # 3. 页面正文出现「登录/注册」或「存续/注册」按钮
        for sel in [
            "xpath://*[contains(@class,'login') or contains(@class,'register') or contains(@class,'signin')]",
            "tag:a@@text():登录",
            "tag:a@@text():注册",
            "tag:span@@text():登录",
            "tag:button@@text():登录",
        ]:
            try:
                el = page.ele(sel, timeout=1)
                if el and el.text.strip():
                    return f"页面元素: {el.text.strip()[:30]!r}"
            except Exception:
                continue
        return None

    reason = _needs_login()
    if reason:
        # 若被重定向，先导回首页
        if "URL" in reason:
            page.get(SITE, timeout=30)
            _jitter(1.0, 1.5)

        log.warning("=" * 60)
        log.warning(f"检测到未登录（{reason}）")
        log.warning("请在浏览器窗口中完成登录，登录后回到终端按 Enter 继续...")
        log.warning("=" * 60)
        input()
        page.get(SITE, timeout=30)   # 登录后刷新回首页
        _jitter(1.5, 2.5)

        # 二次确认
        reason2 = _needs_login()
        if reason2:
            log.warning(f"仍未确认登录（{reason2}），继续尝试...")
        else:
            log.info("登录确认，开始抓取。")
    else:
        log.info("登录确认，开始抓取。")


# ---------------------------------------------------------------------------
# 浏览器初始化
# ---------------------------------------------------------------------------

# Windows 常见 Chrome 安装路径
_CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    r"C:\Program Files\Google\Chrome Beta\Application\chrome.exe",
    # macOS
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    # Linux
    "/usr/bin/google-chrome",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
]

DEBUG_PORT = 9222
DEBUG_DIR  = str(Path("./chrome_debug").resolve())


def _probe_port(port: int) -> bool:
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/json/version", timeout=1
        ) as r:
            return r.status == 200
    except Exception:
        return False


def _find_chrome_port() -> int | None:
    """扫描 9222-9230，找到已开启调试端口的 Chrome。"""
    for p in range(9222, 9231):
        if _probe_port(p):
            log.info(f"检测到 Chrome 调试端口: {p}")
            return p
    return None


def _find_chrome_exe() -> str | None:
    """在常见路径中寻找 Chrome 可执行文件。"""
    for path in _CHROME_PATHS:
        if path and os.path.exists(path):
            return path
    return None


def _launch_chrome(port: int = DEBUG_PORT) -> bool:
    """
    自动找到 Chrome 并以调试端口启动，打开维运网。
    返回是否成功启动并监听。
    """
    exe = _find_chrome_exe()
    if not exe:
        log.error(
            "未找到 Chrome。请手动打开 Chrome 并访问 weiyun001.com 登录，"
            "同时在 Chrome 快捷方式中加入参数：\n"
            f"  --remote-debugging-port={port} --user-data-dir={DEBUG_DIR}"
        )
        return False

    log.info(f"启动 Chrome: {exe}")
    Path(DEBUG_DIR).mkdir(exist_ok=True)
    subprocess.Popen([
        exe,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={DEBUG_DIR}",
        "--disable-blink-features=AutomationControlled",
        "--no-first-run",               # 不显示"欢迎使用 Chrome"
        "--no-default-browser-check",   # 不提示设为默认浏览器
        "--disable-sync",               # 禁用 Google 账号同步提示
        "--disable-extensions",         # 不加载扩展（加速启动）
        SITE,   # 直接打开维运网
    ])

    # 等待 Chrome 就绪（最多 15 秒）
    for i in range(15):
        time.sleep(1)
        if _probe_port(port):
            log.info("Chrome 已就绪。")
            return True
    log.error("Chrome 启动超时。")
    return False


def _build_page(port: int) -> ChromiumPage:
    opt = ChromiumOptions()
    opt.set_address(f"127.0.0.1:{port}")
    page = ChromiumPage(addr_or_opts=opt)
    page.run_js(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return page


# ---------------------------------------------------------------------------
# 批量执行
# ---------------------------------------------------------------------------

def run(ship_names: list[str], output: str, port: int) -> pd.DataFrame:
    log.info(f"共 {len(ship_names)} 艘船 → {output}")
    page = _build_page(port)
    results: list[dict] = []

    try:
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
            # 每条都增量保存（文件被占用时自动换名）
            _safe_save(pd.DataFrame(results), output)

            if idx < len(ship_names):
                _jitter(3.0, 5.0)
    finally:
        pass  # 保持 Chrome 开着，方便用户检查

    df = pd.DataFrame(results)
    saved = _safe_save(df, output)
    ok = (df["status"] == "ok").sum()
    log.info(f"完成。成功率 {ok}/{len(df)} ({ok/max(len(df),1)*100:.1f}%) → {saved}")
    return df


def _safe_save(df: pd.DataFrame, path: str) -> str:
    """保存 Excel，文件被占用时自动加时间戳换名。"""
    try:
        df.to_excel(path, index=False)
        return path
    except PermissionError:
        ts = time.strftime("%H%M%S")
        alt = path.replace(".xlsx", f"_{ts}.xlsx")
        df.to_excel(alt, index=False)
        log.warning(f"文件被占用，已另存为 {alt}（请关闭 Excel 后重试）")
        return alt


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

    # 确定调试端口
    if args.port and str(args.port).lower() != "auto":
        port = int(args.port)
    else:
        # 先扫描已有 Chrome 调试实例
        port = _find_chrome_port()

    if port is None:
        # 没有找到，自动启动 Chrome
        log.info("未检测到 Chrome 调试实例，尝试自动启动...")
        if not _launch_chrome(DEBUG_PORT):
            return
        port = DEBUG_PORT
        # 提示用户登录
        log.warning("=" * 60)
        log.warning("Chrome 已打开并访问维运网。")
        log.warning("请在浏览器中完成登录，登录后回到此终端按 Enter 继续...")
        log.warning("=" * 60)
        input()

    run(ship_names, args.output, port=port)


if __name__ == "__main__":
    main()

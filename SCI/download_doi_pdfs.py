# -*- coding: utf-8 -*-
"""按 DOI 下载九类期刊题录对应的开放获取全文 PDF。

参考 工作流程/爬虫/sci.py：从题录表读 DOI，请求页面、解析 PDF 链接并保存。
只读取 D:\\2026年工作\\SCI数据库下载\\九类期刊题录 下的原表，不修改、不覆盖。
PDF 以 DOI 命名；成败写入独立日志，不写回原 Excel。

获取来源为 DOI 官方跳转与 Unpaywall 登记的开放获取地址（出版社 OA / PMC 等），
不使用未授权镜像站。

示例：
  python download_doi_pdfs.py --dry-run
  python download_doi_pdfs.py --domain 集成电路 --limit 5
  python download_doi_pdfs.py --domain 集成电路 --email your@email.com
  python download_doi_pdfs.py --resume
  python download_doi_pdfs.py --browser   # 可选：独立 Chrome 过 Cloudflare
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import warnings

try:
    import undetected_chromedriver as uc
except ImportError:
    uc = None

try:
    import winreg
except ImportError:
    winreg = None

warnings.filterwarnings("ignore", category=Warning, module="requests")

RECORDS_ROOT = Path(r"D:\2026年工作\SCI数据库下载\九类期刊题录")
OUTPUT_ROOT = Path(r"D:\2026年工作\SCI数据库下载\九类期刊全文")
LOG_NAME = "_doi_download_log.csv"
_SCRIPT_DIR = Path(__file__).resolve().parent
# 独立 Chrome 用户目录 + 调试端口，与 sci_journals(9223) / CNKI(9222) 隔离
CHROME_PROFILE_ROOT = str(_SCRIPT_DIR / ".chrome_profiles" / "doi_pdfs")
CHROME_DEBUG_PORT_BASE = 9322
CF_WAIT_TIMEOUT = 45
BROWSER_LAUNCH_TIMEOUT = 25
# 批量下载默认不用浏览器，避免与已打开的 WoS/日常 Chrome 抢锁
USE_BROWSER_CF = False
_BROWSER_DISABLED = False
_DEBUG_PORT_SEQ = 0

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36 "
        "mailto:research@example.edu"
    ),
    "Accept": "*/*",
}

CLOUDFLARE_MARKERS = (
    b"just a moment",
    b"cf-browser-verification",
    b"cf-turnstile",
    b"check-your-browser",
    b"challenge-platform",
    b"checking your browser",
    b"pow_challenge",
    b"preparing to download",
    b"cloudpmc-viewer-pow",
    b'content="noindex, nofollow"',
    b"content='noindex, nofollow'",
    b"making sure you",
    b"you're not a bot",
    b"you&#39;re not a bot",
    b"/.within.website/",
)

HCAPTCHA_MARKERS = (
    b"hcaptcha.com",
    b"h-captcha",
    b"data-hcaptcha-widget-id",
    b"h-captcha-response",
    b"newassets.hcaptcha.com",
)

# 页面源码（str）用的标记
CF_MARKERS_STR = (
    "cf-turnstile",
    "check-your-browser",
    "just a moment",
    "cf-browser-verification",
    "challenge-platform",
    "checking your browser",
    "enable javascript and cookies",
)

HCAPTCHA_MARKERS_STR = (
    "hcaptcha.com",
    "h-captcha",
    "data-hcaptcha-widget-id",
    "h-captcha-response",
    "newassets.hcaptcha.com",
)

# 这些站点常有 hCaptcha / 强反爬；优先尝试其他仓储，失败后再用浏览器过验证重试
BLOCKED_HOST_HINTS = (
    "pubs.acs.org",
    "spiedigitallibrary.org",
    "ieeexplore.ieee.org",
    "opg.optica.org",
    "hal.science",
    "hal.archives-ouvertes.fr",
)

# 仅对这些站点在 --browser 开启时用浏览器过 CF（不含出版社付费站、不含 OpenAlex 作品页）
CLOUDFLARE_HOST_HINTS = (
    "doi.org",
    "dx.doi.org",
    "content.openalex.org",
)

DOI_COL_CANDIDATES = ("DOI", "doi", "Digital Object Identifier", "DI")
TITLE_COL_CANDIDATES = ("Article Title", "Title", "Article Title - English")
DOCTYPE_COL_CANDIDATES = ("Document Type", "DT")
EXCEL_SUFFIXES = {".xls", ".xlsx"}
ILLEGAL_FS = re.compile(r'[<>:"/\\|?*]+')
DOI_RE = re.compile(r"^10\.\d{4,9}/\S+$", re.I)
# ACS 期次级 DOI，例如 10.1021/phv011i004_1792243
ISSUE_DOI_RE = re.compile(r"^10\.1021/phv", re.I)
# SPIE Advanced Photonics 封面/补白：每期 019901–069901
SPIE_FRONT_MATTER_RE = re.compile(r"^10\.1117/.+\.0[1-6]9901$", re.I)
SKIP_TITLE_RE = re.compile(
    r"(issue publication information|issue editorial masthead|"
    r"editorial (board|masthead)|table of contents|cover image|"
    r"front cover|inside cover|back cover|masthead|about the cover)",
    re.I,
)
SKIP_DOCTYPE_RE = re.compile(
    r"^(correction|retraction|biographical|news item|book review)$",
    re.I,
)

_HAS_CALAMINE = False
try:
    import python_calamine  # noqa: F401

    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def read_table(path: Path) -> pd.DataFrame:
    """只读打开 Excel，不写入源文件。"""
    suffix = path.suffix.lower()
    if _HAS_CALAMINE:
        return pd.read_excel(path, engine="calamine")
    engine = "openpyxl" if suffix == ".xlsx" else "xlrd"
    return pd.read_excel(path, engine=engine)


def pick_col(columns, candidates) -> str | None:
    colset = list(columns)
    for name in candidates:
        if name in colset:
            return name
    lower = {str(c).strip().lower(): c for c in colset}
    for name in candidates:
        hit = lower.get(name.lower())
        if hit is not None:
            return hit
    return None


def normalize_doi(raw) -> str:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    text = str(raw).strip()
    if not text or text.lower() in ("nan", "none"):
        return ""
    text = text.replace("https://doi.org/", "").replace("http://doi.org/", "")
    text = text.replace("https://dx.doi.org/", "").replace("http://dx.doi.org/", "")
    text = text.strip().strip("/")
    if not DOI_RE.match(text):
        return ""
    return text


def skip_reason(doi: str, title: str, doctype: str) -> str:
    """非论文 / 期次信息，无需请求网络。"""
    if ISSUE_DOI_RE.match(doi or ""):
        return "ACS 期次级 DOI，不是论文"
    if SPIE_FRONT_MATTER_RE.match(doi or ""):
        return "SPIE 封面/补白 DOI，不是论文"
    if (doi or "").lower().count("10.") >= 2:
        return "DOI 格式异常"
    if title and SKIP_TITLE_RE.search(title):
        return f"非论文题名: {title[:80]}"
    if doctype and SKIP_DOCTYPE_RE.match(doctype.strip()):
        return f"跳过文献类型: {doctype}"
    return ""


def doi_filename(doi: str) -> str:
    """Windows 文件名：斜杠等非法字符换成 _，扩展名 .pdf。"""
    stem = ILLEGAL_FS.sub("_", doi).strip(" ._")
    return (stem[:180] or "unnamed") + ".pdf"


def is_pdf_bytes(data: bytes) -> bool:
    return data[:8].startswith(b"%PDF")


def find_pdf_urls(html: str, page_url: str) -> list[str]:
    """从页面解析 PDF 链接（对齐 sci.py 的正则思路，并补 citation_pdf_url）。"""
    urls: list[str] = []
    seen = set()

    def add(u: str):
        u = (u or "").strip().strip("'\"")
        if not u:
            return
        full = urljoin(page_url, u)
        if full not in seen:
            seen.add(full)
            urls.append(full)

    for m in re.finditer(
        r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)["\']',
        html,
        re.I,
    ):
        add(m.group(1))
    for m in re.finditer(
        r'<meta[^>]+content=["\']([^"\']+\.pdf[^"\']*)["\'][^>]+name=["\']citation_pdf_url["\']',
        html,
        re.I,
    ):
        add(m.group(1))
    for m in re.finditer(r'(?:href|src)=["\']([^"\']+\.pdf[^"\']*)["\']', html, re.I):
        add(m.group(1))
    for m in re.finditer(r'["\'](https?:[^"\']+\.pdf[^"\']*)["\']', html, re.I):
        add(m.group(1))
    # sci.py: /path/file.pdf
    for m in re.finditer(r'(//[^\'"\s<>]+\.pdf(?:\?[^\'"\s<>]*)?)', html, re.I):
        add("https:" + m.group(1) if m.group(1).startswith("//") else m.group(1))
    return urls


def is_challenge_page(data: bytes) -> bool:
    head = data[:8000].lower()
    if any(m in head for m in CLOUDFLARE_MARKERS):
        return True
    if any(m in head for m in HCAPTCHA_MARKERS):
        return True
    if len(data) < 2500 and b"robots" in head and b"noindex" in head:
        return True
    return False


def is_likely_blocked(url: str) -> bool:
    u = (url or "").lower()
    return any(h in u for h in BLOCKED_HOST_HINTS)


def is_openalex_work_page(url: str) -> bool:
    """OpenAlex 作品页（https://openalex.org/W...）是 HTML，不是 PDF。"""
    parsed = urlparse(url or "")
    host = (parsed.hostname or "").lower()
    path = parsed.path or ""
    if host != "openalex.org" and not host.endswith(".openalex.org"):
        return False
    if host.startswith("api.") or host.startswith("content."):
        return False
    return bool(re.search(r"/W\d+", path, re.I)) and ".pdf" not in path.lower()


def is_cloudflare_host(url: str) -> bool:
    """doi.org / OpenAlex 等已知 Cloudflare 保护站点。"""
    host = (urlparse(url).hostname or "").lower()
    if not host:
        return False
    for hint in CLOUDFLARE_HOST_HINTS:
        if host == hint or host.endswith("." + hint):
            return True
    return False


# ================= 浏览器验证：Cloudflare / hCaptcha =================

_CHROME_MAJOR = None


def is_cloudflare_challenge(html: str) -> bool:
    if not html:
        return False
    lower = html.lower()
    return any(m in lower for m in CF_MARKERS_STR)


def is_hcaptcha_challenge(html: str) -> bool:
    if not html:
        return False
    lower = html.lower()
    return any(m in lower for m in HCAPTCHA_MARKERS_STR)


def page_has_bot_challenge(html: str) -> bool:
    return is_cloudflare_challenge(html) or is_hcaptcha_challenge(html)


def challenge_kind(html: str | bytes | None) -> str:
    """返回 'hcaptcha' / 'cloudflare' / ''。"""
    if html is None:
        return ""
    if isinstance(html, bytes):
        head = html[:8000].lower()
        if any(m in head for m in HCAPTCHA_MARKERS):
            return "hcaptcha"
        if any(m in head for m in CLOUDFLARE_MARKERS):
            return "cloudflare"
        return ""
    if is_hcaptcha_challenge(html):
        return "hcaptcha"
    if is_cloudflare_challenge(html):
        return "cloudflare"
    return ""


def _chrome_major_version() -> int:
    """读取本机 Chrome 主版本，避免 UC 拉取不匹配的 ChromeDriver。"""
    global _CHROME_MAJOR
    if _CHROME_MAJOR is not None:
        return _CHROME_MAJOR

    version = None
    if winreg is not None:
        for root, path in (
            (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Google\Chrome\BLBeacon"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon"),
        ):
            try:
                with winreg.OpenKey(root, path) as key:
                    version, _ = winreg.QueryValueEx(key, "version")
                    break
            except OSError:
                continue

    if not version:
        candidates = [
            os.environ.get("CHROME_PATH"),
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
        for exe in candidates:
            if not exe or not os.path.isfile(exe):
                continue
            try:
                out = subprocess.check_output(
                    [exe, "--version"], stderr=subprocess.STDOUT, text=True, timeout=10
                )
                if re.search(r"(\d+)\.", out):
                    version = out.strip()
                    break
            except Exception:
                continue

    if version:
        m = re.search(r"(\d+)\.", str(version))
        if m:
            _CHROME_MAJOR = int(m.group(1))
            return _CHROME_MAJOR

    _CHROME_MAJOR = 0
    return _CHROME_MAJOR


def _sync_cookies_to_session(driver, session: requests.Session, url: str) -> None:
    for cookie in driver.get_cookies():
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain") or urlparse(url).hostname,
            path=cookie.get("path") or "/",
        )


def _try_click_hcaptcha(driver) -> bool:
    """切入 hCaptcha 复选框 iframe 并点击。出现图片题时返回 False。"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.wait import WebDriverWait

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    checkbox_iframes = []
    try:
        checkbox_iframes = driver.find_elements(
            By.CSS_SELECTOR,
            "iframe[src*='hcaptcha.com'][src*='frame=checkbox'],"
            "iframe[title*='hCaptcha'][title*='复选框'],"
            "iframe[title*='checkbox'],"
            "iframe[data-hcaptcha-widget-id]",
        )
    except Exception:
        checkbox_iframes = []

    if not checkbox_iframes:
        try:
            checkbox_iframes = [
                f
                for f in driver.find_elements(By.TAG_NAME, "iframe")
                if "hcaptcha" in ((f.get_attribute("src") or "") + (f.get_attribute("title") or "")).lower()
                and "challenge" not in (f.get_attribute("src") or "").lower()
            ]
        except Exception:
            checkbox_iframes = []

    if not checkbox_iframes:
        return False

    clicked = False
    try:
        driver.switch_to.frame(checkbox_iframes[0])
        wait = WebDriverWait(driver, 8)
        for selector in ("#checkbox", "div#checkbox", "#anchor-state", ".check"):
            try:
                el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                el.click()
                clicked = True
                break
            except Exception:
                continue
        if not clicked:
            try:
                driver.find_element(By.TAG_NAME, "body").click()
                clicked = True
            except Exception:
                pass
    except Exception:
        clicked = False
    finally:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

    if not clicked:
        return False

    time.sleep(2.0)
    # 若弹出图片挑战 iframe，自动无法继续
    try:
        challenge_frames = driver.find_elements(
            By.CSS_SELECTOR,
            "iframe[src*='hcaptcha.com'][src*='frame=challenge'],"
            "iframe[title*='挑战'],"
            "iframe[title*='challenge']",
        )
        visible = []
        for fr in challenge_frames:
            try:
                if fr.is_displayed() and fr.size.get("height", 0) > 50:
                    visible.append(fr)
            except Exception:
                continue
        if visible:
            print("  [hCaptcha] 出现图片点选挑战，自动跳过")
            return False
    except Exception:
        pass
    return True


def disable_browser(reason: str) -> None:
    """Chrome 起不来时，本进程不再反复尝试，避免每条 DOI 卡 25 秒。"""
    global USE_BROWSER_CF, _BROWSER_DISABLED
    if _BROWSER_DISABLED:
        return
    _BROWSER_DISABLED = True
    USE_BROWSER_CF = False
    print(f"  [验证] {reason}；本进程不再启动浏览器，继续只用仓储直链")


def _free_debug_port() -> int:
    global _DEBUG_PORT_SEQ
    _DEBUG_PORT_SEQ += 1
    base = CHROME_DEBUG_PORT_BASE + (os.getpid() % 80) + _DEBUG_PORT_SEQ
    for port in range(base, base + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _create_uc_driver(proxy: str | None = None):
    if _BROWSER_DISABLED or not USE_BROWSER_CF:
        return None
    if uc is None:
        print("  [验证] 缺少 undetected-chromedriver，请运行: pip install undetected-chromedriver")
        disable_browser("未安装 undetected-chromedriver")
        return None
    os.makedirs(CHROME_PROFILE_ROOT, exist_ok=True)
    profile = tempfile.mkdtemp(prefix=f"doi_{os.getpid()}_", dir=CHROME_PROFILE_ROOT)
    debug_port = _free_debug_port()
    options = uc.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--no-service-autorun")
    options.add_argument("--password-store=basic")
    options.add_argument(f"--user-data-dir={profile}")
    options.add_argument(f"--remote-debugging-port={debug_port}")
    if proxy:
        options.add_argument(f"--proxy-server={proxy}")
    kwargs = {"options": options, "user_data_dir": profile}
    major = _chrome_major_version()
    if major:
        kwargs["version_main"] = major
    print(f"  [验证] 独立 Chrome profile={profile}  debug-port={debug_port}")
    holder: list = []

    def launch():
        try:
            holder.append(uc.Chrome(**kwargs))
        except Exception as exc:
            holder.append(exc)

    thread = threading.Thread(target=launch, daemon=True)
    thread.start()
    thread.join(BROWSER_LAUNCH_TIMEOUT)
    if thread.is_alive():
        disable_browser(
            f"Chrome {BROWSER_LAUNCH_TIMEOUT}s 内未启动（可能与已有 Chrome/chromedriver 冲突）"
        )
        return None
    if not holder:
        disable_browser("Chrome 启动无返回")
        return None
    if isinstance(holder[0], Exception):
        print(f"  [验证] 启动失败: {holder[0]}")
        disable_browser("Chrome 启动异常")
        return None
    return holder[0]


def bypass_browser_challenge(
    session: requests.Session,
    url: str,
    proxy: str | None = None,
    wait_timeout: int = CF_WAIT_TIMEOUT,
) -> tuple[bool, str, str]:
    """
    处理 Cloudflare / hCaptcha。
    返回 (ok, page_html, final_url)；成功时 Cookie 已写入 session。
    """
    if _BROWSER_DISABLED or not USE_BROWSER_CF:
        return False, "", url
    driver = None
    kind_hint = ""
    try:
        driver = _create_uc_driver(proxy=proxy)
        if driver is None:
            return False, "", url
        driver.set_page_load_timeout(wait_timeout)
        major = _chrome_major_version()
        print(
            f"  [验证] 启动浏览器"
            + (f" (Chrome {major})" if major else "")
            + f": {url}"
        )
        driver.get(url)
        time.sleep(1.5)

        deadline = time.time() + wait_timeout
        hcaptcha_clicked = False
        while time.time() < deadline:
            html = driver.page_source or ""
            kind = challenge_kind(html) or kind_hint
            if kind:
                kind_hint = kind

            if not page_has_bot_challenge(html):
                _sync_cookies_to_session(driver, session, driver.current_url or url)
                label = kind_hint or "验证"
                print(f"  [{label}] 验证通过，Cookie 已同步")
                return True, html, driver.current_url or url

            if kind == "hcaptcha" and not hcaptcha_clicked:
                print("  [hCaptcha] 检测到复选框，尝试点击…")
                if _try_click_hcaptcha(driver):
                    hcaptcha_clicked = True
                    print("  [hCaptcha] 已点击复选框，等待通过…")
                else:
                    print("  [hCaptcha] 无法自动完成验证")
                    _sync_cookies_to_session(driver, session, driver.current_url or url)
                    return False, html, driver.current_url or url

            time.sleep(1.5)

        html = driver.page_source or ""
        _sync_cookies_to_session(driver, session, driver.current_url or url)
        label = kind_hint or challenge_kind(html) or "验证"
        print(f"  [{label}] 等待验证超时")
        return False, html, driver.current_url or url
    except Exception as e:
        print(f"  [验证] 处理失败: {e}")
        return False, "", url
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


# 兼容旧名
def bypass_cloudflare(session, url, proxy=None, wait_timeout=CF_WAIT_TIMEOUT) -> bool:
    ok, _, _ = bypass_browser_challenge(session, url, proxy=proxy, wait_timeout=wait_timeout)
    return ok


def _request_bytes(url: str, timeout: int) -> requests.Response | None:
    try:
        resp = SESSION.get(url, timeout=timeout, allow_redirects=True)
    except Exception:
        return None
    if resp.status_code >= 400:
        return None
    return resp


def _request_bytes_raw(url: str, timeout: int) -> requests.Response | None:
    """与 _request_bytes 类似，但保留 4xx 响应体（便于识别 CF 验证页）。"""
    try:
        return SESSION.get(url, timeout=timeout, allow_redirects=True)
    except Exception:
        return None


class _HtmlResponse:
    """用浏览器拿到的 HTML 伪装成轻量 Response，供 fetch_pdf 解析。"""

    def __init__(self, url: str, html: str):
        self.url = url
        self.content = html.encode("utf-8", errors="ignore")
        self.encoding = "utf-8"
        self.headers = {"Content-Type": "text/html; charset=utf-8"}
        self.status_code = 200


def _request_bytes_with_cf(url: str, timeout: int, *, allow_cf: bool | None = None):
    """
    请求 URL；默认不用浏览器。
    --browser 开启后，才对 doi.org / content.openalex.org 等过 Cloudflare。
    ACS/IEEE/SPIE 等付费站以及 OpenAlex 作品页一律不启动浏览器。
    """
    if is_openalex_work_page(url):
        return None
    if allow_cf is None:
        allow_cf = USE_BROWSER_CF and not _BROWSER_DISABLED

    raw = _request_bytes_raw(url, timeout)
    final_url = (raw.url if raw is not None else None) or url
    # doi.org 常 302 到 ACS/SPIE；按最终落地站判断，绝不为付费站开 Chrome
    if is_likely_blocked(url) or is_likely_blocked(final_url):
        allow_cf = False
    data = (raw.content if raw is not None else b"") or b""
    ok_http = raw is not None and raw.status_code < 400
    challenged = bool(data) and is_challenge_page(data)
    force_cf = is_cloudflare_host(url) or is_cloudflare_host(final_url)

    # 已拿到 PDF，直接返回
    if ok_http and is_pdf_bytes(data) and len(data) > 2048:
        return raw

    # 普通站点且无验证：直接返回
    if ok_http and not challenged and not force_cf:
        return raw

    # 已知 CF 站点但本次未触发验证页：信任本次结果（可能已有 Cookie）
    if ok_http and not challenged and force_cf:
        return raw

    # 需要浏览器过验证
    need_browser = allow_cf and (challenged or (force_cf and not ok_http))
    if not need_browser:
        return raw if ok_http else None

    start_url = final_url
    if is_likely_blocked(start_url):
        return raw if ok_http else None
    kind = challenge_kind(data) or ("cloudflare" if force_cf else "验证")
    print(f"  [{kind}] 对 {urlparse(start_url).hostname} 使用浏览器验证…")
    ok, page_html, final_url = bypass_browser_challenge(SESSION, start_url)
    if not ok and not page_html:
        return None

    # 优先用已同步 Cookie 再请求一次（跟随后续跳转）
    resp2 = _request_bytes(final_url or start_url, timeout)
    if resp2 is not None and not is_challenge_page(resp2.content or b""):
        return resp2

    if page_html and not page_has_bot_challenge(page_html):
        return _HtmlResponse(final_url or start_url, page_html)
    return None


def fetch_pdf(url: str, timeout: int = 90) -> bytes | None:
    resp = _request_bytes_with_cf(url, timeout)
    if resp is None:
        return None
    data = resp.content or b""
    if is_pdf_bytes(data) and len(data) > 2048:
        return data
    ctype = (resp.headers.get("Content-Type") or "").lower()
    if "pdf" in ctype and len(data) > 2048 and not is_challenge_page(data):
        return data
    html_like = (
        "html" in ctype
        or data.lstrip()[:15].lower().startswith(b"<!doctype")
        or data.lstrip()[:6].lower().startswith(b"<html")
    )
    if not html_like:
        return None
    try:
        html = data.decode(resp.encoding or "utf-8", errors="ignore")
    except Exception:
        html = ""
    for pdf_url in find_pdf_urls(html, resp.url)[:8]:
        pdf_resp = _request_bytes_with_cf(pdf_url, timeout)
        if pdf_resp is None:
            continue
        blob = pdf_resp.content or b""
        if is_pdf_bytes(blob) and len(blob) > 2048:
            return blob
    return None


def unpaywall_pdf_urls(doi: str, email: str, timeout: int = 30) -> list[str]:
    if not email:
        return []
    api = f"https://api.unpaywall.org/v2/{doi}"
    try:
        resp = SESSION.get(api, params={"email": email}, timeout=timeout)
        if resp.status_code == 404:
            return []
        resp.raise_for_status()
        data = resp.json()
    except Exception:
        return []
    urls: list[str] = []
    locs = []
    best = data.get("best_oa_location") or {}
    if best:
        locs.append(best)
    locs.extend(data.get("oa_locations") or [])
    seen = set()
    for loc in locs:
        for key in ("url_for_pdf", "url"):
            u = (loc.get(key) or "").strip()
            if u and u not in seen:
                seen.add(u)
                urls.append(u)
    return urls


def crossref_pdf_urls(doi: str, timeout: int = 30) -> list[str]:
    try:
        resp = SESSION.get(f"https://api.crossref.org/works/{doi}", timeout=timeout)
        if resp.status_code != 200:
            return []
        msg = (resp.json() or {}).get("message") or {}
    except Exception:
        return []
    urls: list[str] = []
    for item in msg.get("link") or []:
        u = (item.get("URL") or "").strip()
        ctype = (item.get("content-type") or "").lower()
        if u and ("pdf" in ctype or u.lower().endswith(".pdf")):
            urls.append(u)
    return urls


def openalex_pdf_urls(doi: str, timeout: int = 30) -> list[str]:
    """从 OpenAlex API 取全文候选；优先 OpenAlex 自托管 PDF / 作品页。"""
    try:
        resp = SESSION.get(
            f"https://api.openalex.org/works/doi:{doi}", timeout=timeout)
        if resp.status_code != 200:
            return []
        data = resp.json() or {}
    except Exception:
        return []

    urls: list[str] = []
    seen: set[str] = set()

    def add(u: str | None):
        u = (u or "").strip()
        if not u or u in seen:
            return
        if is_openalex_work_page(u):
            return
        seen.add(u)
        urls.append(u)

    # 1) OpenAlex 自托管全文（最高优先）；不加入 openalex.org/W... 作品页
    content_urls = data.get("content_urls") or {}
    add(content_urls.get("pdf"))

    work_id = ""
    oid = (data.get("id") or "").strip()
    if oid:
        m = re.search(r"(W\d+)", oid, re.I)
        if m:
            work_id = m.group(1)
    if not work_id:
        m = re.search(r"(W\d+)", str((data.get("ids") or {}).get("openalex") or ""), re.I)
        if m:
            work_id = m.group(1)

    has_content = data.get("has_content") or {}
    if work_id and (has_content.get("pdf") or content_urls.get("pdf")):
        add(f"https://content.openalex.org/works/{work_id}.pdf")

    # 2) OpenAlex 登记的 OA / 落地页（仍属 OpenAlex 来源，后续再按域名排序）
    oa = data.get("open_access") or {}
    add(oa.get("oa_url"))
    for loc_key in ("best_oa_location", "primary_location"):
        loc = data.get(loc_key) or {}
        for key in ("pdf_url", "oa_url", "landing_page_url"):
            add(loc.get(key))
    for loc in data.get("locations") or []:
        for key in ("pdf_url", "landing_page_url"):
            add(loc.get(key))
    return urls


def semantic_scholar_pdf_urls(doi: str, timeout: int = 30) -> list[str]:
    try:
        resp = SESSION.get(
            f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}",
            params={"fields": "isOpenAccess,openAccessPdf"},
            timeout=timeout,
        )
        if resp.status_code != 200:
            return []
        data = resp.json() or {}
    except Exception:
        return []
    pdf = data.get("openAccessPdf") or {}
    u = (pdf.get("url") or "").strip()
    return [u] if u else []


def europepmc_pdf_urls(doi: str, timeout: int = 30) -> list[str]:
    try:
        resp = SESSION.get(
            "https://www.ebi.ac.uk/europepmc/webservices/rest/search",
            params={
                "query": f'DOI:"{doi}"',
                "format": "json",
                "resultType": "core",
                "pageSize": 1,
            },
            timeout=timeout,
        )
        if resp.status_code != 200:
            return []
        hits = ((resp.json() or {}).get("resultList") or {}).get("result") or []
    except Exception:
        return []
    if not hits:
        return []
    rec = hits[0]
    urls: list[str] = []
    pmcid = (rec.get("pmcid") or "").strip()
    if pmcid:
        urls.extend(pmcid_pdf_urls(pmcid))
    ft = rec.get("fullTextUrlList") or {}
    for item in ft.get("fullTextUrl") or []:
        if str(item.get("documentStyle") or "").lower() == "pdf":
            u = (item.get("url") or "").strip()
            if u:
                urls.append(u)
    return urls


def pmcid_from_doi(doi: str, email: str = "", timeout: int = 30) -> str:
    """NCBI ID Converter：DOI -> PMCID。"""
    try:
        resp = SESSION.get(
            "https://pmc.ncbi.nlm.nih.gov/tools/idconv/api/v1/articles/",
            params={
                "ids": doi,
                "idtype": "doi",
                "format": "json",
                "tool": "download_doi_pdfs",
                "email": email or "research@example.edu",
            },
            timeout=timeout,
        )
        if resp.status_code != 200:
            return ""
        recs = (resp.json() or {}).get("records") or []
    except Exception:
        return ""
    if not recs:
        return ""
    pmcid = str(recs[0].get("pmcid") or "").strip()
    if pmcid.lower() in ("", "null", "none"):
        return ""
    return pmcid if pmcid.upper().startswith("PMC") else f"PMC{pmcid}"


def pmcid_pdf_urls(pmcid: str) -> list[str]:
    pmcid = (pmcid or "").strip()
    if not pmcid:
        return []
    if not pmcid.upper().startswith("PMC"):
        pmcid = "PMC" + pmcid
    pmcid = pmcid.upper()
    return [
        f"https://europepmc.org/api/getPdf?pmcid={pmcid}",
        f"https://europepmc.org/articles/{pmcid}?pdf=render",
    ]


def extract_pmcids(urls: list[str]) -> list[str]:
    found: list[str] = []
    seen = set()
    for u in urls:
        for m in re.finditer(r"PMC(\d+)", u or "", re.I):
            pmcid = "PMC" + m.group(1)
            if pmcid not in seen:
                seen.add(pmcid)
                found.append(pmcid)
    return found


def expand_arxiv_urls(urls: list[str]) -> list[str]:
    extra: list[str] = []
    for u in urls:
        m = re.search(r"arxiv\.org/(?:abs|pdf)/([^/\s?#]+)", u or "", re.I)
        if m:
            ident = m.group(1).replace(".pdf", "")
            extra.append(f"https://arxiv.org/pdf/{ident}.pdf")
    return extra


def is_openalex_host(url: str) -> bool:
    host = (urlparse(url).hostname or "").lower()
    return host == "openalex.org" or host.endswith(".openalex.org")


def url_priority(url: str) -> int:
    """数值越小越优先。OpenAlex 域名最优先（站点为 Cloudflare 验证）。"""
    u = (url or "").lower()
    if "content.openalex.org" in u:
        return -20
    if "openalex.org" in u:
        return -10
    if "europepmc.org/api/getpdf" in u:
        return 0
    if "europepmc.org" in u:
        return 1
    if "arxiv.org/pdf" in u:
        return 2
    if any(x in u for x in ("biorxiv.org", "chemrxiv.org", "medrxiv.org")):
        return 3
    if "pmc.ncbi.nlm.nih.gov" in u or "/pmc/articles/" in u:
        return 8
    if "pubs.acs.org" in u or "spiedigitallibrary" in u or "ieeexplore" in u:
        return 70
    if "doi.org/" in u:
        return 80
    return 40


def collect_candidate_urls(doi: str, email: str) -> tuple[list[str], list[str]]:
    """返回 (候选 URL, 已查询的来源说明)。OpenAlex 最优先，仓储次之，出版社最后。"""
    urls: list[str] = []
    seen = set()
    sources: list[str] = []

    def add_all(label: str, found: list[str]):
        added = 0
        for u in found:
            u = (u or "").strip()
            if u and u not in seen:
                seen.add(u)
                urls.append(u)
                added += 1
        sources.append(f"{label}:{added}")

    # OpenAlex 最先查询并参与排序（含 openalex.org / content.openalex.org）
    add_all("openalex", openalex_pdf_urls(doi))
    pmcid = pmcid_from_doi(doi, email)
    add_all("pmcid", pmcid_pdf_urls(pmcid) if pmcid else [])
    add_all("unpaywall", unpaywall_pdf_urls(doi, email))
    add_all("semanticscholar", semantic_scholar_pdf_urls(doi))
    add_all("europepmc", europepmc_pdf_urls(doi))
    add_all("crossref", crossref_pdf_urls(doi))

    for extra_pmcid in extract_pmcids(urls):
        for u in pmcid_pdf_urls(extra_pmcid):
            if u not in seen:
                seen.add(u)
                urls.append(u)
    for u in expand_arxiv_urls(urls):
        if u not in seen:
            seen.add(u)
            urls.append(u)

    doi_url = f"https://doi.org/{doi}"
    if doi_url not in seen:
        urls.append(doi_url)
        sources.append("doi.org:1")

    urls.sort(key=url_priority)
    return urls, sources


def save_pdf(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".part")
    tmp.write_bytes(data)
    tmp.replace(path)


def iter_record_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES:
            if path.name.startswith("~$") or path.name.startswith("_"):
                continue
            files.append(path)
    files.sort()
    return files


def domain_journal_of(path: Path, root: Path) -> tuple[str, str]:
    try:
        rel = path.relative_to(root)
    except ValueError:
        return "", path.parent.name
    parts = rel.parts
    domain = parts[0] if len(parts) >= 2 else ""
    journal = parts[1] if len(parts) >= 3 else path.parent.name
    return domain, journal


def load_jobs(root: Path, only_domain: str = "") -> list[dict]:
    jobs: list[dict] = []
    seen = set()
    files = iter_record_files(root)
    for fpath in files:
        domain, journal = domain_journal_of(fpath, root)
        if only_domain and domain != only_domain:
            continue
        try:
            df = read_table(fpath)
        except Exception as e:
            print(f"读表失败（跳过，不改原文件）: {fpath} | {e}")
            continue
        doi_col = pick_col(df.columns, DOI_COL_CANDIDATES)
        if not doi_col:
            print(f"无 DOI 列，跳过: {fpath}")
            continue
        title_col = pick_col(df.columns, TITLE_COL_CANDIDATES)
        doctype_col = pick_col(df.columns, DOCTYPE_COL_CANDIDATES)
        for idx, row in df.iterrows():
            doi = normalize_doi(row.get(doi_col))
            if not doi:
                continue
            key = doi.casefold()
            if key in seen:
                continue
            seen.add(key)
            title = ""
            if title_col is not None:
                val = row.get(title_col)
                if val is not None and not (isinstance(val, float) and pd.isna(val)):
                    title = str(val).strip()
            doctype = ""
            if doctype_col is not None:
                val = row.get(doctype_col)
                if val is not None and not (isinstance(val, float) and pd.isna(val)):
                    doctype = str(val).strip()
            jobs.append(
                {
                    "doi": doi,
                    "domain": domain,
                    "journal": journal,
                    "title": title,
                    "doctype": doctype,
                    "skip": skip_reason(doi, title, doctype),
                    "source_file": str(fpath),
                    "row": int(idx) + 2,
                }
            )
    return jobs


def append_log(log_path: Path, row: dict) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "time", "status", "doi", "domain", "journal", "title",
        "pdf_path", "message", "source_file",
    ]
    new_file = not log_path.is_file()
    with log_path.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        if new_file:
            writer.writeheader()
        writer.writerow(row)


def already_ok(pdf_path: Path) -> bool:
    try:
        if not pdf_path.is_file() or pdf_path.stat().st_size < 1024:
            return False
        with pdf_path.open("rb") as f:
            return is_pdf_bytes(f.read(8))
    except Exception:
        return False


def download_one(job: dict, out_root: Path, email: str) -> dict:
    doi = job["doi"]
    pdf_path = out_root / job["domain"] / job["journal"] / doi_filename(doi)
    rec = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "doi": doi,
        "domain": job["domain"],
        "journal": job["journal"],
        "title": job["title"],
        "source_file": job["source_file"],
        "pdf_path": str(pdf_path),
        "status": "fail",
        "message": "",
    }
    if already_ok(pdf_path):
        rec["status"] = "skip_exists"
        rec["message"] = "已有 PDF"
        return rec

    last_err = "无开放获取 PDF"
    urls, sources = collect_candidate_urls(doi, email)
    # OpenAlex 作品页不是 PDF；仓储优先，出版社付费站最后且不弹浏览器
    rest = [u for u in urls if not is_openalex_work_page(u)]
    preferred = [u for u in rest if not is_likely_blocked(u)]
    blocked = [u for u in rest if is_likely_blocked(u)]
    ordered = preferred + blocked
    if not ordered:
        rec["status"] = "fail"
        rec["message"] = (
            "无可用仓储PDF（出版社/HAL 等站点有反爬） | " + ",".join(sources)
        )
        rec["pdf_path"] = ""
        return rec
    for url in ordered:
        timeout = 15 if is_likely_blocked(url) else 90
        data = fetch_pdf(url, timeout=timeout)
        if not data:
            host = urlparse(url).netloc or url
            if is_openalex_work_page(url):
                continue
            if is_cloudflare_host(url):
                kind = "Cloudflare"
            elif is_likely_blocked(url):
                kind = "hCaptcha/反爬"
            else:
                kind = "取数失败"
            last_err = f"未能从 {host} 取到 PDF（{kind}）"
            continue
        if not is_pdf_bytes(data):
            last_err = f"{urlparse(url).netloc} 返回非 PDF"
            continue
        save_pdf(pdf_path, data)
        rec["status"] = "ok"
        rec["message"] = f"{len(data)} bytes from {urlparse(url).netloc}"
        rec["pdf_path"] = str(pdf_path)
        return rec
    rec["status"] = "fail"
    rec["message"] = last_err + " | " + ",".join(sources)
    rec["pdf_path"] = ""
    return rec


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按 DOI 下载九类期刊开放获取 PDF")
    parser.add_argument("--records", type=Path, default=RECORDS_ROOT)
    parser.add_argument("--out", type=Path, default=OUTPUT_ROOT)
    parser.add_argument("--domain", type=str, default="", help="只处理某一个领域")
    parser.add_argument("--doi", type=str, default="", help="只处理包含该字符串的 DOI")
    parser.add_argument("--limit", type=int, default=0, help="最多处理多少条 DOI")
    parser.add_argument("--email", type=str, default=os.environ.get("UNPAYWALL_EMAIL", ""))
    parser.add_argument("--sleep", type=float, default=0.8, help="每篇间隔秒")
    parser.add_argument("--resume", action="store_true", default=True)
    parser.add_argument("--no-resume", dest="resume", action="store_false")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument(
        "--browser",
        action="store_true",
        help="遇 doi.org Cloudflare 时另开独立 Chrome（默认关闭；ACS/SPIE/IEEE 等出版社站点不会开浏览器）",
    )
    parser.add_argument(
        "--no-browser",
        dest="browser",
        action="store_false",
        help="不启动浏览器（默认）",
    )
    parser.set_defaults(browser=False)
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    if not args.records.is_dir():
        print(f"题录目录不存在: {args.records}")
        return 1

    print(f"题录目录（只读）: {args.records}")
    print(f"PDF 保存: {args.out}")
    print(f"Excel 引擎: {'calamine' if _HAS_CALAMINE else 'openpyxl/xlrd'}")
    global USE_BROWSER_CF
    USE_BROWSER_CF = bool(args.browser)
    if USE_BROWSER_CF:
        print(
            "浏览器验证: 开启（仅 doi.org / OpenAlex 直链；"
            "ACS/SPIE/IEEE 等出版社不弹窗；"
            f"启动超时 {BROWSER_LAUNCH_TIMEOUT}s，失败后自动关闭）"
        )
    else:
        print("浏览器验证: 关闭（只走仓储/OA 直链；与已打开的 Chrome 不冲突）")
    jobs = load_jobs(args.records, args.domain.strip())
    doi_filter = args.doi.strip().casefold()
    if doi_filter:
        jobs = [j for j in jobs if doi_filter in j["doi"].casefold()]
    skip_jobs = [j for j in jobs if j.get("skip")]
    fetch_jobs = [j for j in jobs if not j.get("skip")]
    if args.limit and args.limit > 0:
        fetch_jobs = fetch_jobs[: args.limit]
    if not jobs:
        print("没有可处理的 DOI")
        return 1

    by_domain: dict[str, int] = {}
    for job in fetch_jobs:
        by_domain[job["domain"]] = by_domain.get(job["domain"], 0) + 1
    print(f"题录 DOI: {len(jobs)} 条（跳过期次/非论文 {len(skip_jobs)}，待获取 {len(fetch_jobs)}）")
    for name, n in by_domain.items():
        print(f"  {name}: {n}")

    if args.dry_run:
        for i, job in enumerate(fetch_jobs[:30], 1):
            print(f"  {i}. [{job['domain']}] {job['doi']}  ->  {doi_filename(job['doi'])}")
        if len(fetch_jobs) > 30:
            print(f"  ... 其余 {len(fetch_jobs) - 30} 条")
        if skip_jobs:
            print(f"示例跳过: {skip_jobs[0]['doi']} | {skip_jobs[0]['skip']}")
        return 0

    if not args.email:
        print("未提供 Unpaywall 邮箱，将只尝试 Crossref / doi.org。建议加 --email")

    SESSION.headers["User-Agent"] = (
        HEADERS["User-Agent"] + (f" mailto:{args.email}" if args.email else "")
    )

    log_path = args.out / LOG_NAME
    stats = {"ok": 0, "fail": 0, "skip_exists": 0, "skip_invalid": 0}
    args.out.mkdir(parents=True, exist_ok=True)

    skip_path = args.out / "_skipped_invalid.csv"
    skip_path.parent.mkdir(parents=True, exist_ok=True)
    if skip_jobs:
        with skip_path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["doi", "domain", "journal", "title", "message", "source_file"],
            )
            writer.writeheader()
            for job in skip_jobs:
                writer.writerow(
                    {
                        "doi": job["doi"],
                        "domain": job["domain"],
                        "journal": job["journal"],
                        "title": job["title"],
                        "message": job["skip"],
                        "source_file": job["source_file"],
                    }
                )
        stats["skip_invalid"] = len(skip_jobs)
        print(f"已记录跳过 {len(skip_jobs)} 条期次/非论文 DOI -> {skip_path}")

    for i, job in enumerate(fetch_jobs, 1):
        pdf_path = args.out / job["domain"] / job["journal"] / doi_filename(job["doi"])
        if args.resume and already_ok(pdf_path):
            rec = {
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "status": "skip_exists",
                "doi": job["doi"],
                "domain": job["domain"],
                "journal": job["journal"],
                "title": job["title"],
                "pdf_path": str(pdf_path),
                "message": "已有 PDF",
                "source_file": job["source_file"],
            }
            print(f"[{i}/{len(fetch_jobs)}] 已存在，跳过 {job['doi']}")
            stats["skip_exists"] += 1
            append_log(log_path, rec)
            continue
        print(f"[{i}/{len(fetch_jobs)}] {job['doi']}")
        rec = download_one(job, args.out, args.email.strip())
        stats[rec["status"]] = stats.get(rec["status"], 0) + 1
        print(f"  {rec['status']}: {rec['message']}")
        append_log(log_path, rec)
        time.sleep(max(0.0, args.sleep))

    print("\n完成（原题录表未修改）:", json.dumps(stats, ensure_ascii=False))
    print(f"日志: {log_path}")
    return 0 if stats.get("fail", 0) == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())

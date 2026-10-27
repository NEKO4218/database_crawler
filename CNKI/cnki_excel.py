# -*- coding: utf-8 -*-
"""从知网题录 Excel 的 URL 批量下载 PDF。

题录来源：D:\\2026年工作\\知网数据库下载 下各子文件夹中的 CNKI-*.xls/xlsx
验证码：复用 cnki_captcha.py（经 cnkiforexcel.handle_captcha）
机构登录：复用 cnkiforexcel.ensure_unit_login_if_needed
下载点击：对齐 图书情报/batch_download_cnki.py 的 pdfDown 流程

每个 Excel 使用锁文件 + 完成标记：
  {excel}.lock  处理中（进程锁；异常退出后可检测 PID 失效并收回）
  {excel}.done  本表全部成功/已存在后写入；下次自动跳过
手动重启会跳过：已有 .done 的表、已有 PDF 或已记下作者单位的条目。

处理某表时先把题录复制到指定目录（HTML .xls 转为可写的 .xlsx），
下载 PDF 的同时从详情页 wx-tit 抓取作者标注与单位对应，逐条写回副本。

示例：
  python cnki_excel.py
  python cnki_excel.py --folder 信息科技
  python cnki_excel.py --folder 信息科技 --folder 基础科学
  python cnki_excel.py --folder 信息科技 --out D:\\pdf全文 --excel-out D:\\题录副本
  python cnki_excel.py --folder 英文期刊 --limit-files 1 --limit-papers 3
"""
from __future__ import annotations

import argparse
import csv
import gc
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from html import unescape
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from cnki_captcha import (
    TENCENT_RATE_LIMIT_SLEEP,
    TENCENT_RELOGIN_RETRIES,
    has_visible_captcha_widget,
    preload_click_captcha,
    tencent_last_fail,
)
from cnkiforexcel import (
    ensure_captcha_passed,
    ensure_cnki_captcha_after_nav,
    ensure_unit_login_if_needed,
    get_login_home_url,
    has_unit_login_anywhere,
    is_cnki_404_url,
    is_driver_alive,
    open_cnki_url,
    set_chrome_download_dir,
    unwrap_webvpn_url,
)

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
RECORDS_ROOT = Path(r'D:\2026年工作\知网数据库下载')
DEFAULT_PDF_ROOT = Path(r'D:\2026年工作\知网数据库下载\_pdf全文')
DEFAULT_EXCEL_OUT = Path(r'D:\2026年工作\知网数据库下载\_题录')
SKIP_DIR_NAMES = {'captcha_debug', '_pdf全文', '_题录', '_locks'}
XLS_PREFIX = 'CNKI-'
EXCEL_SUFFIXES = {'.xls', '.xlsx'}

_SCRIPT_DIR = Path(__file__).resolve().parent
CHROME_PROFILE_ROOT = _SCRIPT_DIR / '.chrome_profiles' / 'cnki_excel'
CHROME_DEBUG_PORT_BASE = 9224

PAGE_TIMEOUT = 30
MAX_RETRY = 3
NEW_WINDOW_WAIT = 15
PDF_WAIT = 90
SESSION_DEAD_MARKS = (
    'max retries exceeded',
    'failed to establish a new connection',
    'connection refused',
    '积极拒绝',
    'winerror 10061',
    'httpconnectionpool',
    'session deleted',
    'invalid session id',
    'chrome not reachable',
    'disconnected: not connected to devtools',
    'unable to connect to renderer',
    'connectionreseterror',
    'session not created',
)


def is_session_dead_error(exc) -> bool:
    text = str(exc or '').lower()
    return any(m in text for m in SESSION_DEAD_MARKS)

URL_COLUMNS = ('URL-网址', 'URL', '网址', 'Link', '链接')
TITLE_COLUMNS = ('Title-题名', 'Title', '题名', '篇名')
YEAR_COLUMNS = ('Year-年', '年', 'PubTime-发表时间')
COL_AUTHOR_MARK = '作者标注'
COL_ORG_LIST = '单位列表'
COL_AUTHOR_ORG = '作者单位对应'
AUTHOR_ORG_COLS = (COL_AUTHOR_MARK, COL_ORG_LIST, COL_AUTHOR_ORG)
MARK_URL_404 = '网址失效404'
MARK_NO_AUTHOR_ORG = '无作者单位'
MARK_UNAVAILABLE = '机构未购入'
MARK_CATALOG_ONLY = '仅题录'
MARK_NO_PDF = '无PDF'
SKIP_ALL_MARKS = (MARK_URL_404,)
SKIP_PDF_MARKS = (MARK_URL_404, MARK_UNAVAILABLE, MARK_CATALOG_ONLY, MARK_NO_PDF)
CNKI_URL_PREFIXES = (
    'https://link.cnki.net',
    'http://link.cnki.net',
    'https://kns.cnki.net',
    'http://kns.cnki.net',
)

_chrome_profile_dir = None
_debug_port_seq = 0


def _setup_stdio():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')


def pid_is_running(pid) -> bool:
    """判断 PID 是否仍在运行（用于回收异常退出留下的锁）。"""
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if os.name == 'nt':
        import ctypes
        from ctypes import wintypes
        process_query = 0x1000
        still_active = 259
        handle = ctypes.windll.kernel32.OpenProcess(process_query, False, pid)
        if not handle:
            return False
        code = wintypes.DWORD()
        ok = ctypes.windll.kernel32.GetExitCodeProcess(handle, ctypes.byref(code))
        ctypes.windll.kernel32.CloseHandle(handle)
        return bool(ok) and code.value == still_active
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


# ---------------------------------------------------------------------------
# Excel 锁 / 完成标记
# ---------------------------------------------------------------------------
class ExcelLock:
    """单表互斥锁 + 完成标记。

    Windows 用 msvcrt 字节锁：进程异常退出后操作系统会释放文件锁。
    若残留 .lock 且其中 PID 已不存在，则删除后重新占用，避免整表被永久跳过。
    """

    def __init__(self, excel_path: Path):
        self.excel_path = Path(excel_path)
        self.lock_path = Path(str(self.excel_path) + '.lock')
        self.done_path = Path(str(self.excel_path) + '.done')
        self._fp = None
        self._acquired = False

    def is_done(self) -> bool:
        return self.done_path.is_file()

    def lock_pid(self) -> int | None:
        try:
            text = self.lock_path.read_text(encoding='utf-8', errors='ignore')
        except OSError:
            return None
        m = re.search(r'pid=(\d+)', text)
        if not m:
            return None
        try:
            return int(m.group(1))
        except ValueError:
            return None

    def is_stale(self) -> bool:
        if not self.lock_path.is_file():
            return False
        pid = self.lock_pid()
        if pid is None:
            return True
        return not pid_is_running(pid)

    def _close_fp(self) -> None:
        if not self._fp:
            return
        try:
            self._fp.close()
        except OSError:
            pass
        self._fp = None

    def _try_steal_stale(self) -> bool:
        if not self.is_stale():
            return False
        pid = self.lock_pid()
        print(
            f'  发现残留锁（pid={pid or "未知"} 已退出），重新占用: '
            f'{self.lock_path.name}'
        )
        try:
            self.lock_path.unlink()
        except OSError:
            return False
        return True

    def acquire(self, _retry: bool = True) -> bool:
        if self._acquired:
            return True
        self.lock_path.parent.mkdir(parents=True, exist_ok=True)
        self._fp = open(self.lock_path, 'a+b')
        try:
            self._fp.seek(0, os.SEEK_END)
            if self._fp.tell() == 0:
                self._fp.write(b'\0')
                self._fp.flush()
            self._fp.seek(0)
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self._fp.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self._fp.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self._close_fp()
            if _retry and self._try_steal_stale():
                return self.acquire(_retry=False)
            return False
        payload = (
            f'pid={os.getpid()}\n'
            f'started={datetime.now().isoformat(timespec="seconds")}\n'
            f'excel={self.excel_path}\n'
        ).encode('utf-8')
        self._fp.seek(1)
        self._fp.write(payload)
        self._fp.flush()
        self._acquired = True
        return True

    def mark_done(self, success: int, fail: int, skip: int, total: int) -> None:
        text = (
            f'finished={datetime.now().isoformat(timespec="seconds")}\n'
            f'pid={os.getpid()}\n'
            f'total={total}\n'
            f'success={success}\n'
            f'fail={fail}\n'
            f'skip={skip}\n'
        )
        self.done_path.write_text(text, encoding='utf-8')

    def release(self) -> None:
        if not self._fp:
            self._acquired = False
            return
        try:
            self._fp.seek(0)
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(self._fp.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self._fp.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        self._close_fp()
        self._acquired = False
        try:
            if self.lock_path.is_file():
                self.lock_path.unlink()
        except OSError:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.release()
        return False


# ---------------------------------------------------------------------------
# 题录扫描 / 解析
# ---------------------------------------------------------------------------
def list_category_folders(root: Path) -> list[Path]:
    if not root.is_dir():
        return []
    folders = []
    for p in sorted(root.iterdir(), key=lambda x: x.name):
        if p.is_dir() and p.name not in SKIP_DIR_NAMES and not p.name.startswith('.'):
            folders.append(p)
    return folders


def iter_excel_files(folder: Path) -> list[Path]:
    files = []
    for path in folder.rglob('*'):
        if any(part in SKIP_DIR_NAMES for part in path.parts):
            continue
        if not path.is_file():
            continue
        if path.name.startswith('~$') or path.name.startswith('_'):
            continue
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            continue
        if not path.name.startswith(XLS_PREFIX):
            continue
        files.append(path)
    files.sort()
    return files


def count_excels(folder: Path) -> int:
    return len(iter_excel_files(folder))


def _strip_html(cell: str) -> str:
    text = re.sub(r'<[^>]+>', '', cell or '')
    text = unescape(text)
    return re.sub(r'\s+', ' ', text).strip()


def _td_texts(tr_html: str) -> list[str]:
    tds = re.findall(r'<td[^>]*>(.*?)</td>', tr_html, flags=re.I | re.S)
    return [_strip_html(td) for td in tds]


def read_cnki_html_xls(path: Path) -> pd.DataFrame:
    raw = path.read_bytes()
    text = None
    for enc in ('utf-8-sig', 'utf-8', 'gb18030'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode('utf-8', errors='replace')
    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', text, flags=re.I | re.S)
    if not rows:
        raise ValueError('未解析到表格行')
    headers = _td_texts(rows[0])
    if not headers:
        raise ValueError('表头为空')
    records = []
    for row_html in rows[1:]:
        cells = _td_texts(row_html)
        if not any(cells):
            continue
        if len(cells) < len(headers):
            cells.extend([''] * (len(headers) - len(cells)))
        records.append(dict(zip(headers, cells[:len(headers)])))
    return pd.DataFrame(records)


def read_excel_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == '.xls':
        head = path.read_bytes()[:200].lower()
        if b'<html' in head or b'<table' in head or b'<tr' in head:
            return read_cnki_html_xls(path)
        return pd.read_excel(path)
    return pd.read_excel(path)


def _pick_column(columns, candidates) -> str | None:
    colset = {str(c): c for c in columns}
    for name in candidates:
        if name in colset:
            return colset[name]
    lower = {str(c).strip().lower(): c for c in columns}
    for name in candidates:
        key = name.lower()
        if key in lower:
            return lower[key]
    return None


def sanitize_filename(name: str, max_len: int = 80) -> str:
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '', name or '')
    name = re.sub(r'\s+', ' ', name).strip(' .')
    if len(name) > max_len:
        name = name[:max_len].rstrip(' .')
    return name or 'paper'


def is_cnki_url(url: str) -> bool:
    u = (url or '').strip()
    if any(u.startswith(p) for p in CNKI_URL_PREFIXES):
        return True
    inner = unwrap_webvpn_url(u)
    if inner != u and any(inner.startswith(p) for p in CNKI_URL_PREFIXES):
        return True
    return 'webvpn.uibe.edu.cn' in u.lower() and 'cnki.net' in inner.lower()


def cell_text(value) -> str:
    if value is None:
        return ''
    text = str(value).strip()
    if text.lower() in ('nan', 'none'):
        return ''
    return text


def has_author_org(value) -> bool:
    return bool(cell_text(value))


def row_has_author_org(row) -> bool:
    """任一作者/单位列有内容即视为已提取，重启后不再抓作者单位。"""
    return any(has_author_org(row.get(col)) for col in AUTHOR_ORG_COLS)


def row_joined_marks(row) -> str:
    return ' '.join(cell_text(row.get(col)) for col in AUTHOR_ORG_COLS)


def row_has_skip_mark(row, marks) -> bool:
    text = row_joined_marks(row)
    return any(m in text for m in marks)


def load_papers_from_df(df: pd.DataFrame) -> list[dict]:
    url_col = _pick_column(df.columns, URL_COLUMNS)
    if not url_col:
        raise ValueError(f'无 URL 列，现有列: {list(df.columns)}')
    title_col = _pick_column(df.columns, TITLE_COLUMNS)
    year_col = _pick_column(df.columns, YEAR_COLUMNS)
    papers = []
    for idx, row in df.iterrows():
        url = cell_text(row.get(url_col))
        if not url or not is_cnki_url(url):
            continue
        title = cell_text(row.get(title_col)) if title_col is not None else ''
        year = cell_text(row.get(year_col)) if year_col is not None else ''
        year = re.sub(r'\D', '', year)[:4]
        safe = sanitize_filename(title)
        if year and len(year) == 4:
            fname = f'{safe}_{year}.pdf'
        else:
            fname = f'{safe}.pdf'
        papers.append({
            'url': url,
            'title': title or safe,
            'filename': fname,
            'row': int(idx) + 2,
            'df_index': idx,
            'has_meta': row_has_author_org(row),
        })
    return papers


def load_papers_from_excel(path: Path) -> list[dict]:
    return load_papers_from_df(read_excel_table(path))


def working_excel_path(excel_path: Path, records_root: Path, excel_out: Path) -> Path:
    rel = relative_category(excel_path, records_root)
    return excel_out / rel / f'{excel_path.stem}.xlsx'


def save_working_excel(path: Path, df: pd.DataFrame) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    illegal = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
    for col in df.columns:
        if df[col].dtype == object or str(df[col].dtype).startswith('string'):
            df[col] = df[col].map(
                lambda v: illegal.sub('', v) if isinstance(v, str) else v)
    tmp = path.with_name(path.name + '.tmp')
    df.to_excel(tmp, index=False, engine='openpyxl')
    tmp.replace(path)


def prepare_working_excel(src: Path, dest: Path) -> pd.DataFrame:
    """复制题录到可写 xlsx，并补齐作者-单位列；已有副本则合并已抓字段。"""
    df = read_excel_table(src)
    for col in AUTHOR_ORG_COLS:
        if col not in df.columns:
            df[col] = ''
        else:
            df[col] = df[col].fillna('').astype(str).replace({'nan': '', 'None': ''})
    dest = Path(dest)
    if dest.is_file():
        try:
            old = pd.read_excel(dest)
            url_col = _pick_column(df.columns, URL_COLUMNS)
            old_url = _pick_column(old.columns, URL_COLUMNS)
            if url_col and old_url:
                old_map = {}
                for _, row in old.iterrows():
                    url = cell_text(row.get(old_url))
                    if not url:
                        continue
                    old_map[url] = {
                        col: cell_text(row.get(col))
                        for col in AUTHOR_ORG_COLS
                        if col in old.columns
                    }
                for idx, row in df.iterrows():
                    url = cell_text(row.get(url_col))
                    prev = old_map.get(url)
                    if not prev:
                        continue
                    for col, val in prev.items():
                        if val:
                            df.at[idx, col] = val
            print(f'  已载入题录副本并合并已抓字段: {dest}')
        except Exception as e:
            print(f'  读取已有题录副本失败，将重建: {e}')
    save_working_excel(dest, df)
    print(f'  题录副本: {dest}')
    return df


def format_author_org(meta: dict) -> dict:
    authors = meta.get('authors') or []
    orgs = meta.get('orgs') or {}
    marks = []
    pairs = []
    for item in authors:
        name = (item.get('name') or '').strip()
        if not name:
            continue
        nums = [str(n) for n in (item.get('orgs') or []) if str(n).strip()]
        if nums:
            marks.append(f'{name}[{",".join(nums)}]')
        else:
            marks.append(name)
        names_org = []
        for n in nums:
            org = (orgs.get(n) or orgs.get(str(n)) or '').strip()
            if org:
                names_org.append(org)
        if names_org:
            pairs.append(f'{name}：{"；".join(names_org)}')
        else:
            pairs.append(name)
    org_items = []
    keys = sorted(orgs.keys(), key=lambda x: int(x) if str(x).isdigit() else str(x))
    for key in keys:
        name = (orgs.get(key) or '').strip()
        if name:
            org_items.append(f'{key}.{name}')
    return {
        COL_AUTHOR_MARK: '; '.join(marks),
        COL_ORG_LIST: '; '.join(org_items),
        COL_AUTHOR_ORG: ' | '.join(pairs),
    }


def apply_author_org(df: pd.DataFrame, df_index, meta: dict) -> bool:
    formatted = format_author_org(meta)
    if not any(formatted.values()):
        return False
    for col, val in formatted.items():
        if val:
            df.at[df_index, col] = val
    return True


def apply_skip_mark(df: pd.DataFrame, df_index, mark: str) -> None:
    """在作者/单位列写入跳过标记，下次启动不再打开该条。"""
    mark = (mark or '').strip()
    if not mark:
        return
    for col in AUTHOR_ORG_COLS:
        df.at[df_index, col] = mark


def append_skip_mark(df: pd.DataFrame, df_index, mark: str) -> None:
    """在已有作者单位后追加跳过标记（如机构未购入）。"""
    mark = (mark or '').strip()
    if not mark:
        return
    cur = cell_text(df.at[df_index, COL_AUTHOR_ORG] if COL_AUTHOR_ORG in df.columns else '')
    if mark in cur:
        return
    if cur:
        df.at[df_index, COL_AUTHOR_ORG] = f'{cur} | {mark}'
    else:
        apply_skip_mark(df, df_index, mark)


def meta_author_org_counts(meta: dict) -> tuple[int, int]:
    authors = [
        a for a in (meta.get('authors') or [])
        if (a.get('name') or '').strip()
    ]
    orgs = meta.get('orgs') or {}
    org_n = 0
    if isinstance(orgs, dict):
        org_n = len([v for v in orgs.values() if str(v).strip()])
    return len(authors), org_n


def relative_category(excel_path: Path, records_root: Path) -> Path:
    try:
        rel = excel_path.parent.resolve().relative_to(records_root.resolve())
        return rel
    except ValueError:
        return Path(excel_path.parent.name)


# ---------------------------------------------------------------------------
# 浏览器
# ---------------------------------------------------------------------------
def _next_debug_port() -> int:
    global _debug_port_seq
    _debug_port_seq += 1
    port = CHROME_DEBUG_PORT_BASE + (os.getpid() % 80) + _debug_port_seq * 3
    if port > 60999:
        _debug_port_seq = 1
        port = CHROME_DEBUG_PORT_BASE + 11
    return port


def _force_quit_driver(driver) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except Exception:
        pass


def _driver_chromedriver_pid(driver) -> int | None:
    if driver is None:
        return None
    pid = getattr(driver, '_cnki_chromedriver_pid', None)
    if pid:
        return int(pid)
    try:
        svc = getattr(driver, 'service', None)
        proc = getattr(svc, 'process', None) if svc else None
        if proc is not None and getattr(proc, 'pid', None):
            return int(proc.pid)
    except Exception:
        pass
    return None


def _kill_pid_tree(pid: int) -> None:
    if not pid:
        return
    try:
        if sys.platform == 'win32':
            subprocess.run(
                ['taskkill', '/F', '/T', '/PID', str(pid)],
                capture_output=True, text=True, timeout=30,
                check=False)
        else:
            os.kill(pid, 9)
    except Exception as e:
        print(f'结束进程树 PID={pid} 失败: {e}')


def _kill_chrome_by_profile(profile_dir: str | None) -> None:
    """按 user-data-dir 标记杀掉残留 chrome.exe（Windows）。"""
    if not profile_dir or sys.platform != 'win32':
        return
    marker = Path(profile_dir).name
    if not marker or len(marker) < 8:
        return
    try:
        # WMIC 按命令行匹配临时 profile 名，避免误杀其他 Chrome
        like = marker.replace("'", "")
        out = subprocess.run(
            [
                'wmic', 'process', 'where',
                f"name='chrome.exe' and CommandLine like '%{like}%'",
                'get', 'ProcessId',
            ],
            capture_output=True, text=True, timeout=25, check=False)
        for line in (out.stdout or '').splitlines():
            line = line.strip()
            if line.isdigit():
                _kill_pid_tree(int(line))
    except Exception as e:
        print('按 profile 清理 chrome 残留失败:', e)


def _force_kill_browser(driver=None) -> None:
    """quit → 杀 chromedriver/chrome 进程树 → 清 profile。限流等待前必须先调用。"""
    profile = _chrome_profile_dir
    pids = set()
    pid = _driver_chromedriver_pid(driver)
    if pid:
        pids.add(pid)
    _force_quit_driver(driver)
    time.sleep(0.4)
    for p in list(pids):
        _kill_pid_tree(p)
    _kill_chrome_by_profile(profile)
    cleanup_chrome_profile()
    gc.collect()


def create_driver(download_dir: Path):
    global _chrome_profile_dir
    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    abs_dir = str(download_dir.resolve())

    CHROME_PROFILE_ROOT.mkdir(parents=True, exist_ok=True)
    _chrome_profile_dir = tempfile.mkdtemp(
        prefix=f'cnki_excel_{os.getpid()}_', dir=str(CHROME_PROFILE_ROOT))
    debug_port = _next_debug_port()
    cache_dir = os.path.join(_chrome_profile_dir, 'Cache')
    os.makedirs(cache_dir, exist_ok=True)

    chrome_options = Options()
    prefs = {
        'profile.default_content_settings.popups': 0,
        'download.default_directory': abs_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'plugins.always_open_pdf_externally': True,
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument(f'--user-data-dir={_chrome_profile_dir}')
    chrome_options.add_argument(f'--disk-cache-dir={cache_dir}')
    chrome_options.add_argument('--disk-cache-size=1')
    chrome_options.add_argument('--disable-gpu-shader-disk-cache')
    chrome_options.add_argument(f'--remote-debugging-port={debug_port}')
    chrome_options.add_argument('--no-first-run')
    chrome_options.add_argument('--no-default-browser-check')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--enable-unsafe-swiftshader')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_argument(
        '--disable-features=ScreenAI,PdfOcr,LiveCaption,ReadAnything')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    print(f'Chrome profile: {_chrome_profile_dir} | debug-port={debug_port}')
    driver = webdriver.Chrome(options=chrome_options)
    try:
        svc = getattr(driver, 'service', None)
        proc = getattr(svc, 'process', None) if svc else None
        if proc is not None and getattr(proc, 'pid', None):
            driver._cnki_chromedriver_pid = int(proc.pid)
    except Exception:
        pass
    driver.set_page_load_timeout(PAGE_TIMEOUT)
    try:
        driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
        })
    except Exception:
        pass
    set_chrome_download_dir(driver, abs_dir)
    return driver


def cleanup_chrome_profile():
    global _chrome_profile_dir
    if not _chrome_profile_dir:
        return
    path = _chrome_profile_dir
    _chrome_profile_dir = None
    try:
        shutil.rmtree(path, ignore_errors=True)
        print(f'已清理 Chrome profile: {path}')
    except Exception as e:
        print('清理 Chrome profile 失败:', e)


def recreate_driver(download_dir: Path, old_driver=None, max_tries: int = 3):
    """关掉旧会话后换新端口启动。失败返回 None，不向外抛，避免重启风暴。"""
    _force_kill_browser(old_driver)
    last_err = None
    for attempt in range(1, max_tries + 1):
        time.sleep(1.2 * attempt)
        try:
            driver = create_driver(download_dir)
            prepare_session(driver, download_dir)
            return driver
        except Exception as e:
            last_err = e
            print(f'启动 Chrome 失败（{attempt}/{max_tries}）: {e}')
            _force_kill_browser(None)
    print(f'⚠ 连续 {max_tries} 次无法启动 Chrome，暂停本篇/本表后续重试')
    if last_err:
        print(f'  最后错误: {last_err}')
    return None


def restart_browser_after_rate_limit(download_dir: Path, old_driver, n: int, total: int):
    """限流后强制关掉浏览器进程，再等，换新 profile 重新登录。"""
    wait_s = int(TENCENT_RATE_LIMIT_SLEEP)
    print(
        f'    腾讯限流：强制关闭浏览器进程，等待 {wait_s // 60} 分钟后重新认证'
        f'（{n}/{total}）'
    )
    # 必须先杀干净再 sleep，避免等待期间旧 Chrome 占内存/端口
    _force_kill_browser(old_driver)
    print('    旧浏览器进程已结束，开始等待…')
    time.sleep(wait_s)
    driver = recreate_driver(download_dir, old_driver=None)
    if driver is None:
        print('    重新启动浏览器失败')
    return driver


def _prune_captcha_debug(debug_dir, keep_latest: int = 3) -> None:
    """只保留最近若干次验证码调试图，释放磁盘与句柄占用。"""
    root = Path(debug_dir)
    if not root.is_dir():
        return
    try:
        attempts = sorted(
            [p for p in root.glob('attempt_*') if p.is_dir()],
            key=lambda p: p.stat().st_mtime,
            reverse=True)
        for old in attempts[keep_latest:]:
            shutil.rmtree(old, ignore_errors=True)
        # 清理过大的散落 png（保留最近 20 个）
        pngs = sorted(
            [p for p in root.rglob('*.png') if p.is_file()],
            key=lambda p: p.stat().st_mtime,
            reverse=True)
        for old in pngs[20:]:
            try:
                old.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception as e:
        print('清理 captcha_debug 失败:', e)


def close_other_windows_keep(driver, keep_handle):
    if not keep_handle:
        return
    try:
        for h in list(driver.window_handles):
            if h != keep_handle:
                driver.switch_to.window(h)
                driver.close()
        remaining = driver.window_handles
        if remaining:
            if keep_handle in remaining:
                driver.switch_to.window(keep_handle)
            else:
                driver.switch_to.window(remaining[0])
    except Exception:
        pass


def ensure_session_window(driver) -> bool:
    if driver is None:
        return False
    try:
        handles = driver.window_handles
    except Exception:
        return False
    if not handles:
        return False
    try:
        driver.switch_to.window(handles[0])
    except Exception:
        return False
    return True


def ensure_single_main_tab(driver) -> bool:
    if driver is None:
        return False
    try:
        handles = driver.window_handles
    except Exception:
        return False
    if len(handles) <= 1:
        return ensure_session_window(driver)
    close_other_windows_keep(driver, handles[0])
    return ensure_session_window(driver)


def wait_for_new_window(driver, before_handles, timeout=NEW_WINDOW_WAIT) -> bool:
    before_handles = before_handles or []
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            handles = driver.window_handles
        except Exception:
            return False
        if len(handles) > len(before_handles):
            return True
        time.sleep(0.25)
    return False


UNAVAILABLE_URL_MARKS = (
    'htmlread.html',
    'settlementhtml',
    'audownload.html',
)
UNAVAILABLE_TEXT_MARKS = (
    '产品不在有效期范围之内',
    '不在有效期范围',
    '不在您的订购范围',
    '未订购该',
    '无权下载',
    '请使用个人账号下载阅读',
)


def is_verify_page(driver) -> bool:
    """下载站 / 检索站的滑块或点选验证页。"""
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        url = ''
    inner = (unwrap_webvpn_url(url) or '').lower()
    return (
        'bar.cnki.net/bar/verify' in inner
        or 'kns.cnki.net/verify' in inner
        or '/bar/verify' in url
        or '/verify/index' in url
        or '/verify/home' in url
    )


def is_cnki_404_page(driver) -> bool:
    """详情页失效：地址含 /404，或标题/正文为找不到页面。"""
    try:
        url = driver.current_url or ''
    except Exception:
        url = ''
    if is_cnki_404_url(url):
        return True
    try:
        title = driver.title or ''
    except Exception:
        title = ''
    t = title.lower()
    if '404' in t or 'not found' in t or '页面不存在' in title or '找不到' in title:
        return True
    try:
        body = (driver.find_element(By.TAG_NAME, 'body').text or '')[:1200]
    except Exception:
        body = ''
    if not body:
        return False
    if '页面不存在' in body or '您访问的页面不存在' in body:
        return True
    if re.search(r'\b404\b', body) and (
            '找不到' in body or '不存在' in body or 'Not Found' in body):
        return True
    return False


def is_article_detail_page(driver) -> bool:
    """知网论文详情（摘要）页，不是结算/未购入提示页。"""
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        url = ''
    inner = (unwrap_webvpn_url(url) or url).lower()
    if '/kcms2/article/abstract' in inner or '/kcms/detail' in inner:
        return True
    try:
        if driver.find_elements(By.CSS_SELECTOR, 'div.wx-tit, #authorpart, #pdfDown'):
            return True
    except Exception:
        pass
    return False


def is_catalog_only_page(driver) -> bool:
    """详情页标题带「题录」类型，通常无全文 PDF。"""
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, 'h1 span.type, .wx-tit span.type'):
            text = (el.text or '').replace(' ', '').strip()
            if text != '题录':
                continue
            try:
                if el.is_displayed():
                    return True
            except Exception:
                return True
    except Exception:
        pass
    try:
        h1 = driver.find_elements(By.CSS_SELECTOR, 'h1')
        if h1 and re.search(r'题录', (h1[0].text or '')):
            if not driver.find_elements(By.ID, 'pdfDown'):
                return True
    except Exception:
        pass
    return False


def is_unavailable_page(driver) -> bool:
    """机构未购入 / 不在有效期 的结算或提示页。验证码、404、正文详情页不算。"""
    if is_verify_page(driver) or is_cnki_404_page(driver):
        return False
    try:
        url = (driver.current_url or '').lower()
    except Exception:
        url = ''
    inner = (unwrap_webvpn_url(url) or url).lower()
    if any(m in url or m in inner for m in UNAVAILABLE_URL_MARKS):
        return True
    try:
        tips = driver.find_elements(
            By.CSS_SELECTOR, '.organizationTip, .tip-body.o-style2')
        for el in tips:
            text = (el.text or '')
            if any(m in text for m in UNAVAILABLE_TEXT_MARKS):
                return True
    except Exception:
        pass
    if is_article_detail_page(driver):
        return False
    try:
        body = driver.find_element(By.TAG_NAME, 'body').text or ''
    except Exception:
        body = ''
    return any(m in body for m in UNAVAILABLE_TEXT_MARKS)


def find_verify_anywhere(driver) -> bool:
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    found = False
    for handle in list(driver.window_handles):
        try:
            driver.switch_to.window(handle)
            if is_verify_page(driver):
                found = True
                break
        except Exception:
            continue
    if current and current in driver.window_handles:
        try:
            driver.switch_to.window(current)
        except Exception:
            pass
    return found


def has_visible_captcha_anywhere(driver) -> bool:
    """任意标签页是否有可交互验证码控件（已通过的验证页不算）。"""
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    found = False
    for handle in list(driver.window_handles):
        try:
            driver.switch_to.window(handle)
            if has_visible_captcha_widget(driver):
                found = True
                break
        except Exception:
            continue
    if current and current in driver.window_handles:
        try:
            driver.switch_to.window(current)
        except Exception:
            pass
    return found


def find_unavailable_anywhere(driver) -> bool:
    current = None
    try:
        current = driver.current_window_handle
    except Exception:
        pass
    found = False
    for handle in list(driver.window_handles):
        try:
            driver.switch_to.window(handle)
            if is_unavailable_page(driver):
                found = True
                break
        except Exception:
            continue
    if current and current in driver.window_handles:
        try:
            driver.switch_to.window(current)
        except Exception:
            pass
    return found


def close_unavailable_and_extras(driver, keep_handle):
    """关闭结算/未购入提示页及其它多余标签，回到正文。"""
    for handle in list(driver.window_handles):
        if keep_handle and handle == keep_handle:
            continue
        try:
            driver.switch_to.window(handle)
            driver.close()
        except Exception:
            pass
    if keep_handle and keep_handle in driver.window_handles:
        try:
            driver.switch_to.window(keep_handle)
            return
        except Exception:
            pass
    ensure_single_main_tab(driver)


def snapshot_download_files(download_dir: Path) -> set[str]:
    if not download_dir.is_dir():
        return set()
    return {
        name for name in os.listdir(download_dir)
        if not name.endswith('.lock') and not name.endswith('.done')
    }


def _is_pdf_file(path: Path, min_size=1024) -> bool:
    try:
        if not path.is_file() or path.stat().st_size < min_size:
            return False
        if path.name.lower().endswith('.pdf'):
            return True
        with open(path, 'rb') as f:
            return f.read(5).startswith(b'%PDF')
    except OSError:
        return False


def _still_downloading(path: Path) -> bool:
    name = path.name.lower()
    if name.endswith('.crdownload') or name.endswith('.tmp'):
        return True
    return path.with_name(path.name + '.crdownload').exists()


def iter_finished_pdfs(download_dir: Path):
    if not download_dir.is_dir():
        return
    for name in os.listdir(download_dir):
        path = download_dir / name
        if _still_downloading(path):
            continue
        if _is_pdf_file(path):
            yield path


def find_new_pdf(
    download_dir: Path, before: set[str], since_ts: float | None = None,
    dest_pdf: Path | None = None,
) -> Path | None:
    """找出点击下载后新出现（或刚写完）的 PDF，不因目录里其它 .crdownload 误判。"""
    if dest_pdf and dest_pdf.is_file() and dest_pdf.stat().st_size > 1024:
        return dest_pdf
    found = []
    stem = (dest_pdf.stem[:24] if dest_pdf and dest_pdf.stem else '')
    for path in iter_finished_pdfs(download_dir):
        name = path.name
        try:
            mtime = path.stat().st_mtime
            size = path.stat().st_size
        except OSError:
            continue
        is_new = name not in (before or set())
        is_recent = since_ts is not None and mtime >= since_ts - 2 and size > 1024
        if not is_new and not is_recent:
            continue
        prio = 0 if (stem and stem in path.stem) else 1
        found.append((prio, -mtime, path))
    if not found:
        return None
    found.sort()
    return found[0][2]


def paper_pdf_exists(download_dir: Path, dest_pdf: Path) -> Path | None:
    """目标文件名已存在，或目录中已有同题名 PDF（上次未改名成功）。"""
    if dest_pdf.is_file() and dest_pdf.stat().st_size > 1024:
        return dest_pdf
    stem = dest_pdf.stem
    if len(stem) < 12:
        return None
    key = stem[:30]
    for path in iter_finished_pdfs(download_dir):
        if key in path.stem and path.stat().st_size > 1024:
            return path
    return None


def wait_for_new_pdf(
    download_dir: Path, before: set[str], timeout=PDF_WAIT,
    since_ts: float | None = None, dest_pdf: Path | None = None,
) -> Path | None:
    deadline = time.time() + timeout
    last_size = -1
    stable = 0
    candidate = None
    while time.time() < deadline:
        newest = find_new_pdf(download_dir, before, since_ts=since_ts, dest_pdf=dest_pdf)
        if newest:
            try:
                size = newest.stat().st_size
            except OSError:
                size = 0
            if size > 1024 and size == last_size:
                stable += 1
                if stable >= 2:
                    return newest
            else:
                stable = 0
                last_size = size
            candidate = newest
        time.sleep(0.4)
    if candidate and _is_pdf_file(candidate):
        return candidate
    return find_new_pdf(download_dir, before, since_ts=since_ts, dest_pdf=dest_pdf)


def wait_pdf_or_unavailable(
    driver, download_dir: Path, before: set[str], timeout=PDF_WAIT,
    save_dir=None, main_handle=None, since_ts=None, dest_pdf=None,
) -> tuple[str, Path | None]:
    """等待 PDF 落地；遇验证码先过码，遇机构未购入则返回。"""
    deadline = time.time() + timeout
    last_size = -1
    stable = 0
    candidate = None
    while time.time() < deadline:
        newest = find_new_pdf(
            download_dir, before, since_ts=since_ts, dest_pdf=dest_pdf)
        if newest:
            try:
                size = newest.stat().st_size
            except OSError:
                size = 0
            if size > 1024 and size == last_size:
                stable += 1
                if stable >= 2:
                    return 'ok', newest
            else:
                stable = 0
                last_size = size
            candidate = newest
            time.sleep(0.4)
            continue
        if has_visible_captcha_anywhere(driver):
            print('    下载过程出现验证码，开始处理...')
            if not save_dir:
                return 'captcha', None
            if not ensure_captcha_passed(
                driver, str(save_dir), main_handle,
                download_dir=str(download_dir), before_files=before,
            ):
                return 'captcha_fail', None
            time.sleep(0.8)
            continue
        if find_unavailable_anywhere(driver):
            return 'unavailable', None
        time.sleep(0.4)
    if candidate and _is_pdf_file(candidate):
        return 'ok', candidate
    leftover = find_new_pdf(
        download_dir, before, since_ts=since_ts, dest_pdf=dest_pdf)
    if leftover:
        print(f'    等待结束时发现已下载文件: {leftover.name}')
        return 'ok', leftover
    new_names = sorted(snapshot_download_files(download_dir) - (before or set()))
    if new_names:
        print(f'    下载目录新文件但未识别为 PDF: {new_names[:8]}')
    return 'timeout', None


def rename_pdf(src: Path, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.resolve() == src.resolve():
        return dest
    if dest.exists():
        stem, suffix = dest.stem, dest.suffix
        i = 2
        while True:
            alt = dest.with_name(f'{stem}_{i}{suffix}')
            if not alt.exists():
                dest = alt
                break
            i += 1
    try:
        src.replace(dest)
        return dest
    except OSError:
        shutil.copy2(src, dest)
        try:
            src.unlink()
        except OSError:
            pass
        return dest


def pass_captcha_and_login(
    driver, save_dir: Path, main_handle=None,
    download_dir=None, before_files=None,
) -> bool:
    save_dir = Path(save_dir)
    save_dir.mkdir(parents=True, exist_ok=True)
    from cnkiforexcel import captcha_needs_browser_restart
    cap0 = ensure_cnki_captcha_after_nav(driver, str(save_dir), main_handle)
    if captcha_needs_browser_restart(cap0) or not cap0:
        return False
    if has_visible_captcha_anywhere(driver):
        dl = str(download_dir) if download_dir else None
        cap = ensure_captcha_passed(
            driver, str(save_dir), main_handle,
            download_dir=dl, before_files=before_files,
        )
        if captcha_needs_browser_restart(cap) or not cap:
            return False
    if not ensure_unit_login_if_needed(driver, restore_handle=main_handle):
        time.sleep(1)
        if not ensure_unit_login_if_needed(driver, restore_handle=main_handle):
            return False
    cap1 = ensure_cnki_captcha_after_nav(driver, str(save_dir), main_handle)
    if captcha_needs_browser_restart(cap1) or not cap1:
        return False
    return True


def prepare_session(driver, save_dir: Path) -> None:
    driver.get(get_login_home_url())
    time.sleep(2)
    try:
        driver.maximize_window()
    except Exception:
        pass
    main_handle = driver.current_window_handle
    pass_captcha_and_login(driver, save_dir, main_handle)
    print('浏览器会话已就绪（验证码/机构登录）')


def has_pdf_download_button(driver) -> bool:
    """页面上是否有可点的 PDF 下载（不含仅「下载/CAJ/HTML」按钮）。"""
    try:
        for el in driver.find_elements(By.ID, 'pdfDown'):
            try:
                if el.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR, 'a#pdfDown, li#pdfDown, [id*="pdfDown"]'):
            try:
                if el.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    try:
        for btn in driver.find_elements(By.CLASS_NAME, 'operate-btn'):
            try:
                if not btn.is_displayed():
                    continue
            except Exception:
                continue
            txt = (btn.text or '') + ' ' + (btn.get_attribute('id') or '')
            if re.search(r'pdf', txt, re.I):
                return True
    except Exception:
        pass
    return False


def click_pdf_button(driver) -> bool:
    try:
        pdf_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, 'pdfDown'))
        )
        driver.execute_script('arguments[0].click();', pdf_btn)
        return True
    except Exception:
        pass
    try:
        for btn in driver.find_elements(By.CLASS_NAME, 'operate-btn'):
            txt = (btn.text or '') + ' ' + (btn.get_attribute('id') or '')
            if re.search(r'pdf', txt, re.I):
                driver.execute_script('arguments[0].click();', btn)
                return True
    except Exception:
        pass
    try:
        links = driver.find_elements(
            By.CSS_SELECTOR, 'a#pdfDown, a[id*="pdfDown"], [id="pdfDown"]')
        for a in links:
            driver.execute_script('arguments[0].click();', a)
            return True
    except Exception:
        pass
    return False


EXTRACT_AUTHOR_ORG_JS = r"""
var result = {authors: [], orgs: {}};
function clean(s) {
    return (s || '').replace(/\s+/g, ' ').trim();
}
function cleanName(s) {
    return clean(s).replace(/[\s;；,，]+$/g, '');
}
function childSpans(root) {
    var out = [];
    if (!root) return out;
    for (var i = 0; i < root.children.length; i++) {
        if ((root.children[i].tagName || '').toLowerCase() === 'span') {
            out.push(root.children[i]);
        }
    }
    if (!out.length) out = Array.prototype.slice.call(root.querySelectorAll('span'));
    return out;
}
function addOrg(key, name) {
    key = String(key || '').trim();
    name = clean(name);
    if (!key || !name) return;
    if (!result.orgs[key]) result.orgs[key] = name;
}
function parseOrgText(text) {
    text = (text || '').replace(/\r/g, '');
    var re = /(\d+)\s*[.．、]\s*([^\d\n]+?)(?=(?:\n|\d+\s*[.．、]|$))/g;
    var m;
    while ((m = re.exec(text))) {
        addOrg(m[1], m[2]);
    }
    var lines = text.split(/\n+/);
    for (var i = 0; i < lines.length; i++) {
        var mm = clean(lines[i]).match(/^(\d+)\s*[.．、]\s*(.+)$/);
        if (mm) addOrg(mm[1], mm[2]);
    }
}
var authorRoot = document.querySelector('#authorpart');
if (!authorRoot) {
    var heads = document.querySelectorAll('div.wx-tit h3.author, h3.author');
    if (heads.length) authorRoot = heads[0];
}
if (authorRoot) {
    childSpans(authorRoot).forEach(function(span) {
        var tip = span.querySelector('.authortip, .icon-email');
        var sup = span.querySelector('sup');
        var nameEl = span.querySelector('a') || span;
        var name = cleanName(nameEl.textContent || '');
        if (sup) {
            var supText = clean(sup.textContent);
            if (supText) name = cleanName(name.replace(supText, ''));
        }
        if (tip) {
            var tipText = clean(tip.textContent);
            if (tipText) name = cleanName(name.replace(tipText, ''));
        }
        name = name.replace(/\d+(,\d+)*/g, '').trim();
        if (!name) return;
        var nums = [];
        if (sup) {
            var mm = String(sup.textContent || '').match(/\d+/g);
            if (mm) nums = mm;
        }
        result.authors.push({name: name, orgs: nums});
    });
}
var wx = document.querySelector('div.wx-tit');
var orgRoots = [];
if (wx) {
    orgRoots = orgRoots.concat(Array.prototype.slice.call(
        wx.querySelectorAll('h3, p, .orgn, .org, .author')));
} else {
    orgRoots = Array.prototype.slice.call(
        document.querySelectorAll('h3.author, h3, .orgn, .org'));
}
orgRoots.forEach(function(node) {
    if (!node || node.id === 'authorpart') return;
    if (authorRoot && node === authorRoot) return;
    childSpans(node).forEach(function(span) {
        var text = clean(span.textContent || '');
        var m = text.match(/^(\d+)\s*[.．、]\s*(.+)$/);
        if (m) addOrg(m[1], m[2]);
    });
    parseOrgText(node.innerText || node.textContent || '');
});
if (wx && Object.keys(result.orgs).length === 0) {
    parseOrgText(wx.innerText || '');
}
return result;
"""


def wait_detail_ready(driver, timeout=12) -> bool:
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.ID, 'authorpart')
                or d.find_elements(By.CSS_SELECTOR, 'div.wx-tit')
                or d.find_elements(By.ID, 'pdfDown')
            )
        )
        return True
    except Exception:
        return False


def extract_author_org(driver) -> dict:
    """从知网详情页 wx-tit 提取作者上标与单位编号对应关系。"""
    try:
        data = driver.execute_script(EXTRACT_AUTHOR_ORG_JS) or {}
    except Exception as e:
        print(f'    提取作者/单位失败: {e}')
        return {}
    authors = []
    for item in data.get('authors') or []:
        name = cell_text(item.get('name'))
        if not name:
            continue
        nums = [str(n) for n in (item.get('orgs') or []) if str(n).strip()]
        authors.append({'name': name, 'orgs': nums})
    orgs = {}
    raw_orgs = data.get('orgs') or {}
    if isinstance(raw_orgs, dict):
        for key, val in raw_orgs.items():
            text = cell_text(val)
            if text:
                orgs[str(key)] = text
    return {'authors': authors, 'orgs': orgs}


def download_paper(
    driver,
    url: str,
    dest_pdf: Path,
    download_dir: Path,
    debug_dir: Path,
    need_pdf: bool = True,
    need_meta: bool = True,
) -> tuple[str, dict]:
    """下载一篇。返回 (ok|unavailable|not_found|no_meta|fail, 作者单位信息)。

    unavailable：机构未购入/不在有效期，已抓作者单位，应关页并继续下一条。
    not_found：网址 404/失效。
    no_meta：详情页无作者也无单位（仍会尝试下载 PDF）。
    rate_limit：腾讯滑块过于频繁，需关浏览器后重新认证。
    单作者/单单位仍下载 PDF；need_meta=False 时不再抓作者单位。
    """
    dest_pdf = Path(dest_pdf)
    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    meta = {}
    paper_before = None
    click_time = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            if not ensure_session_window(driver):
                print('    无可用浏览器窗口')
                return 'fail', meta
            driver.set_page_load_timeout(PAGE_TIMEOUT)
            open_cnki_url(driver, url)
            time.sleep(random.uniform(0.8, 1.5))
            try:
                print(f'    当前页: {driver.current_url}')
            except Exception:
                pass
            if is_cnki_404_page(driver):
                print('    网址失效 404，标记后跳过')
                ensure_single_main_tab(driver)
                return 'not_found', meta
            main_handle = driver.window_handles[0] if driver.window_handles else None

            if not pass_captcha_and_login(driver, debug_dir, main_handle):
                print(f'    验证码/登录失败（尝试 {attempt}/{MAX_RETRY}）')
                if tencent_last_fail() == 'rate_limit':
                    ensure_single_main_tab(driver)
                    return 'rate_limit', meta
                ensure_single_main_tab(driver)
                time.sleep(2)
                continue

            if is_cnki_404_page(driver):
                print('    网址失效 404，标记后跳过')
                ensure_single_main_tab(driver)
                return 'not_found', meta

            main_handle = driver.current_window_handle
            wait_detail_ready(driver)
            if need_meta:
                got = extract_author_org(driver)
                if got.get('authors') or got.get('orgs'):
                    meta = got
                    print(
                        f'    作者 {len(meta.get("authors") or [])} / '
                        f'单位 {len(meta.get("orgs") or {})}'
                    )
            else:
                print('    作者单位已有，跳过抓取')

            if is_cnki_404_page(driver):
                print('    网址失效 404，标记后跳过')
                ensure_single_main_tab(driver)
                return 'not_found', meta

            if need_meta:
                n_auth, n_org = meta_author_org_counts(meta)
                if n_auth == 0 and n_org == 0:
                    meta = {'_no_author_org': True, 'authors': [], 'orgs': {}}
                    if not need_pdf:
                        print('    无作者/单位，标记后跳过抓取')
                        if main_handle in driver.window_handles:
                            close_other_windows_keep(driver, main_handle)
                        else:
                            ensure_single_main_tab(driver)
                        return 'no_meta', meta
                    print('    无作者/单位，继续下载 PDF')

            if find_unavailable_anywhere(driver):
                print('    机构未购入/不在有效期，记录作者单位后跳过 PDF')
                close_unavailable_and_extras(driver, main_handle)
                return 'unavailable', meta

            if not need_pdf:
                if main_handle in driver.window_handles:
                    close_other_windows_keep(driver, main_handle)
                else:
                    ensure_single_main_tab(driver)
                return 'ok', meta

            if not has_pdf_download_button(driver):
                if is_catalog_only_page(driver):
                    print('    该条为仅题录，无 PDF，标记后跳过')
                    status = 'catalog_only'
                else:
                    print('    无 PDF 下载按钮，标记后跳过，不再重试')
                    status = 'no_pdf'
                if main_handle in driver.window_handles:
                    close_other_windows_keep(driver, main_handle)
                else:
                    ensure_single_main_tab(driver)
                return status, meta

            if paper_before is None:
                paper_before = snapshot_download_files(download_dir)
                click_time = time.time()
            before = paper_before
            before_handles = list(driver.window_handles)
            if not click_pdf_button(driver):
                try:
                    now_u = driver.current_url or ''
                except Exception:
                    now_u = ''
                print(
                    f'    未找到 PDF 下载按钮（尝试 {attempt}/{MAX_RETRY}）'
                    f' 当前页: {now_u}'
                )
                if is_cnki_404_page(driver):
                    print('    网址失效 404，标记后跳过')
                    ensure_single_main_tab(driver)
                    return 'not_found', meta
                if is_catalog_only_page(driver):
                    print('    未找到 PDF 且为仅题录，标记后跳过，不再重试')
                    ensure_single_main_tab(driver)
                    return 'catalog_only', meta
                print('    无 PDF 下载按钮，标记后跳过，不再重试')
                ensure_single_main_tab(driver)
                return 'no_pdf', meta

            wait_for_new_window(driver, before_handles, timeout=8)
            time.sleep(1.0)
            if has_visible_captcha_anywhere(driver) or has_unit_login_anywhere(driver):
                if has_visible_captcha_anywhere(driver):
                    print('    检测到下载验证码页，开始处理...')
                if not pass_captcha_and_login(
                    driver, debug_dir, main_handle,
                    download_dir=download_dir, before_files=before,
                ):
                    print(f'    下载后验证码/登录失败（尝试 {attempt}/{MAX_RETRY}）')
                    if tencent_last_fail() == 'rate_limit':
                        ensure_single_main_tab(driver)
                        return 'rate_limit', meta
                    ensure_single_main_tab(driver)
                    time.sleep(2)
                    continue
                time.sleep(1.0)

            if find_unavailable_anywhere(driver):
                print('    机构未购入/不在有效期，记录作者单位后跳过 PDF')
                close_unavailable_and_extras(driver, main_handle)
                return 'unavailable', meta

            outcome, pdf = wait_pdf_or_unavailable(
                driver, download_dir, before, timeout=PDF_WAIT,
                save_dir=debug_dir, main_handle=main_handle,
                since_ts=click_time, dest_pdf=dest_pdf,
            )
            if outcome == 'captcha_fail':
                print(f'    下载验证码未通过（尝试 {attempt}/{MAX_RETRY}）')
                if tencent_last_fail() == 'rate_limit':
                    ensure_single_main_tab(driver)
                    return 'rate_limit', meta
                ensure_single_main_tab(driver)
                time.sleep(2)
                continue
            if outcome == 'unavailable':
                print('    机构未购入/不在有效期，记录作者单位后跳过 PDF')
                close_unavailable_and_extras(driver, main_handle)
                return 'unavailable', meta
            if outcome == 'ok' and pdf:
                saved = rename_pdf(pdf, dest_pdf)
                print(f'    已保存: {saved.name}')
                if main_handle in driver.window_handles:
                    close_other_windows_keep(driver, main_handle)
                else:
                    ensure_single_main_tab(driver)
                return 'ok', meta

            print(f'    等待 PDF 超时（尝试 {attempt}/{MAX_RETRY}）')
            if not has_pdf_download_button(driver):
                if is_catalog_only_page(driver):
                    print('    等待超时且无 PDF 按钮（仅题录），标记后跳过')
                    ensure_single_main_tab(driver)
                    return 'catalog_only', meta
                print('    等待超时且无 PDF 按钮，标记后跳过，不再重试')
                ensure_single_main_tab(driver)
                return 'no_pdf', meta
            pdf = find_new_pdf(
                download_dir, paper_before or set(),
                since_ts=click_time, dest_pdf=dest_pdf)
            if pdf:
                saved = rename_pdf(pdf, dest_pdf)
                print(f'    补识别到已下载 PDF: {saved.name}')
                if main_handle in driver.window_handles:
                    close_other_windows_keep(driver, main_handle)
                else:
                    ensure_single_main_tab(driver)
                return 'ok', meta
            ensure_single_main_tab(driver)
            time.sleep(2)
        except TimeoutException:
            print(f'    页面加载超时（尝试 {attempt}/{MAX_RETRY}）')
        except Exception as e:
            if is_session_dead_error(e):
                print(f'    浏览器已断开（尝试 {attempt}/{MAX_RETRY}）')
                return 'session_dead', meta
            if isinstance(e, WebDriverException):
                err = str(e).lower()
                print(f'    会话错误（尝试 {attempt}/{MAX_RETRY}）: {e}')
                if 'no such window' in err or 'web view not found' in err:
                    ensure_single_main_tab(driver)
                    if not ensure_session_window(driver):
                        return 'session_dead', meta
                time.sleep(2)
                continue
            print(f'    未知错误（尝试 {attempt}/{MAX_RETRY}）: {e}')
    return 'fail', meta


# ---------------------------------------------------------------------------
# 文件夹选择
# ---------------------------------------------------------------------------
def prompt_folders(root: Path) -> list[Path]:
    folders = list_category_folders(root)
    if not folders:
        print(f'未找到子文件夹: {root}')
        return []
    print('\n可选子文件夹：')
    for i, folder in enumerate(folders, 1):
        n = count_excels(folder)
        print(f'  {i:2d}. {folder.name}  ({n} 个 Excel)')
    print('输入编号或名称（逗号分隔）；回车=全部；q=退出')
    raw = input('请选择: ').strip()
    if raw.lower() in ('q', 'quit', 'exit'):
        return []
    if not raw or raw.lower() in ('all', 'a', '*'):
        return folders
    chosen = []
    tokens = [t.strip() for t in re.split(r'[,，\s]+', raw) if t.strip()]
    name_map = {f.name: f for f in folders}
    for token in tokens:
        if token.isdigit():
            idx = int(token)
            if 1 <= idx <= len(folders):
                chosen.append(folders[idx - 1])
                continue
        if token in name_map:
            chosen.append(name_map[token])
            continue
        rel = root / token
        if rel.is_dir():
            chosen.append(rel)
            continue
        print(f'  忽略无效选项: {token}')
    # 去重且保持顺序
    seen = set()
    uniq = []
    for p in chosen:
        key = str(p.resolve())
        if key not in seen:
            seen.add(key)
            uniq.append(p)
    return uniq


def resolve_folders(root: Path, names: list[str], use_all: bool) -> list[Path]:
    if names:
        folders = []
        available = {f.name: f for f in list_category_folders(root)}
        for name in names:
            name = name.strip().strip('/\\')
            if not name:
                continue
            if name in available:
                folders.append(available[name])
                continue
            path = Path(name)
            if not path.is_absolute():
                path = root / name
            if path.is_dir():
                folders.append(path)
            else:
                print(f'文件夹不存在: {name}')
        return folders
    if use_all:
        return list_category_folders(root)
    return prompt_folders(root)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def append_fail_log(log_path: Path, excel_path: Path, paper: dict, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    new_file = not log_path.is_file()
    with log_path.open('a', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(
            f, fieldnames=['time', 'excel', 'row', 'title', 'url', 'message'])
        if new_file:
            writer.writeheader()
        writer.writerow({
            'time': datetime.now().isoformat(timespec='seconds'),
            'excel': str(excel_path),
            'row': paper.get('row', ''),
            'title': paper.get('title', ''),
            'url': paper.get('url', ''),
            'message': message,
        })


def process_excel(
    driver,
    excel_path: Path,
    records_root: Path,
    pdf_root: Path,
    excel_out: Path,
    resume: bool,
    mark_done_on_fail: bool,
    limit_papers: int,
    fail_log: Path,
    paper_counter: list[int],
):
    lock = ExcelLock(excel_path)
    if resume and lock.is_done():
        print(f'  ⊘ 已完成，跳过: {excel_path.name}')
        return driver, 'skipped_done', 0, 0, 0

    if not lock.acquire():
        holder = lock.lock_pid()
        extra = f'（pid={holder}）' if holder else ''
        print(f'  ⊘ 正在被其他进程处理{extra}，跳过: {excel_path.name}')
        return driver, 'locked', 0, 0, 0

    success = fail = skip = 0
    try:
        copy_path = working_excel_path(excel_path, records_root, excel_out)
        try:
            df = prepare_working_excel(excel_path, copy_path)
            papers = load_papers_from_df(df)
        except Exception as e:
            print(f'  读取/复制失败: {e}')
            return driver, 'read_error', 0, 0, 0

        if limit_papers and limit_papers > 0:
            papers = papers[:limit_papers]
        total = len(papers)
        if total == 0:
            print('  URL 0 条')
            lock.mark_done(0, 0, 0, 0)
            print('  ✓ 无有效 URL，已标记完成')
            return driver, 'empty', 0, 0, 0

        rel = relative_category(excel_path, records_root)
        download_dir = pdf_root / rel
        download_dir.mkdir(parents=True, exist_ok=True)
        debug_dir = pdf_root / '_captcha_debug'

        pending = []
        for paper in papers:
            dest = download_dir / paper['filename']
            existing = paper_pdf_exists(download_dir, dest)
            dest_missing = (not dest.is_file()) or dest.stat().st_size <= 1024
            if (
                existing is not None
                and dest_missing
                and existing.resolve() != dest.resolve()
            ):
                try:
                    dest = rename_pdf(existing, dest)
                    print(f'  已将已下载 PDF 更名为: {dest.name}')
                except Exception:
                    dest = existing
            pdf_exists = dest.is_file() and dest.stat().st_size > 1024
            has_meta = bool(paper.get('has_meta'))
            row = df.loc[paper['df_index']]
            if row_has_skip_mark(row, SKIP_ALL_MARKS):
                continue
            if pdf_exists and has_meta:
                continue
            if not pdf_exists and row_has_skip_mark(row, SKIP_PDF_MARKS):
                continue
            need_pdf = not pdf_exists
            need_meta = not has_meta
            if not need_pdf and not need_meta:
                continue
            pending.append((paper, dest, need_pdf, need_meta, pdf_exists))
        already_ok = total - len(pending)
        print(
            f'  URL {total} 条，已完成跳过 {already_ok}，待处理 {len(pending)}'
        )
        if not pending:
            lock.mark_done(0, 0, already_ok, total)
            print('  ✓ 本表均已有 PDF 或已标记跳过，已标记完成')
            return driver, 'done', 0, 0, already_ok

        try:
            set_chrome_download_dir(driver, str(download_dir.resolve()))
        except Exception:
            pass
        skip = already_ok
        success = fail = 0

        for n, (paper, dest, need_pdf, need_meta, pdf_exists) in enumerate(pending, 1):
            display = paper['title'][:40] + ('...' if len(paper['title']) > 40 else '')
            print(f'  [{n}/{len(pending)} · 全表 {total}] {display}')
            if pdf_exists and need_meta:
                print('    PDF 已存在，仅补作者/单位')
            elif (not need_meta) and need_pdf:
                print('    作者单位已有，仅下载 PDF')
            session_retried = False
            rate_restarts = 0
            status, meta = 'fail', {}
            while True:
                try:
                    if driver is None or not is_driver_alive(driver) or not ensure_session_window(driver):
                        print('    ⚠ 浏览器失效，正在重启…')
                        driver = recreate_driver(download_dir, old_driver=driver)
                        if driver is None:
                            print('    ⚠ 无法启动浏览器，本篇失败')
                            status = 'fail'
                            break
                    status, meta = download_paper(
                        driver, paper['url'], dest, download_dir, debug_dir,
                        need_pdf=need_pdf, need_meta=need_meta)
                    if status == 'session_dead':
                        if session_retried:
                            print('    ⚠ 重启后浏览器仍不可用，本篇失败')
                            status = 'fail'
                            break
                        print('    ⚠ 浏览器断开，正在重启后重试本篇…')
                        driver = recreate_driver(download_dir, old_driver=driver)
                        if driver is None:
                            print('    ⚠ 无法启动浏览器，本篇失败')
                            status = 'fail'
                            break
                        session_retried = True
                        continue
                    if status == 'rate_limit':
                        if rate_restarts >= TENCENT_RELOGIN_RETRIES:
                            print(
                                f'    腾讯限流已重新认证 {TENCENT_RELOGIN_RETRIES} 次仍失败，本篇放弃'
                            )
                            status = 'fail'
                            break
                        rate_restarts += 1
                        nxt = restart_browser_after_rate_limit(
                            download_dir, driver, rate_restarts,
                            TENCENT_RELOGIN_RETRIES)
                        driver = nxt
                        if driver is None:
                            status = 'fail'
                            break
                        continue
                    break
                except Exception as e:
                    if is_session_dead_error(e) and not session_retried:
                        print('    ⚠ 浏览器断开，正在重启后重试本篇…')
                        driver = recreate_driver(download_dir, old_driver=driver)
                        if driver is None:
                            print('    ⚠ 无法启动浏览器，本篇失败')
                            status = 'fail'
                            break
                        session_retried = True
                        continue
                    print(f'    处理本篇异常: {e}')
                    status, meta = 'fail', {}
                    break
            if meta and meta.get('_no_author_org'):
                apply_skip_mark(df, paper['df_index'], MARK_NO_AUTHOR_ORG)
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')
            elif meta and apply_author_org(df, paper['df_index'], meta):
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')
            elif status == 'not_found':
                apply_skip_mark(df, paper['df_index'], MARK_URL_404)
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')
            elif status == 'no_meta':
                apply_skip_mark(df, paper['df_index'], MARK_NO_AUTHOR_ORG)
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')
            elif need_meta:
                print('    ! 未抓到作者/单位对应')

            if status == 'unavailable' and paper.get('has_meta'):
                append_skip_mark(df, paper['df_index'], MARK_UNAVAILABLE)
                try:
                    save_working_excel(copy_path, df)
                except Exception:
                    pass
            elif status == 'unavailable':
                apply_skip_mark(df, paper['df_index'], MARK_UNAVAILABLE)
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')

            if status == 'catalog_only' and paper.get('has_meta'):
                append_skip_mark(df, paper['df_index'], MARK_CATALOG_ONLY)
                try:
                    save_working_excel(copy_path, df)
                except Exception:
                    pass
            elif status == 'catalog_only':
                apply_skip_mark(df, paper['df_index'], MARK_CATALOG_ONLY)
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')

            if status == 'no_pdf' and paper.get('has_meta'):
                append_skip_mark(df, paper['df_index'], MARK_NO_PDF)
                try:
                    save_working_excel(copy_path, df)
                except Exception:
                    pass
            elif status == 'no_pdf':
                apply_skip_mark(df, paper['df_index'], MARK_NO_PDF)
                paper['has_meta'] = True
                try:
                    save_working_excel(copy_path, df)
                except Exception as e:
                    print(f'    写回题录副本失败: {e}')

            if status == 'not_found':
                print('    ⊘ 网址失效 404，已标记并继续')
                skip += 1
                append_fail_log(fail_log, excel_path, paper, 'url_404')
            elif status == 'no_meta':
                print('    ⊘ 无作者/单位（PDF 已有），已标记并继续')
                skip += 1
            elif status == 'unavailable':
                print('    ⊘ 机构未购入，已记作者单位并继续')
                skip += 1
                append_fail_log(
                    fail_log, excel_path, paper, 'institution_unavailable')
            elif status == 'catalog_only':
                print('    ⊘ 仅题录无 PDF，已标记并继续')
                skip += 1
            elif status == 'no_pdf':
                print('    ⊘ 无 PDF 下载按钮，已标记并继续')
                skip += 1
            elif need_pdf:
                if status == 'ok':
                    print('    ✓ PDF 成功')
                    success += 1
                else:
                    print('    ✗ PDF 失败')
                    fail += 1
                    append_fail_log(fail_log, excel_path, paper, 'download_failed')
            else:
                skip += 1
                if need_meta and not paper.get('has_meta'):
                    fail += 1
                    append_fail_log(fail_log, excel_path, paper, 'author_org_failed')

            paper_counter[0] += 1
            try:
                if driver is None or not is_driver_alive(driver) or not ensure_session_window(driver):
                    print('    ⚠ 无可用窗口，下一篇前再尝试启动')
                    nxt = recreate_driver(download_dir, old_driver=driver)
                    if nxt is not None:
                        driver = nxt
                else:
                    ensure_single_main_tab(driver)
            except Exception as e:
                print(f'    会话检查失败，继续下一篇: {e}')
                driver = None

            # 每篇后释放临时对象；每 10 篇清理验证码调试图
            meta = None
            if n % 10 == 0:
                _prune_captcha_debug(debug_dir)
            gc.collect()
            time.sleep(random.uniform(1.0, 2.5))

        leftover = 0
        for p in papers:
            dest = download_dir / p['filename']
            pdf_ok = paper_pdf_exists(download_dir, dest) is not None
            row = df.loc[p['df_index']]
            if row_has_skip_mark(row, SKIP_PDF_MARKS):
                continue
            if not pdf_ok:
                leftover += 1
        try:
            save_working_excel(copy_path, df)
        except Exception as e:
            print(f'  最终保存题录副本失败: {e}')

        can_done = (fail == 0 and leftover == 0) or mark_done_on_fail
        if can_done:
            lock.mark_done(success, fail, skip, total)
            print(
                f'  ✓ 本表完成并标记: 成功 {success} / 失败 {fail} / 跳过 {skip}'
                + (f' / 待处理 {leftover}' if leftover else '')
            )
            status = 'done'
        else:
            print(
                f'  ! 有失败或仍有未处理条目，未标记完成: '
                f'成功 {success} / 失败 {fail} / 跳过 {skip}'
                f' / 待处理 {leftover}'
            )
            status = 'partial'
        return driver, status, success, fail, skip
    finally:
        lock.release()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='从知网题录 Excel 的 URL 下载 PDF（可选子文件夹，完成标记+锁）')
    parser.add_argument(
        '--records', type=Path, default=RECORDS_ROOT,
        help='题录根目录（默认 知网数据库下载）')
    parser.add_argument(
        '--out', type=Path, default=DEFAULT_PDF_ROOT,
        help='PDF 保存根目录')
    parser.add_argument(
        '--excel-out', type=Path, default=DEFAULT_EXCEL_OUT,
        help='题录副本保存根目录（写入作者单位对应）')
    parser.add_argument(
        '--folder', action='append', default=[],
        help='只处理指定一级子文件夹，可重复；不传则交互选择')
    parser.add_argument(
        '--all', action='store_true',
        help='处理全部子文件夹（不交互）')
    parser.add_argument(
        '--limit-files', type=int, default=0,
        help='每个文件夹最多处理多少个 Excel（调试）')
    parser.add_argument(
        '--limit-papers', type=int, default=0,
        help='每个 Excel 最多下载多少篇（调试）')
    parser.add_argument(
        '--resume', dest='resume', action='store_true', default=True,
        help='跳过已有 .done 的 Excel（默认开启）')
    parser.add_argument(
        '--no-resume', dest='resume', action='store_false',
        help='忽略 .done，重新处理')
    parser.add_argument(
        '--mark-done-on-fail', action='store_true',
        help='即使有失败也写 .done（默认失败不标记，便于下次重试）')
    parser.add_argument(
        '--dry-run', action='store_true',
        help='只列出将处理的 Excel，不打开浏览器')
    parser.add_argument(
        '--list', action='store_true',
        help='列出子文件夹后退出')
    return parser.parse_args()


def main() -> int:
    _setup_stdio()
    args = parse_args()
    root = args.records
    if not root.is_dir():
        print(f'题录目录不存在: {root}')
        return 1

    if args.list:
        print(f'题录根目录: {root}')
        for folder in list_category_folders(root):
            print(f'  {folder.name}: {count_excels(folder)} 个 Excel')
        return 0

    folders = resolve_folders(root, args.folder, args.all)
    if not folders:
        print('未选择文件夹')
        return 1

    if '--out' not in sys.argv and not args.dry_run and not args.list:
        raw = input(f'PDF 保存目录（回车={args.out}）: ').strip().strip('"')
        if raw:
            args.out = Path(raw)
    if '--excel-out' not in sys.argv and not args.dry_run and not args.list:
        raw = input(f'题录副本目录（回车={args.excel_out}）: ').strip().strip('"')
        if raw:
            args.excel_out = Path(raw)

    print(f'题录根目录: {root}')
    print(f'PDF 保存: {args.out}')
    print(f'题录副本: {args.excel_out}')
    print('将处理:')
    jobs: list[Path] = []
    for folder in folders:
        files = iter_excel_files(folder)
        if args.limit_files and args.limit_files > 0:
            files = files[:args.limit_files]
        already = 0
        if args.resume:
            already = sum(1 for f in files if ExcelLock(f).is_done())
        print(
            f'  {folder.name}: {len(files)} 个 Excel'
            + (f'（已完成将跳过 {already}）' if already else '')
        )
        jobs.extend(files)

    if not jobs:
        print('没有可处理的 Excel')
        return 1

    if args.dry_run:
        for i, path in enumerate(jobs[:40], 1):
            mark = ' [done]' if ExcelLock(path).is_done() else ''
            print(f'  {i}. {path.relative_to(root)}{mark}')
        if len(jobs) > 40:
            print(f'  ... 其余 {len(jobs) - 40} 个')
        return 0

    pending_jobs = [
        p for p in jobs
        if not (args.resume and ExcelLock(p).is_done())
    ]
    if args.resume and len(pending_jobs) < len(jobs):
        print(
            f'续跑：{len(jobs)} 个 Excel 中 {len(jobs) - len(pending_jobs)} 个已完成将跳过，'
            f'待处理 {len(pending_jobs)} 个'
        )
    if not pending_jobs:
        print('全部 Excel 已有完成标记，无需打开浏览器')
        return 0

    args.out.mkdir(parents=True, exist_ok=True)
    args.excel_out.mkdir(parents=True, exist_ok=True)
    fail_log = args.out / '_fail.csv'
    debug_dir = args.out / '_captcha_debug'
    debug_dir.mkdir(parents=True, exist_ok=True)

    print('预加载点选/滑块验证码模型…')
    preload_click_captcha()

    driver = None
    totals = {'success': 0, 'fail': 0, 'skip': 0, 'excel_done': 0, 'excel_skip': 0}
    paper_counter = [0]

    def ensure_browser():
        nonlocal driver
        if driver is None or not is_driver_alive(driver):
            driver = recreate_driver(debug_dir, old_driver=driver)
        return driver

    try:
        for i, excel_path in enumerate(jobs, 1):
            rel = excel_path.relative_to(root)
            lock = ExcelLock(excel_path)
            if args.resume and lock.is_done():
                print(f'\n⊘ [{i}/{len(jobs)}] 已完成，跳过: {rel}')
                totals['excel_skip'] += 1
                continue
            print(f'\n======== Excel {i}/{len(jobs)}: {rel} ========')
            driver = ensure_browser()
            if driver is None:
                print('无法启动浏览器，跳过本表，稍后对下一表再试')
                time.sleep(3)
                continue
            try:
                driver, status, success, fail, skip = process_excel(
                    driver,
                    excel_path,
                    root,
                    args.out,
                    args.excel_out,
                    resume=args.resume,
                    mark_done_on_fail=args.mark_done_on_fail,
                    limit_papers=args.limit_papers,
                    fail_log=fail_log,
                    paper_counter=paper_counter,
                )
            except Exception as e:
                print(f'处理 Excel 异常，将继续下一表: {e}')
                driver = recreate_driver(debug_dir, old_driver=driver)
                continue
            totals['success'] += success
            totals['fail'] += fail
            totals['skip'] += skip
            if status in ('skipped_done', 'locked'):
                totals['excel_skip'] += 1
            elif status in ('done', 'empty'):
                totals['excel_done'] += 1
    except KeyboardInterrupt:
        print('\n已中断')
    finally:
        _force_kill_browser(driver)
        driver = None
        gc.collect()

    print('\n' + '=' * 60)
    print(
        f"Excel 完成标记: {totals['excel_done']} | 跳过: {totals['excel_skip']}\n"
        f"PDF 成功: {totals['success']} | 失败: {totals['fail']} | 已存在跳过: {totals['skip']}"
    )
    print(f"失败日志: {fail_log}" if totals['fail'] else '无失败记录')
    print('=' * 60)
    return 0 if totals['fail'] == 0 else 2


if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    raise SystemExit(main())

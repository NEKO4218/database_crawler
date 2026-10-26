# -*- coding: utf-8 -*-
"""用 requests 从知网详情页标记作者与单位对应，不下载 PDF。

对照 cnki_excel.py 的 wx-tit / #authorpart 解析，将
「作者标注」「单位列表」「作者单位对应」写入指定输出目录的 xlsx。

示例：
  python cnki_author.py                 # 默认交互：询问源路径、输出目录等
  python cnki_author.py --folder "D:\\2026年工作\\知网数据库下载\\信息科技"
  python cnki_author.py --folder 英文期刊 --out "D:\\2026年工作\\知网数据库下载\\_作者单位"
  python cnki_author.py --all --workers 12 --limit-files 1 --limit-papers 20
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from html import unescape
from pathlib import Path

import pandas as pd
import requests

from cnki_summary import (
    DOWNLOAD_ROOT,
    TITLE_COLUMNS,
    URL_COLUMNS,
    _setup_stdio,
)

RECORDS_ROOT = DOWNLOAD_ROOT
DEFAULT_OUT = RECORDS_ROOT / '_作者单位'
CATEGORY_ROOTS = [
    RECORDS_ROOT / '工程科技Ⅰ辑',
    RECORDS_ROOT / '工程科技Ⅱ辑',
    RECORDS_ROOT / '基础科学',
    RECORDS_ROOT / '农业科技',
    RECORDS_ROOT / '信息科技',
    RECORDS_ROOT / '医药卫生科技',
    RECORDS_ROOT / '英文期刊',
]
SKIP_DIR_NAMES = {
    'captcha_debug', '_pdf全文', '_题录', '_locks', '_作者单位',
}
XLS_PREFIX = 'CNKI-'
EXCEL_SUFFIXES = {'.xls', '.xlsx'}
COL_AUTHOR_MARK = '作者标注'
COL_ORG_LIST = '单位列表'
COL_AUTHOR_ORG = '作者单位对应'
AUTHOR_ORG_COLS = (COL_AUTHOR_MARK, COL_ORG_LIST, COL_AUTHOR_ORG)
MARK_URL_404 = '网址失效404'
MARK_NO_AUTHOR_ORG = '无作者单位'
CNKI_URL_PREFIXES = (
    'https://link.cnki.net',
    'http://link.cnki.net',
    'https://kns.cnki.net',
    'http://kns.cnki.net',
)
HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
        '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'Referer': 'https://kns.cnki.net/',
}
TAG_RE = re.compile(r'<[^>]+>', re.I)
SPACE_RE = re.compile(r'\s+')
AUTHORPART_RE = re.compile(
    r'<h3[^>]*\bid=["\']authorpart["\'][^>]*>(.*?)</h3>',
    re.I | re.S,
)
H3_AUTHOR_RE = re.compile(
    r'<h3([^>]*)>(.*?)</h3>',
    re.I | re.S,
)
WX_TIT_RE = re.compile(
    r'<div[^>]*class=["\'][^"\']*wx-tit[^"\']*["\'][^>]*>(.*?)</div>',
    re.I | re.S,
)
SPAN_RE = re.compile(r'<span\b[^>]*>(.*?)</span>', re.I | re.S)
SUP_RE = re.compile(r'<sup\b[^>]*>(.*?)</sup>', re.I | re.S)
ORG_LINE_RE = re.compile(r'^(\d+)\s*[.．、]\s*(.+)$')
ORG_SCAN_RE = re.compile(
    r'(\d+)\s*[.．、]\s*([^\d\n]+?)(?=(?:\n|\d+\s*[.．、]|$))'
)


def _pick_column(columns, candidates) -> str | None:
    colset = {str(c): c for c in columns}
    for name in candidates:
        if name in colset:
            return colset[name]
    lower = {str(c).strip().lower(): c for c in columns}
    for name in candidates:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def _td_texts(tr_html: str) -> list[str]:
    tds = re.findall(r'<td[^>]*>(.*?)</td>', tr_html, flags=re.I | re.S)
    return [strip_html(td) for td in tds]


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


def cell_text(value) -> str:
    if value is None:
        return ''
    text = str(value).strip()
    if text.lower() in ('nan', 'none'):
        return ''
    return text


def strip_html(text: str) -> str:
    text = TAG_RE.sub('', unescape(text or ''))
    return SPACE_RE.sub(' ', text).strip()


def is_cnki_url(url: str) -> bool:
    u = (url or '').strip().lower()
    return any(u.startswith(p.lower()) for p in CNKI_URL_PREFIXES) or 'cnki.net' in u


def has_author_org(value) -> bool:
    return bool(cell_text(value))


def row_has_author_org(row) -> bool:
    return any(has_author_org(row.get(col)) for col in AUTHOR_ORG_COLS)


def iter_excel_files(folder: Path) -> list[Path]:
    files = []
    if not folder.is_dir():
        return files
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


def dest_excel_path(src: Path, records_root: Path, out_root: Path) -> Path:
    try:
        rel = src.parent.resolve().relative_to(records_root.resolve())
    except ValueError:
        rel = Path(src.parent.name)
    return out_root / rel / f'{src.stem}.xlsx'


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
    mark = (mark or '').strip()
    if not mark:
        return
    for col in AUTHOR_ORG_COLS:
        df.at[df_index, col] = mark


# openpyxl 不允许的控制字符（与 openpyxl.cell.cell.ILLEGAL_CHARACTERS_RE 一致）
ILLEGAL_XLSX_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def sanitize_excel_value(value):
    if isinstance(value, str) and ILLEGAL_XLSX_RE.search(value):
        return ILLEGAL_XLSX_RE.sub('', value)
    return value


def sanitize_dataframe_for_xlsx(df: pd.DataFrame) -> pd.DataFrame:
    """摘要等字段里的控制字符会导致 openpyxl 报 cannot be used in worksheets。"""
    for col in df.columns:
        if df[col].dtype == object or str(df[col].dtype).startswith('string'):
            df[col] = df[col].map(sanitize_excel_value)
    return df


def save_working_excel(path: Path, df: pd.DataFrame) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    sanitize_dataframe_for_xlsx(df)
    tmp = path.with_name(path.name + '.tmp')
    df.to_excel(tmp, index=False, engine='openpyxl')
    tmp.replace(path)


def prepare_working_excel(src: Path, dest: Path) -> pd.DataFrame:
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
            print(f'  已合并已有标记: {dest}')
        except Exception as e:
            print(f'  读取已有输出失败，将重建: {e}')
    return df


def add_org(orgs: dict, key: str, name: str) -> None:
    key = str(key or '').strip()
    name = SPACE_RE.sub(' ', (name or '').strip()).strip(' ;；,，')
    if key and name and key not in orgs:
        orgs[key] = name


def parse_org_text(text: str, orgs: dict) -> None:
    text = (text or '').replace('\r', '')
    for m in ORG_SCAN_RE.finditer(text):
        add_org(orgs, m.group(1), m.group(2))
    for line in text.split('\n'):
        m = ORG_LINE_RE.match(SPACE_RE.sub(' ', line).strip())
        if m:
            add_org(orgs, m.group(1), m.group(2))


def extract_author_org_from_html(html: str) -> dict:
    """对照 cnki_excel.EXTRACT_AUTHOR_ORG_JS，从详情页 HTML 抽作者上标与单位。"""
    authors = []
    orgs = {}
    m = AUTHORPART_RE.search(html or '')
    author_html = m.group(1) if m else ''
    if not author_html:
        wx = WX_TIT_RE.search(html or '')
        block = wx.group(1) if wx else (html or '')
        for hm in H3_AUTHOR_RE.finditer(block):
            if re.search(r'class=["\'][^"\']*author', hm.group(1) or '', re.I):
                author_html = hm.group(2)
                break
    for span in SPAN_RE.findall(author_html or ''):
        sups = SUP_RE.findall(span)
        nums = []
        for sup in sups:
            nums.extend(re.findall(r'\d+', strip_html(sup)))
        name_html = SUP_RE.sub('', span)
        name = strip_html(name_html)
        name = re.sub(r'[\s;；,，]+$', '', name)
        name = re.sub(r'\d+(,\d+)*', '', name).strip()
        if name:
            authors.append({'name': name, 'orgs': nums})

    wx = WX_TIT_RE.search(html or '')
    block = wx.group(1) if wx else (html or '')
    for hm in H3_AUTHOR_RE.finditer(block):
        attrs, body = hm.group(1) or '', hm.group(2) or ''
        if re.search(r'id=["\']authorpart["\']', attrs, re.I):
            continue
        for span in SPAN_RE.findall(body):
            text = strip_html(span)
            mm = ORG_LINE_RE.match(text)
            if mm:
                add_org(orgs, mm.group(1), mm.group(2))
        plain = TAG_RE.sub('\n', unescape(body))
        parse_org_text(plain, orgs)
    if wx and not orgs:
        parse_org_text(TAG_RE.sub('\n', unescape(wx.group(1))), orgs)
    return {'authors': authors, 'orgs': orgs}


def classify_html(url: str, status: int, html: str) -> str:
    loc = (url or '').lower()
    text = html or ''
    if status == 404 or '页面不存在' in text or '您访问的页面不存在' in text:
        return 'not_found'
    title_m = re.search(r'<title[^>]*>(.*?)</title>', text, re.I | re.S)
    title = strip_html(title_m.group(1) if title_m else '')
    if '/verify/' in loc or ('验证' in title and ('码' in title or 'verify' in loc)):
        return 'verify'
    if any(x in loc for x in ('htmlread.html', 'settlementhtml', 'audownload.html')):
        return 'unavailable'
    if 'cas/login' in loc:
        return 'login'
    return 'ok'


def fetch_html(session: requests.Session, url: str, timeout: int) -> tuple[int, str, str]:
    r = session.get(url, timeout=timeout, allow_redirects=True)
    r.encoding = r.apparent_encoding or r.encoding or 'utf-8'
    return r.status_code, r.url or url, r.text or ''


def fetch_author_org(url: str, timeout: int = 20, retries: int = 3) -> tuple[str, dict]:
    session = requests.Session()
    session.headers.update(HEADERS)
    last_err = ''
    for attempt in range(1, retries + 1):
        try:
            status, final, html = fetch_html(session, url, timeout)
            kind = classify_html(final, status, html)
            if kind == 'not_found':
                return 'not_found', {}
            if kind == 'verify':
                last_err = 'verify'
                time.sleep(1.2 * attempt)
                continue
            if kind == 'login':
                return 'login', {}
            meta = extract_author_org_from_html(html)
            if meta.get('authors') or meta.get('orgs'):
                return 'ok', meta
            if kind == 'unavailable':
                return 'unavailable', meta
            if attempt < retries:
                time.sleep(0.4 * attempt)
                continue
            return 'no_meta', meta
        except requests.RequestException as e:
            last_err = str(e)
            time.sleep(0.6 * attempt)
    return 'fail', {'_error': last_err}


def papers_from_df(df: pd.DataFrame) -> list[dict]:
    url_col = _pick_column(df.columns, URL_COLUMNS)
    if not url_col:
        raise ValueError(f'无 URL 列，现有列: {list(df.columns)}')
    title_col = _pick_column(df.columns, TITLE_COLUMNS)
    papers = []
    for idx, row in df.iterrows():
        url = cell_text(row.get(url_col))
        if not url or not is_cnki_url(url):
            continue
        papers.append({
            'url': url,
            'title': cell_text(row.get(title_col)) if title_col else '',
            'df_index': idx,
            'has_meta': row_has_author_org(row),
        })
    return papers


def process_excel(
    src: Path,
    records_root: Path,
    out_root: Path,
    workers: int,
    limit_papers: int,
    resume: bool,
) -> tuple[int, int, int]:
    dest = dest_excel_path(src, records_root, out_root)
    df = prepare_working_excel(src, dest)
    papers = papers_from_df(df)
    if resume:
        already = [p for p in papers if p['has_meta']]
        pending = [p for p in papers if not p['has_meta']]
    else:
        already = []
        pending = list(papers)
    skip = len(already)
    if limit_papers and limit_papers > 0:
        pending = pending[:limit_papers]
    print(
        f'  条目 {len(papers)}，待标记 {len(pending)}'
        + (f'，已有跳过 {skip}' if skip else '')
    )
    if not pending:
        save_working_excel(dest, df)
        return 0, 0, skip

    ok = fail = 0
    workers = max(1, int(workers or 1))
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = {
            pool.submit(fetch_author_org, p['url']): p
            for p in pending
        }
        for fut in as_completed(futs):
            paper = futs[fut]
            done += 1
            try:
                status, meta = fut.result()
            except Exception as e:
                status, meta = 'fail', {'_error': str(e)}
            idx = paper['df_index']
            title = (paper.get('title') or '')[:40]
            if status == 'ok' and apply_author_org(df, idx, meta):
                ok += 1
                n_a = len(meta.get('authors') or [])
                n_o = len(meta.get('orgs') or {})
                print(f'    [{done}/{len(pending)}] {title}  作者{n_a}/单位{n_o}')
            elif status == 'not_found':
                apply_skip_mark(df, idx, MARK_URL_404)
                ok += 1
                print(f'    [{done}/{len(pending)}] {title}  网址失效404')
            elif status in ('no_meta', 'unavailable') and not (
                    meta.get('authors') or meta.get('orgs')):
                apply_skip_mark(df, idx, MARK_NO_AUTHOR_ORG)
                ok += 1
                print(f'    [{done}/{len(pending)}] {title}  无作者单位')
            elif meta and apply_author_org(df, idx, meta):
                ok += 1
                print(f'    [{done}/{len(pending)}] {title}  已标记')
            else:
                fail += 1
                print(f'    [{done}/{len(pending)}] {title}  失败 {status}')
            if done % 30 == 0:
                try:
                    save_working_excel(dest, df)
                except Exception as e:
                    print(f'    中途保存失败: {e}')
    save_working_excel(dest, df)
    print(f'  已写入: {dest}')
    return ok, fail, skip


def _split_input_paths(raw: str) -> list[str]:
    parts = re.split(r'[;；]', raw or '')
    return [p.strip().strip('"').strip("'") for p in parts if p.strip()]


def _resolve_input_path(raw: str) -> Path:
    """完整路径 / 题录根下的大类名 / 当前目录相对路径。"""
    path = Path(raw).expanduser()
    if path.is_absolute():
        return path
    by_root = RECORDS_ROOT / raw
    if by_root.is_dir():
        return by_root
    return (Path.cwd() / path).resolve()


def _input_line(prompt: str, default: str = '', hint: str | None = None) -> str:
    if hint is None:
        hint = f'（回车={default}）' if default else '（直接回车跳过）'
    raw = input(f'{prompt}{hint}: ').strip()
    if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in ('"', "'"):
        raw = raw[1:-1].strip()
    return raw or default


def _prompt_int(
    prompt: str, default: int, minimum: int = 0,
    maximum: int | None = None,
) -> int:
    while True:
        raw = _input_line(prompt, str(default))
        try:
            value = int(raw)
        except ValueError:
            print('  请输入整数。')
            continue
        if value < minimum or (maximum is not None and value > maximum):
            bounds = (f'{minimum}..{maximum}' if maximum is not None
                      else f'>= {minimum}')
            print(f'  请输入 {bounds} 的整数。')
            continue
        return value


def _prompt_yes_no(prompt: str, default: bool = True) -> bool:
    while True:
        raw = _input_line(prompt, 'Y' if default else 'N').lower()
        if raw in ('y', 'yes', '是', '1'):
            return True
        if raw in ('n', 'no', '否', '0'):
            return False
        print('  请输入 y 或 n。')


def interactive_setup(args) -> bool:
    """无命令行参数时，逐项询问源路径、输出目录与抓取参数。"""
    print('=' * 64)
    print('知网作者-单位标记（交互模式）')
    print(f'题录根目录: {RECORDS_ROOT}')
    available = [p for p in CATEGORY_ROOTS if p.is_dir()]
    if available:
        print('可用大类（可输入序号或名称）:')
        for i, path in enumerate(available, 1):
            print(f'  {i}. {path.name}')
        print('多个路径用分号分隔；直接回车处理全部大类。')
    else:
        print('题录根目录下暂未找到大类目录，请输入完整路径。')

    while True:
        raw = _input_line(
            '源文件夹（完整路径/大类名/序号，多个用分号）',
            hint='（回车=全部大类，q=取消）').strip()
        if raw.lower() in ('q', 'quit', 'exit'):
            return False
        if not raw:
            if not available:
                print('  没有可用的默认大类，请先输入源路径。')
                continue
            args.folder = []
            args.all = True
            break
        selected: list[Path] = []
        invalid: list[Path] = []
        for item in _split_input_paths(raw):
            if item.isdigit() and available and 1 <= int(item) <= len(available):
                selected.append(available[int(item) - 1])
                continue
            path = _resolve_input_path(item)
            if path.is_dir():
                selected.append(path)
            else:
                invalid.append(path)
        if invalid:
            for path in invalid:
                kind = '是文件，请输入目录' if path.is_file() else '不存在'
                print(f'  {path} {kind}')
            if selected:
                print('  已识别有效路径:')
                for path in selected:
                    print(f'    {path}')
            print('  请重新输入源路径。')
            continue
        args.folder = [str(path) for path in selected]
        args.all = False
        break

    while True:
        raw = _input_line(
            '标记结果保存目录', str(args.out),
            hint=f'（回车={args.out}，q=取消）')
        if raw.lower() in ('q', 'quit', 'exit'):
            return False
        out = Path(raw).expanduser()
        if out.is_file():
            print('  该路径是文件，请指定目录。')
            continue
        args.out = out if out.is_absolute() else (Path.cwd() / out).resolve()
        break

    args.workers = _prompt_int('并发请求数（1-64）', args.workers, 1, 64)
    args.limit_files = _prompt_int(
        '每个目录最多处理几个 Excel（0=不限）', args.limit_files, 0)
    args.limit_papers = _prompt_int(
        '每个 Excel 最多抓几篇（0=不限）', args.limit_papers, 0)
    args.resume = _prompt_yes_no('跳过已有作者单位标记的条目', True)
    args.list = _prompt_yes_no('只列出将处理的 Excel，不实际抓取', False)

    print('\n即将执行:')
    if args.all:
        print('  源目录: 全部大类')
    for path in args.folder:
        print(f'  源目录: {path}')
    print(f'  输出目录: {args.out}')
    print(f'  并发: {args.workers}')
    print(f'  每目录 Excel 上限: {args.limit_files or "不限"}')
    print(f'  每 Excel 论文上限: {args.limit_papers or "不限"}')
    print(f'  跳过已有标记: {"是" if args.resume else "否"}')
    print(f'  只列出: {"是" if args.list else "否"}')
    return _prompt_yes_no('确认开始', True)


def resolve_folders(names: list[str] | None, use_all: bool) -> list[Path]:
    if use_all or not names:
        return [p for p in CATEGORY_ROOTS if p.is_dir()]
    out = []
    for raw in names:
        path = _resolve_input_path(raw)
        if path.is_dir():
            out.append(path)
        else:
            print(f'目录不存在，跳过: {path}')
    return out


def parse_args():
    parser = argparse.ArgumentParser(description='用 requests 标记知网作者-单位对应')
    parser.add_argument(
        '--folder', action='append', default=[],
        help='源文件夹（完整路径或题录根下的目录名），可重复')
    parser.add_argument('--all', action='store_true', help='处理全部已存在的大类')
    parser.add_argument(
        '--out', type=Path, default=DEFAULT_OUT,
        help=f'标记结果保存根目录（默认 {DEFAULT_OUT}）')
    parser.add_argument('--workers', type=int, default=8, help='并发请求数')
    parser.add_argument('--limit-files', type=int, default=0, help='每个目录最多几个 Excel')
    parser.add_argument('--limit-papers', type=int, default=0, help='每个 Excel 最多几篇')
    parser.add_argument(
        '--no-resume', dest='resume', action='store_false',
        help='忽略已有标记，重新抓取')
    parser.set_defaults(resume=True)
    parser.add_argument('--list', action='store_true', help='只列出将处理的 Excel')
    return parser.parse_args()


def main() -> int:
    _setup_stdio()
    args = parse_args()
    interactive = len(sys.argv) == 1
    if interactive:
        if not interactive_setup(args):
            print('已取消')
            return 0
    folders = resolve_folders(args.folder, args.all)
    if not folders:
        print('未找到可处理目录')
        return 1
    if '--out' not in sys.argv and not args.list and not interactive:
        raw = input(f'标记结果目录（回车={args.out}）: ').strip().strip('"')
        if raw:
            args.out = Path(raw)
    args.out = Path(args.out)
    print(f'输出目录: {args.out}')
    print(f'并发: {args.workers}')

    jobs = []
    for folder in folders:
        files = iter_excel_files(folder)
        if args.limit_files and args.limit_files > 0:
            files = files[:args.limit_files]
        print(f'  {folder}: {len(files)} 个 Excel')
        jobs.extend(files)
    if not jobs:
        print('没有可处理的 Excel')
        return 1
    if args.list:
        for i, path in enumerate(jobs[:40], 1):
            print(f'  {i}. {path}')
        if len(jobs) > 40:
            print(f'  ... 其余 {len(jobs) - 40} 个')
        return 0

    args.out.mkdir(parents=True, exist_ok=True)
    total_ok = total_fail = total_skip = 0
    try:
        for i, src in enumerate(jobs, 1):
            print(f'\n======== Excel {i}/{len(jobs)}: {src} ========')
            ok, fail, skip = process_excel(
                src, RECORDS_ROOT, args.out,
                workers=args.workers,
                limit_papers=args.limit_papers,
                resume=args.resume,
            )
            total_ok += ok
            total_fail += fail
            total_skip += skip
    except KeyboardInterrupt:
        print('\n已中断')
    print('\n======== 合计 ========')
    print(f'新标记 {total_ok} | 失败 {total_fail} | 跳过已有 {total_skip}')
    print(f'输出: {args.out}')
    return 0 if total_fail == 0 else 2


if __name__ == '__main__':
    raise SystemExit(main())

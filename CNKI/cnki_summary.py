# -*- coding: utf-8 -*-
"""统计知网题录导出的总量、缺失时段与相邻时段重复。

对照 cnkiforexcel.py：
  - 导出目录：DOWNLOAD_ROOT 下各大类 / 英文期刊子分类
  - 失败记录：各导出目录 export_failures.txt（YYYY-MM-DD 或 起_止）
  - 期望区间：bgdate ~ initdate
  - 题录文件：CNKI-*.xls（多为 HTML 表，文件名是下载时刻，不是检索日期）

缺失：
  1. 失败文件中的时段（映射到期望粒度后标记）
  2. 期望区间内没有任何题录、也没有对应题录文件的时段（失败未记也标出）

重复：相邻时间段题录内容交叉（优先 URL）；并对比按文件名排序的相邻 Excel。

示例：
  python cnki_summary.py
  python cnki_summary.py --start 2023-01-01 --end 2026-06-30
  python cnki_summary.py --root "D:\\2026年工作\\知网数据库下载\\信息科技"
"""
from __future__ import annotations

import argparse
import calendar
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from html import unescape
from pathlib import Path

import pandas as pd

# 与 cnkiforexcel.py 保持一致
DOWNLOAD_ROOT = Path(r'D:\2026年工作\知网数据库下载')
BGDATE = '2023-01-01'
INITDATE = '2026-06-30'
EXPORT_FAILURE_LOG = 'export_failures.txt'
XLS_PREFIX = 'CNKI-'
SKIP_DIR_NAMES = {'captcha_debug', '_pdf全文', '_题录', '_locks'}
CATEGORY_ROOTS = [
    DOWNLOAD_ROOT / '工程科技Ⅰ辑',
    DOWNLOAD_ROOT / '工程科技Ⅱ辑',
    DOWNLOAD_ROOT / '基础科学',
    DOWNLOAD_ROOT / '农业科技',
    DOWNLOAD_ROOT / '信息科技',
    DOWNLOAD_ROOT / '医药卫生科技',
    DOWNLOAD_ROOT / '英文期刊',
]
URL_COLUMNS = ('URL-网址', 'URL', '网址', 'Link', '链接')
TITLE_COLUMNS = ('Title-题名', 'Title', '题名', '篇名')
AUTHOR_COLUMNS = ('Author-作者', 'Author', '作者')
DATE_COLUMNS = (
    'PubTime-发表时间', '发表时间', 'Year-年', '年', '日期', '出版日期',
)
LABEL_RE = re.compile(
    r'^(\d{4}-\d{2}-\d{2})(?:_(\d{4}-\d{2}-\d{2}))?$'
)
TR_RE = re.compile(r'<tr[^>]*>(.*?)</tr>', flags=re.I | re.S)
TD_RE = re.compile(r'<td[^>]*>(.*?)</td>', flags=re.I | re.S)
TAG_RE = re.compile(r'<[^>]+>')
SPACE_RE = re.compile(r'\s+')


def _setup_stdio():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')


def parse_iso_date(text: str) -> date | None:
    text = (text or '').strip()
    if not text:
        return None
    head = text[:10]
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d'):
        try:
            return datetime.strptime(head, fmt).date()
        except ValueError:
            continue
    m = re.match(r'^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})', text)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    if re.fullmatch(r'\d{4}', text):
        try:
            return date(int(text), 1, 1)
        except ValueError:
            return None
    return None


def iter_crawl_date_ranges(start: date, end: date, granularity: str):
    """与 cnkiforexcel.iter_crawl_date_ranges 相同。"""
    if start > end:
        return
    g = (granularity or 'all').strip().lower()
    if g in ('all', 'total', '整段', '总'):
        yield start, end
        return
    current = start
    while current <= end:
        if g in ('day', 'd', '天'):
            yield current, current
            current += timedelta(days=1)
            continue
        if g in ('month', 'm', '月'):
            last_day = calendar.monthrange(current.year, current.month)[1]
            period_end = min(current.replace(day=last_day), end)
            yield current, period_end
            current = period_end + timedelta(days=1)
            continue
        if g in ('year', 'y', '年'):
            period_end = min(current.replace(month=12, day=31), end)
            yield current, period_end
            current = period_end + timedelta(days=1)
            continue
        raise ValueError(f'未知检索粒度: {granularity}')


def period_label(start: date, end: date) -> str:
    if start == end:
        return start.isoformat()
    return f'{start.isoformat()}_{end.isoformat()}'


def parse_fail_label(text: str) -> tuple[date, date] | None:
    text = (text or '').strip()
    if not text or text.startswith('#'):
        return None
    m = LABEL_RE.match(text)
    if m:
        a = parse_iso_date(m.group(1))
        b = parse_iso_date(m.group(2) or m.group(1))
        if a and b:
            return (a, b) if a <= b else (b, a)
    if re.fullmatch(r'\d{4}', text):
        y = int(text)
        return date(y, 1, 1), date(y, 12, 31)
    if re.fullmatch(r'\d{4}-\d{2}', text):
        y, mth = int(text[:4]), int(text[5:7])
        last = calendar.monthrange(y, mth)[1]
        return date(y, mth, 1), date(y, mth, last)
    return None


def read_fail_labels(folder: Path) -> list[tuple[date, date]]:
    path = folder / EXPORT_FAILURE_LOG
    if not path.is_file():
        return []
    out = []
    seen = set()
    for line in path.read_text(encoding='utf-8', errors='replace').splitlines():
        parsed = parse_fail_label(line)
        if not parsed:
            continue
        if parsed in seen:
            continue
        seen.add(parsed)
        out.append(parsed)
    return out


def infer_granularity(fail_periods: list[tuple[date, date]], default: str) -> str:
    if not fail_periods:
        return default
    min_span = min((b - a).days + 1 for a, b in fail_periods)
    if min_span <= 2:
        return 'day'
    if min_span <= 40:
        return 'month'
    if min_span <= 400:
        return 'year'
    return 'all'


def _strip_html(cell: str) -> str:
    text = TAG_RE.sub('', cell or '')
    return SPACE_RE.sub(' ', unescape(text)).strip()


def _td_texts(tr_html: str) -> list[str]:
    return [_strip_html(td) for td in TD_RE.findall(tr_html)]


def _pick_index(headers: list[str], candidates) -> int | None:
    for name in candidates:
        if name in headers:
            return headers.index(name)
    lower = [h.strip().lower() for h in headers]
    for name in candidates:
        key = name.lower()
        if key in lower:
            return lower.index(key)
    return None


def _decode_xls(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ('utf-8-sig', 'utf-8', 'gb18030'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def _is_html_xls(path: Path) -> bool:
    head = path.read_bytes()[:240].lower()
    return b'<html' in head or b'<table' in head or b'<tr' in head


def iter_records_from_html_xls(path: Path):
    text = _decode_xls(path)
    rows = TR_RE.findall(text)
    if not rows:
        raise ValueError('未解析到表格行')
    headers = _td_texts(rows[0])
    if not headers:
        raise ValueError('表头为空')
    i_url = _pick_index(headers, URL_COLUMNS)
    i_title = _pick_index(headers, TITLE_COLUMNS)
    i_author = _pick_index(headers, AUTHOR_COLUMNS)
    i_date = _pick_index(headers, DATE_COLUMNS)
    for row_html in rows[1:]:
        cells = _td_texts(row_html)
        if not any(cells):
            continue
        def cell(idx):
            if idx is None or idx >= len(cells):
                return ''
            return cells[idx]
        yield cell(i_url), cell(i_title), cell(i_author), cell(i_date)


def iter_records_from_excel(path: Path):
    df = pd.read_excel(path)
    cols = list(df.columns)
    url_col = None
    title_col = None
    author_col = None
    date_col = None
    colset = {str(c): c for c in cols}
    def pick(cands):
        for name in cands:
            if name in colset:
                return colset[name]
        lower = {str(c).strip().lower(): c for c in cols}
        for name in cands:
            if name.lower() in lower:
                return lower[name.lower()]
        return None
    url_col = pick(URL_COLUMNS)
    title_col = pick(TITLE_COLUMNS)
    author_col = pick(AUTHOR_COLUMNS)
    date_col = pick(DATE_COLUMNS)
    for row in df.itertuples(index=False, name=None):
        mapping = {str(c): v for c, v in zip(cols, row)}
        def cell(col):
            if col is None:
                return ''
            val = mapping.get(col)
            if val is None:
                return ''
            text = str(val).strip()
            return '' if text.lower() in ('nan', 'none') else text
        yield cell(url_col), cell(title_col), cell(author_col), cell(date_col)


def iter_file_records(path: Path):
    if path.suffix.lower() == '.xls' and _is_html_xls(path):
        yield from iter_records_from_html_xls(path)
        return
    yield from iter_records_from_excel(path)


def record_key(url: str, title: str, author: str) -> str:
    url = (url or '').strip()
    if url:
        return 'u:' + url
    title = re.sub(r'\s+', '', title or '')
    author = re.sub(r'\s+', '', author or '')
    return 't:' + title + '|' + author


def iter_excel_files(folder: Path) -> list[Path]:
    files = []
    if not folder.is_dir():
        return files
    for path in folder.iterdir():
        if not path.is_file():
            continue
        if path.suffix.lower() not in {'.xls', '.xlsx'}:
            continue
        name = path.name
        if name.endswith('.crdownload') or name.endswith('.tmp'):
            continue
        if not name.startswith(XLS_PREFIX):
            continue
        files.append(path)
    files.sort()
    return files


def is_leaf_export_dir(folder: Path) -> bool:
    return bool(iter_excel_files(folder) or (folder / EXPORT_FAILURE_LOG).is_file())


def iter_export_units(root: Path) -> list[Path]:
    """英文期刊按子分类拆单元；中文大类目录本身即导出目录。"""
    if not root.is_dir():
        return []
    if is_leaf_export_dir(root):
        return [root]
    units = []
    for child in sorted(root.iterdir()):
        if not child.is_dir():
            continue
        if child.name in SKIP_DIR_NAMES or child.name.startswith('.'):
            continue
        if is_leaf_export_dir(child):
            units.append(child)
            continue
        units.extend(iter_export_units(child))
    return units or [root]


def unit_label(folder: Path) -> str:
    try:
        return str(folder.relative_to(DOWNLOAD_ROOT))
    except ValueError:
        return folder.name


def build_day_index(expected: list[tuple[date, date]]) -> dict[date, tuple[date, date]]:
    index = {}
    for start, end in expected:
        cur = start
        while cur <= end:
            index[cur] = (start, end)
            cur += timedelta(days=1)
    return index


def periods_overlap(a: tuple[date, date], b: tuple[date, date]) -> bool:
    return not (a[1] < b[0] or a[0] > b[1])


def fail_reason(in_fail: bool, n_rows: int, n_files: int) -> str | None:
    if in_fail and n_rows == 0 and n_files == 0:
        return '失败文件已标记'
    if in_fail and (n_rows > 0 or n_files > 0):
        return '失败文件已标记（目录中仍有该时段题录）'
    if n_rows == 0 and n_files == 0:
        return '数据中缺少（失败文件未标记）'
    return None


def summarize_unit(
    folder: Path,
    start: date,
    end: date,
    default_granularity: str,
    limit_files: int = 0,
    force_granularity: bool = False,
) -> dict:
    fail_periods = read_fail_labels(folder)
    if force_granularity:
        granularity = default_granularity
    else:
        granularity = infer_granularity(fail_periods, default_granularity)
    expected = list(iter_crawl_date_ranges(start, end, granularity))
    day_index = build_day_index(expected)

    files = iter_excel_files(folder)
    if limit_files and limit_files > 0:
        files = files[:limit_files]

    period_counts: dict[tuple[date, date], int] = defaultdict(int)
    period_unique: dict[tuple[date, date], set[str]] = defaultdict(set)
    period_files: dict[tuple[date, date], set[str]] = defaultdict(set)
    period_samples: dict[tuple[date, date], list[str]] = defaultdict(list)
    file_keys: dict[str, set[str]] = {}
    file_majority: dict[str, tuple[date, date] | None] = {}
    unknown = 0
    total_rows = 0
    errors = []
    total = len(files)

    for i, path in enumerate(files, 1):
        if i == 1 or i % 80 == 0 or i == total:
            print(f'    读取题录 {i}/{total}: {path.name}')
        keys = set()
        period_hits = Counter()
        try:
            for url, title, author, pub_text in iter_file_records(path):
                key = record_key(url, title, author)
                if key in ('u:', 't:|'):
                    continue
                total_rows += 1
                keys.add(key)
                pub = parse_iso_date(pub_text)
                period = day_index.get(pub) if pub else None
                if period is None:
                    unknown += 1
                    continue
                period_counts[period] += 1
                period_unique[period].add(key)
                period_files[period].add(path.name)
                period_hits[period] += 1
                if len(period_samples[period]) < 3:
                    period_samples[period].append(title or url)
        except Exception as e:
            errors.append(f'{path.name}: {e}')
            file_keys[path.name] = keys
            file_majority[path.name] = None
            continue
        file_keys[path.name] = keys
        file_majority[path.name] = period_hits.most_common(1)[0][0] if period_hits else None

    files_by_majority: dict[tuple[date, date], list[str]] = defaultdict(list)
    for fname, period in file_majority.items():
        if period is not None:
            files_by_majority[period].append(fname)

    fail_hit = set()
    for exp in expected:
        for fail in fail_periods:
            if periods_overlap(exp, fail):
                fail_hit.add(exp)
                break

    volume_rows = []
    missing_rows = []
    label = unit_label(folder)
    for ps, pe in expected:
        period = (ps, pe)
        n_rows = period_counts.get(period, 0)
        n_unique = len(period_unique.get(period, ()))
        content_files = sorted(period_files.get(period, ()))
        majority_files = sorted(files_by_majority.get(period, ()))
        n_files = len(set(content_files) | set(majority_files))
        in_fail = period in fail_hit
        volume_rows.append({
            '分类目录': label,
            '时段': period_label(ps, pe),
            '题录条数': n_rows,
            '去重条数': n_unique,
            '对应题录文件数': n_files,
            '失败文件已标记': '是' if in_fail else '否',
        })
        reason = fail_reason(in_fail, n_rows, n_files)
        if reason:
            missing_rows.append({
                '分类目录': label,
                '时段': period_label(ps, pe),
                '题录条数': n_rows,
                '对应题录文件数': n_files,
                '原因': reason,
            })

    for fail in fail_periods:
        if any(periods_overlap(exp, fail) for exp in expected):
            continue
        missing_rows.append({
            '分类目录': label,
            '时段': period_label(*fail),
            '题录条数': 0,
            '对应题录文件数': 0,
            '原因': '失败文件已标记（超出期望区间或粒度无法对齐）',
        })

    dup_period_rows = []
    populated = [p for p in expected if period_unique.get(p)]
    for i in range(len(populated) - 1):
        a, b = populated[i], populated[i + 1]
        overlap = period_unique[a] & period_unique[b]
        if not overlap:
            continue
        samples = period_samples.get(a, [])[:3]
        dup_period_rows.append({
            '分类目录': label,
            '时段A': period_label(*a),
            '时段B': period_label(*b),
            '重复条数': len(overlap),
            '示例': ' | '.join(samples),
        })

    dup_file_rows = []
    names = [p.name for p in files]
    for i in range(len(names) - 1):
        fa, fb = names[i], names[i + 1]
        overlap = file_keys.get(fa, set()) & file_keys.get(fb, set())
        if not overlap:
            continue
        dup_file_rows.append({
            '分类目录': label,
            '文件A': fa,
            '文件B': fb,
            '文件A多数时段': period_label(*file_majority[fa]) if file_majority.get(fa) else '',
            '文件B多数时段': period_label(*file_majority[fb]) if file_majority.get(fb) else '',
            '重复条数': len(overlap),
        })

    unique_all = set()
    for keys in period_unique.values():
        unique_all |= keys
    for keys in file_keys.values():
        unique_all |= keys

    return {
        'unit': label,
        'granularity': granularity,
        'excel_files': len(files),
        'rows': total_rows,
        'unique': len(unique_all),
        'unknown_date': unknown,
        'expected_periods': len(expected),
        'missing_fail': sum(1 for r in missing_rows if r['原因'] == '失败文件已标记'),
        'missing_unmarked': sum(
            1 for r in missing_rows if r['原因'].startswith('数据中缺少')),
        'missing_fail_but_has_data': sum(
            1 for r in missing_rows if '仍有' in r['原因']),
        'dup_period_pairs': len(dup_period_rows),
        'dup_period_rows_count': sum(r['重复条数'] for r in dup_period_rows),
        'dup_file_pairs': len(dup_file_rows),
        'read_errors': errors,
        'volume_rows': volume_rows,
        'missing_rows': missing_rows,
        'dup_period_rows': dup_period_rows,
        'dup_file_rows': dup_file_rows,
    }


def default_granularity_for_root(root: Path) -> str:
    if root.name == '英文期刊':
        return 'month'
    return 'day'


def write_report(out_path: Path, details: list[dict]):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    overview_df = pd.DataFrame([{
        '分类': d['unit'],
        '粒度': d['granularity'],
        '题录文件数': d['excel_files'],
        '题录总条数': d['rows'],
        '去重条数': d['unique'],
        '日期无法解析': d['unknown_date'],
        '期望时段数': d['expected_periods'],
        '缺失_失败已标记': d['missing_fail'],
        '缺失_失败未标记': d['missing_unmarked'],
        '失败已标记但仍有题录': d['missing_fail_but_has_data'],
        '相邻时段重复对数': d['dup_period_pairs'],
        '相邻时段重复条数合计': d['dup_period_rows_count'],
        '相邻文件重复对数': d['dup_file_pairs'],
        '读取失败文件数': len(d['read_errors']),
    } for d in details])
    volume_df = pd.DataFrame([row for d in details for row in d['volume_rows']])
    missing_df = pd.DataFrame([row for d in details for row in d['missing_rows']])
    dup_p_df = pd.DataFrame([row for d in details for row in d['dup_period_rows']])
    dup_f_df = pd.DataFrame([row for d in details for row in d['dup_file_rows']])
    err_df = pd.DataFrame([
        {'分类': d['unit'], '错误': msg}
        for d in details for msg in d['read_errors']
    ])

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        overview_df.to_excel(writer, sheet_name='总览', index=False)
        if not volume_df.empty:
            volume_df.to_excel(writer, sheet_name='各时段数据量', index=False)
        if not missing_df.empty:
            missing_df.to_excel(writer, sheet_name='缺失时段', index=False)
        if not dup_p_df.empty:
            dup_p_df.to_excel(writer, sheet_name='相邻时段重复', index=False)
        if not dup_f_df.empty:
            dup_f_df.to_excel(writer, sheet_name='相邻文件重复', index=False)
        if not err_df.empty:
            err_df.to_excel(writer, sheet_name='读取错误', index=False)
    print(f'\n已写入统计表: {out_path}')


def parse_args():
    parser = argparse.ArgumentParser(description='统计知网题录总量、缺失与重复')
    parser.add_argument(
        '--start', default=BGDATE, help=f'期望起始日期（默认 {BGDATE}）')
    parser.add_argument(
        '--end', default=INITDATE, help=f'期望结束日期（默认 {INITDATE}）')
    parser.add_argument(
        '--out', default=str(DOWNLOAD_ROOT / 'cnki_summary.xlsx'),
        help='统计结果 xlsx 路径')
    parser.add_argument(
        '--root', action='append', default=None,
        help='只统计指定大类目录，可重复；默认七个大类')
    parser.add_argument(
        '--limit-files', type=int, default=0,
        help='调试：每个单元最多读 N 个 Excel，0=全部')
    parser.add_argument(
        '--granularity', default='',
        help='强制粒度 day/month/year/all；默认按失败记录推断')
    return parser.parse_args()


def main():
    _setup_stdio()
    args = parse_args()
    start = parse_iso_date(args.start)
    end = parse_iso_date(args.end)
    if not start or not end or start > end:
        print('日期无效')
        return 2
    roots = [Path(p) for p in args.root] if args.root else CATEGORY_ROOTS
    details = []
    print(f'期望区间: {start} ~ {end}')
    print(f'对照失败文件: {EXPORT_FAILURE_LOG}')
    for root in roots:
        print(f'\n======== {root} ========')
        if not root.is_dir():
            print('  目录不存在，跳过')
            continue
        units = iter_export_units(root)
        force_g = bool(args.granularity.strip())
        default_g = args.granularity.strip() or default_granularity_for_root(root)
        print(f'  导出单元 {len(units)} 个，默认粒度 {default_g}')
        for unit in units:
            print(f'\n  -- {unit} --')
            info = summarize_unit(
                unit, start, end, default_g,
                limit_files=args.limit_files,
                force_granularity=force_g,
            )
            details.append(info)
            print(
                f'    文件 {info["excel_files"]} | 条数 {info["rows"]} | 去重 {info["unique"]}\n'
                f'    缺失: 失败已标 {info["missing_fail"]} / 未标缺数 {info["missing_unmarked"]}'
                f' / 失败但有数 {info["missing_fail_but_has_data"]}\n'
                f'    相邻时段重复 {info["dup_period_pairs"]} 对'
                f'（{info["dup_period_rows_count"]} 条）'
                f' | 相邻文件重复 {info["dup_file_pairs"]} 对'
            )
            if info['read_errors']:
                print(f'    读取失败 {len(info["read_errors"])} 个文件')

    if not details:
        print('无统计结果')
        return 1
    total_rows = sum(d['rows'] for d in details)
    total_unique = sum(d['unique'] for d in details)
    total_missing_u = sum(d['missing_unmarked'] for d in details)
    total_missing_f = sum(d['missing_fail'] for d in details)
    print('\n======== 合计 ========')
    print(f'题录条数 {total_rows} | 各单元去重合计 {total_unique}')
    print(f'缺失时段: 失败已标记 {total_missing_f} | 失败未标记 {total_missing_u}')
    write_report(Path(args.out), details)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

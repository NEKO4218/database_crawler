# -*- coding: utf-8 -*-
"""根据 cnki_summary 结果表，删除相邻且题录相同的文件 B。

读取 cnki_summary.xlsx 的「相邻文件重复」工作表，对比每对 文件A / 文件B
的 URL-网址 与 Title-题名。URL 会去掉一次性查询参数（如摘要页 v=），
题名会去 HTML 标签后再比；两列记录多重集一致则删除文件 B，保留 A。

部分交叉、条数不同的不删。A=B=C 连锁重复时，后一对会跟已保留的 A 再比。

示例：
  python cnki_dedup_files.py
  python cnki_dedup_files.py --dry-run
  python cnki_dedup_files.py --summary "D:\\2026年工作\\知网数据库下载\\cnki_summary.xlsx"
"""
from __future__ import annotations

import argparse
import os
import re
from collections import Counter
from html import unescape
from pathlib import Path

import pandas as pd

from cnki_summary import (
    DOWNLOAD_ROOT,
    _setup_stdio,
    iter_file_records,
)

SHEET_NAME = '相邻文件重复'
DEFAULT_SUMMARY = DOWNLOAD_ROOT / 'cnki_summary.xlsx'
SPACE_RE = re.compile(r'\s+')
TAG_RE = re.compile(r'<[^>]+>')


def norm_text(text: str) -> str:
    text = unescape(text or '')
    text = TAG_RE.sub('', text)
    return SPACE_RE.sub(' ', text).strip()


def normalize_url(url: str) -> str:
    """去掉知网摘要页一次性 v= 等查询参数，只保留可稳定比对的路径。"""
    url = norm_text(url)
    if not url:
        return ''
    url = url.split('#', 1)[0].split('?', 1)[0].rstrip('/')
    return url.lower()


def url_title_counter(path: Path) -> Counter:
    keys = Counter()
    for url, title, _author, _pub in iter_file_records(path):
        url = normalize_url(url)
        title = norm_text(title)
        if not url and not title:
            continue
        keys[(url, title)] += 1
    return keys


def unit_dir(category: str) -> Path:
    text = str(category or '').strip().replace('/', os.sep).replace('\\', os.sep)
    return DOWNLOAD_ROOT / text


def compare_pair(path_a: Path, path_b: Path) -> tuple[bool, str]:
    if path_a.resolve() == path_b.resolve():
        return False, '同一文件'
    try:
        keys_a = url_title_counter(path_a)
        keys_b = url_title_counter(path_b)
    except Exception as e:
        return False, f'解析失败: {e}'
    n_a, n_b = sum(keys_a.values()), sum(keys_b.values())
    if n_a == 0 and n_b == 0:
        return True, 'URL/题名均为空'
    if keys_a == keys_b:
        return True, f'URL与题名完全相同({n_a}条)'
    overlap = sum((keys_a & keys_b).values())
    return False, f'URL/题名不完全相同 A={n_a} B={n_b} 交叉={overlap}'


def load_dup_rows(summary_path: Path) -> pd.DataFrame:
    if not summary_path.is_file():
        raise FileNotFoundError(f'找不到统计结果: {summary_path}')
    xl = pd.ExcelFile(summary_path)
    if SHEET_NAME not in xl.sheet_names:
        raise ValueError(
            f'{summary_path} 没有「{SHEET_NAME}」表（可能没有相邻文件重复）')
    df = pd.read_excel(summary_path, sheet_name=SHEET_NAME)
    needed = {'分类目录', '文件A', '文件B'}
    missing = needed - set(df.columns)
    if missing:
        raise ValueError(f'表头缺少: {", ".join(sorted(missing))}')
    return df


def parse_args():
    parser = argparse.ArgumentParser(
        description='按 URL-网址 与 Title-题名 删除相邻且相同的题录文件 B')
    parser.add_argument(
        '--summary', default=str(DEFAULT_SUMMARY),
        help=f'cnki_summary 结果 xlsx（默认 {DEFAULT_SUMMARY}）')
    parser.add_argument(
        '--dry-run', action='store_true',
        help='只报告，不删除')
    parser.add_argument(
        '--out', default='',
        help='处理日志 xlsx；默认与 summary 同目录 cnki_dedup_files.xlsx')
    return parser.parse_args()


def main():
    _setup_stdio()
    args = parse_args()
    summary_path = Path(args.summary)
    df = load_dup_rows(summary_path)
    out_path = Path(args.out) if args.out else (
        summary_path.parent / 'cnki_dedup_files.xlsx')

    # 被删文件 -> 实际保留文件（处理 A=B=C 连锁）
    kept_for: dict[Path, Path] = {}

    def follow(path: Path) -> Path:
        seen = set()
        while path in kept_for and path not in seen:
            seen.add(path)
            path = kept_for[path]
        return path

    logs = []
    deleted = 0
    skipped = 0
    missing = 0
    total = len(df)
    print(f'结果表: {summary_path}')
    print(f'相邻文件重复 {total} 对')
    print('模式: ' + ('仅预览' if args.dry_run else '删除 URL/题名 完全相同的文件B'))

    cols = list(df.columns)
    for n, raw in enumerate(df.itertuples(index=False, name=None), 1):
        row = {str(c): v for c, v in zip(cols, raw)}
        category = str(row.get('分类目录') or '').strip()
        name_a = str(row.get('文件A') or '').strip()
        name_b = str(row.get('文件B') or '').strip()
        folder = unit_dir(category)
        orig_a = (folder / name_a).resolve()
        orig_b = (folder / name_b).resolve()
        path_a = follow(orig_a)
        path_b = orig_b

        if orig_b in kept_for:
            logs.append({
                '分类目录': category,
                '文件A': name_a,
                '文件B': name_b,
                '结果': '跳过',
                '说明': '文件B已在先前配对中删除',
            })
            skipped += 1
            continue

        if not path_a.is_file():
            logs.append({
                '分类目录': category,
                '文件A': name_a,
                '文件B': name_b,
                '结果': '缺失',
                '说明': f'找不到文件A: {path_a}',
            })
            missing += 1
            continue
        if not path_b.is_file():
            logs.append({
                '分类目录': category,
                '文件A': name_a,
                '文件B': name_b,
                '结果': '缺失',
                '说明': f'找不到文件B: {path_b}',
            })
            missing += 1
            continue

        identical, reason = compare_pair(path_a, path_b)
        if not identical:
            print(f'  [{n}/{total}] 保留 {name_b} | {category} | {reason}')
            logs.append({
                '分类目录': category,
                '文件A': name_a,
                '文件B': name_b,
                '文件A路径': str(path_a),
                '文件B路径': str(path_b),
                '结果': '保留',
                '说明': reason,
            })
            skipped += 1
            continue

        action = '将删除' if args.dry_run else '已删除'
        print(f'  [{n}/{total}] {action} {name_b} | {category} | {reason}')
        if not args.dry_run:
            path_b.unlink()
        kept_for[orig_b] = path_a
        deleted += 1
        logs.append({
            '分类目录': category,
            '文件A': name_a,
            '文件B': name_b,
            '文件A路径': str(path_a),
            '文件B路径': str(orig_b),
            '结果': '预览删除' if args.dry_run else '已删除',
            '说明': reason,
        })

    log_df = pd.DataFrame(logs)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    log_df.to_excel(out_path, index=False)
    print('\n======== 合计 ========')
    print(f'完全相同并{"预览" if args.dry_run else "删除"}: {deleted}')
    print(f'不完全相同/已处理过: {skipped}')
    print(f'文件缺失: {missing}')
    print(f'日志: {out_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

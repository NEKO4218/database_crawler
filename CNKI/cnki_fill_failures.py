# -*- coding: utf-8 -*-
"""根据 cnki_summary.xlsx「缺失时段」把日期补入各目录 export_failures.txt。

补入：
  - 数据中缺少（失败文件未标记）
  - 失败文件已标记（目录中仍有该时段题录）（文件中尚无该行时）

「失败文件已标记」本身已在 txt 中，不改。

示例：
  python cnki_fill_failures.py
  python cnki_fill_failures.py --summary "D:\\2026年工作\\知网数据库下载\\cnki_summary.xlsx"
"""
from __future__ import annotations

import argparse
import os
from collections import defaultdict
from pathlib import Path

import pandas as pd

from cnki_summary import (
    DOWNLOAD_ROOT,
    EXPORT_FAILURE_LOG,
    _setup_stdio,
)

DEFAULT_SUMMARY = DOWNLOAD_ROOT / 'cnki_summary.xlsx'
SHEET_NAME = '缺失时段'
REASONS_APPEND = (
    '数据中缺少（失败文件未标记）',
    '失败文件已标记（目录中仍有该时段题录）',
)


def unit_dir_from_label(label: str) -> Path:
    text = str(label or '').strip().replace('/', os.sep).replace('\\', os.sep)
    return DOWNLOAD_ROOT / text


def should_append(reason: str) -> bool:
    text = str(reason or '').strip()
    return any(text.startswith(r) for r in REASONS_APPEND)


def existing_fail_lines(folder: Path) -> set[str]:
    path = folder / EXPORT_FAILURE_LOG
    if not path.is_file():
        return set()
    return {
        line.strip()
        for line in path.read_text(encoding='utf-8', errors='replace').splitlines()
        if line.strip() and not line.strip().startswith('#')
    }


def append_fail_label(folder: Path, label: str) -> bool:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / EXPORT_FAILURE_LOG
    existing = existing_fail_lines(folder)
    if label in existing:
        return False
    with path.open('a', encoding='utf-8') as f:
        f.write(label + '\n')
    return True


def parse_args():
    parser = argparse.ArgumentParser(description='把缺失时段补入 export_failures.txt')
    parser.add_argument(
        '--summary', default=str(DEFAULT_SUMMARY),
        help=f'cnki_summary 结果（默认 {DEFAULT_SUMMARY}）')
    parser.add_argument(
        '--out', default='',
        help='对照表 xlsx；默认与 summary 同目录 cnki_fill_failures.xlsx')
    parser.add_argument(
        '--dry-run', action='store_true', help='只报告，不写失败文件')
    return parser.parse_args()


def main() -> int:
    _setup_stdio()
    args = parse_args()
    summary_path = Path(args.summary)
    if not summary_path.is_file():
        print(f'找不到统计结果: {summary_path}')
        return 1
    xl = pd.ExcelFile(summary_path)
    if SHEET_NAME not in xl.sheet_names:
        print(f'{summary_path} 没有「{SHEET_NAME}」表')
        return 1
    df = pd.read_excel(summary_path, sheet_name=SHEET_NAME)
    needed = {'分类目录', '时段', '原因'}
    missing = needed - set(df.columns)
    if missing:
        print('表头缺少: ' + ', '.join(sorted(missing)))
        return 1

    out_path = Path(args.out) if args.out else (
        summary_path.parent / 'cnki_fill_failures.xlsx')
    logs = []
    added = defaultdict(int)
    existed = defaultdict(int)
    skipped = defaultdict(int)
    cols = list(df.columns)

    print(f'结果表: {summary_path}')
    print(f'缺失时段 {len(df)} 行')
    print('模式: ' + ('仅预览' if args.dry_run else '追加失败日期'))

    for raw in df.itertuples(index=False, name=None):
        row = {str(c): v for c, v in zip(cols, raw)}
        category = str(row.get('分类目录') or '').strip()
        label = str(row.get('时段') or '').strip()
        reason = str(row.get('原因') or '').strip()
        folder = unit_dir_from_label(category)
        rec = {
            '分类目录': category,
            '时段': label,
            '原因': reason,
            '目录': str(folder),
        }
        if not category or not label or label.lower() == 'nan':
            rec['结果'] = '跳过'
            rec['说明'] = '分类或时段为空'
            skipped[reason] += 1
            logs.append(rec)
            continue
        if not should_append(reason):
            rec['结果'] = '跳过'
            rec['说明'] = '该原因不补入失败文件'
            skipped[reason] += 1
            logs.append(rec)
            continue
        if label in existing_fail_lines(folder):
            rec['结果'] = '已存在'
            rec['说明'] = 'export_failures.txt 已有该时段'
            existed[reason] += 1
            logs.append(rec)
            continue
        if args.dry_run:
            rec['结果'] = '将补入'
            rec['说明'] = 'dry-run'
            added[reason] += 1
            logs.append(rec)
            continue
        if append_fail_label(folder, label):
            rec['结果'] = '已补入'
            rec['说明'] = ''
            added[reason] += 1
        else:
            rec['结果'] = '已存在'
            rec['说明'] = '写入时发现已有'
            existed[reason] += 1
        logs.append(rec)

    pd.DataFrame(logs).to_excel(out_path, index=False)
    print('\n======== 合计 ========')
    print('新补入:')
    for k, n in added.items():
        print(f'  {k}: {n}')
    if not added:
        print('  （无）')
    print('已存在:')
    for k, n in existed.items():
        print(f'  {k}: {n}')
    print('跳过:')
    for k, n in skipped.items():
        print(f'  {k}: {n}')
    print(f'对照表: {out_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())

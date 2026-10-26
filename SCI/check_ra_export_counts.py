# -*- coding: utf-8 -*-
"""按文件名核对 Research Areas 导出批次，并可补导缺失 / 删除多余文件。

1) 核对模式（默认）
   输入统计表两列：Research Areas | Count
   输出四列：Research Areas | Count | 结果 | 多余文件

2) 补导模式（--fill-missing）
   读取核对结果表，解析「结果」列中的缺失批次，仅导出这些批次。

3) 删除多余（--delete-extra）
   按 Count 计算应有批次，删除不在应有范围内的文件，以及同批次重复文件（保留 1 个）。
   可用统计表或核对结果提供 Research Areas / Count；支持 --dry-run 只预览。

示例：
  python check_ra_export_counts.py -s 统计.xlsx -r 导出目录 -d 2023-01-01_2023-12-31
  python check_ra_export_counts.py --fill-missing -c 核对结果.xlsx -r 导出目录 -d 2023-01-01_2023-12-31
  python check_ra_export_counts.py --delete-extra -c 核对结果.xlsx -r 导出目录 -d 2023-01-01_2023-12-31
  python check_ra_export_counts.py --delete-extra --dry-run -s 统计.xlsx -r 导出目录 -d 2023-01-01_2023-12-31
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict

BATCH_SIZE = 1000
RANGE_RE = re.compile(r'_(\d+)-(\d+)$')
MISSING_BATCH_RE = re.compile(r'(\d+)\s*-\s*(\d+)')


def expected_batches(count: int, batch_size: int = BATCH_SIZE):
    """按导出逻辑生成应有批次，如 2762 → (1,1000),(1001,2000),(2001,2762)。"""
    count = int(count or 0)
    if count <= 0:
        return []
    batches = []
    start = 1
    while start <= count:
        end = min(start + batch_size - 1, count)
        batches.append((start, end))
        start = end + 1
    return batches


def _norm_area(name: str) -> str:
    text = re.sub(r'[<>:"/\\|?*]+', '_', str(name or '').strip())
    text = re.sub(r'\s+', ' ', text).strip(' ._')
    return text


def parse_missing_batches_from_result(result_text: str):
    """从「结果」列解析缺失批次，如 '缺失: 1001-2000, 2001-2762；且有多余'。"""
    text = str(result_text or '').strip()
    if '缺失' not in text:
        return []
    # 只取「缺失:」之后、「；」或行尾之前
    m = re.search(r'缺失\s*[:：]\s*(.+)$', text)
    chunk = m.group(1) if m else text
    chunk = re.split(r'[；;]', chunk, maxsplit=1)[0]
    batches = []
    seen = set()
    for mm in MISSING_BATCH_RE.finditer(chunk):
        s, e = int(mm.group(1)), int(mm.group(2))
        if s <= 0 or e < s:
            continue
        key = (s, e)
        if key in seen:
            continue
        seen.add(key)
        batches.append(key)
    return batches


def read_stats_table(path: str):
    """读取统计表，返回 [(area, count), ...]。"""
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f'统计表不存在: {path}')

    lower = path.lower()
    rows = []

    if lower.endswith('.csv'):
        import csv
        with open(path, 'r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if not row or len(row) < 2:
                    continue
                area = str(row[0]).strip()
                if not area:
                    continue
                try:
                    count = int(float(str(row[1]).replace(',', '').strip()))
                except ValueError:
                    continue
                rows.append((area, count))
        return rows

    try:
        import openpyxl
    except ImportError as e:
        raise SystemExit('需要 openpyxl：pip install openpyxl') from e

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    first = True
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        if first:
            first = False
            a0 = str(row[0]).strip().lower()
            if 'research' in a0 or 'area' in a0 or a0 in ('研究方向', '名称'):
                continue
        area = str(row[0]).strip()
        if not area:
            continue
        try:
            count = int(float(str(row[1]).replace(',', '').strip()))
        except (TypeError, ValueError, IndexError):
            continue
        rows.append((area, count))
    wb.close()
    return rows


def read_check_result(path: str):
    """读取核对结果表，返回 [{area, count, result, extra, missing}, ...]。"""
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise FileNotFoundError(f'核对结果不存在: {path}')

    try:
        import openpyxl
    except ImportError as e:
        raise SystemExit('需要 openpyxl：pip install openpyxl') from e

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    first = True
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        if first:
            first = False
            a0 = str(row[0]).strip().lower()
            if 'research' in a0 or 'area' in a0 or a0 in ('研究方向', '名称'):
                continue
        area = str(row[0]).strip()
        if not area:
            continue
        try:
            count = int(float(str(row[1]).replace(',', '').strip()))
        except (TypeError, ValueError, IndexError):
            count = 0
        result = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ''
        extra = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ''
        missing = parse_missing_batches_from_result(result)
        rows.append({
            'name': area,
            'count': count,
            'result': result,
            'extra': extra,
            'batches': missing,
        })
    wb.close()
    return rows


def iter_export_files(root: str, date_label: str):
    """遍历根目录下含日期标签的导出文件。"""
    root = os.path.abspath(root)
    date_token = date_label.strip()
    for dirpath, _dirs, filenames in os.walk(root):
        for name in filenames:
            lower = name.lower()
            if lower.endswith(('.crdownload', '.tmp')):
                continue
            if not lower.endswith(('.xls', '.xlsx', '.csv')):
                continue
            if date_token not in name:
                continue
            stem = os.path.splitext(name)[0]
            m = RANGE_RE.search(stem)
            if not m:
                continue
            start_n, end_n = int(m.group(1)), int(m.group(2))
            prefix = stem[: m.start()]
            parts = os.path.normpath(dirpath).split(os.sep)
            area = None
            if 'ResearchAreas' in parts:
                idx = parts.index('ResearchAreas')
                if idx + 1 < len(parts):
                    area = parts[idx + 1]
            if not area:
                marker = f'SCI_{date_token}_'
                if prefix.startswith(marker):
                    rest = prefix[len(marker):]
                elif prefix.startswith('SCI_') and date_token in prefix:
                    rest = prefix.split(date_token + '_', 1)[-1]
                else:
                    rest = prefix
                area = rest
            rel = os.path.relpath(os.path.join(dirpath, name), root)
            yield area, start_n, end_n, name, rel


def match_files_for_area(area: str, file_index: dict):
    """从索引中取出属于该研究方向的 {(start,end): [relpaths]}。"""
    key = _norm_area(area)
    key_l = key.lower()
    found = defaultdict(list)

    for k, batches in file_index.items():
        k_l = k.lower()
        if k_l == key_l or k_l.endswith('_' + key_l):
            for rng, paths in batches.items():
                found[rng].extend(paths)

    if found:
        return found

    for k, batches in file_index.items():
        if _norm_area(k).lower().endswith(key_l):
            for rng, paths in batches.items():
                found[rng].extend(paths)
    return found


def build_file_index(root: str, date_label: str):
    """area -> {(start,end): [relpath,...]}"""
    index = defaultdict(lambda: defaultdict(list))
    for area, start_n, end_n, _name, rel in iter_export_files(root, date_label):
        area_key = _norm_area(area)
        index[area_key][(start_n, end_n)].append(rel)
    return index


def collect_extra_paths(
        area: str, count: int, file_index: dict, batch_size: int = BATCH_SIZE):
    """找出该研究方向应删除的多余文件相对路径。

    - 不在应有批次内的区间：全部删除
    - 应有批次但存在多个文件：保留路径排序后的第 1 个，其余删除
    返回 (labels, relpaths)。
    """
    expected = expected_batches(count, batch_size)
    expected_set = set(expected)
    present = match_files_for_area(area, file_index)

    labels = []
    relpaths = []
    for rng, paths in sorted(present.items()):
        # 去重并稳定排序，便于“保留第一个”
        uniq = sorted(dict.fromkeys(paths))
        label = f'{rng[0]}-{rng[1]}'
        if rng not in expected_set:
            relpaths.extend(uniq)
            labels.append(label + (f'({len(uniq)})' if len(uniq) > 1 else ''))
        elif len(uniq) > 1:
            dupes = uniq[1:]
            relpaths.extend(dupes)
            labels.append(f'{label}重复{len(dupes)}')
    return labels, relpaths


def check_one_area(area: str, count: int, file_index: dict, batch_size: int = BATCH_SIZE):
    expected = expected_batches(count, batch_size)
    present = match_files_for_area(area, file_index)

    missing = []
    for rng in expected:
        if rng not in present or not present[rng]:
            missing.append(f'{rng[0]}-{rng[1]}')

    extras, _paths = collect_extra_paths(area, count, file_index, batch_size)

    if count <= 0:
        result = 'Count为0，无需文件'
    elif not missing and not extras:
        result = '完整'
    elif not missing and extras:
        result = '完整(有多余)'
    elif missing and not extras:
        result = '缺失: ' + ', '.join(missing)
    else:
        result = '缺失: ' + ', '.join(missing) + '；且有多余'

    extra_col = '; '.join(extras) if extras else ''
    return result, extra_col


def write_result_xlsx(path: str, rows):
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '核对结果'
    headers = ['Research Areas', 'Count', '结果', '多余文件']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for area, count, result, extra in rows:
        ws.append([area, count, result, extra])
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 40
    wb.save(path)


def resolve_save_dir(root: str):
    """定位含 ResearchAreas 的导出任务目录。"""
    root = os.path.abspath(root)
    if os.path.isdir(os.path.join(root, 'ResearchAreas')):
        return root
    # 若 root 即 ResearchAreas
    if os.path.basename(root.rstrip('\\/')) == 'ResearchAreas':
        return os.path.dirname(root)
    # 向下找一层
    try:
        for name in os.listdir(root):
            p = os.path.join(root, name)
            if os.path.isdir(os.path.join(p, 'ResearchAreas')):
                return p
    except OSError:
        pass
    return root


def parse_date_label(date_label: str):
    """2023-01-01_2023-12-31 → (2023-01-01, 2023-12-31)。"""
    text = (date_label or '').strip()
    parts = text.split('_')
    if len(parts) >= 2 and re.match(r'\d{4}-\d{2}-\d{2}$', parts[0]):
        return parts[0], parts[1]
    if re.match(r'\d{4}-\d{2}-\d{2}$', text):
        return text, text
    raise ValueError(f'无法解析日期段: {date_label!r}，请用 2023-01-01_2023-12-31')


def load_area_counts(args):
    """从核对结果或统计表加载 [(area, count), ...]（取前两列）。"""
    path = args.check or args.stats
    if not path:
        raise ValueError('请指定 --check 或 --stats')
    return read_stats_table(os.path.abspath(path))


def run_check(args):
    stats = read_stats_table(args.stats)
    if not stats:
        print('统计表无有效数据')
        return 1

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        print(f'导出目录不存在: {root}')
        return 1

    date_label = args.date_range.strip()
    print(f'统计表: {args.stats}（{len(stats)} 类）')
    print(f'导出目录: {root}')
    print(f'日期段: {date_label}')
    print(f'批次大小: {args.batch_size}')

    file_index = build_file_index(root, date_label)
    print(f'匹配到含该日期的导出文件研究方向数: {len(file_index)}')

    out_rows = []
    ok_n = miss_n = extra_n = 0
    for area, count in stats:
        result, extra = check_one_area(
            area, count, file_index, batch_size=args.batch_size)
        out_rows.append((area, count, result, extra))
        if result == '完整':
            ok_n += 1
        elif '缺失' in result:
            miss_n += 1
        if extra:
            extra_n += 1
        print(f'  [{result}] {area} ({count})'
              + (f'  多余={extra}' if extra else ''))

    out_path = os.path.abspath(args.out)
    os.makedirs(os.path.dirname(out_path) or '.', exist_ok=True)
    write_result_xlsx(out_path, out_rows)
    print(
        f'\n完成: 完整 {ok_n} / 缺失 {miss_n} / 有多余文件的类 {extra_n}'
        f'\n结果已写入: {out_path}')
    return 0


def _is_under_root(root: str, path: str) -> bool:
    root_n = os.path.normcase(os.path.abspath(root))
    path_n = os.path.normcase(os.path.abspath(path))
    try:
        common = os.path.normcase(os.path.commonpath([root_n, path_n]))
    except ValueError:
        return False
    return common == root_n


def run_delete_extra(args):
    """按 Count 计算应有批次，删除多余/重复导出文件。"""
    try:
        stats = load_area_counts(args)
    except Exception as e:
        print(e)
        return 1
    if not stats:
        print('无有效研究方向数据')
        return 1

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        print(f'导出目录不存在: {root}')
        return 1

    date_label = args.date_range.strip()
    src = args.check or args.stats
    print(f'数量来源: {src}（{len(stats)} 类）')
    print(f'导出目录: {root}')
    print(f'日期段: {date_label}')
    print(f'批次大小: {args.batch_size}')
    if args.dry_run:
        print('[--dry-run] 仅列出将删除的文件，不真正删除')

    file_index = build_file_index(root, date_label)
    print(f'匹配到含该日期的导出文件研究方向数: {len(file_index)}')

    plan = []  # (area, labels, abs_paths)
    for area, count in stats:
        labels, rels = collect_extra_paths(
            area, count, file_index, batch_size=args.batch_size)
        if not rels:
            continue
        abs_paths = []
        for rel in rels:
            abs_path = os.path.normpath(os.path.join(root, rel))
            if not _is_under_root(root, abs_path):
                print(f'  跳过路径越界: {rel}')
                continue
            if not os.path.isfile(abs_path):
                print(f'  文件已不存在: {rel}')
                continue
            abs_paths.append(abs_path)
        if abs_paths:
            plan.append((area, labels, abs_paths))

    if not plan:
        print('\n没有可删除的多余文件')
        return 0

    total_files = sum(len(p) for _, _, p in plan)
    print(f'\n待删除: {len(plan)} 个研究方向，共 {total_files} 个文件')
    for area, labels, paths in plan:
        print(f'  - {area}: {"; ".join(labels)}')
        for p in paths:
            print(f'      {p}')

    if args.dry_run:
        print(f'\n[--dry-run] 预览结束，未删除（共 {total_files} 个）')
        return 0

    deleted = 0
    failed = 0
    for area, _labels, paths in plan:
        for p in paths:
            try:
                os.remove(p)
                deleted += 1
                print(f'已删除: {p}')
            except OSError as e:
                failed += 1
                print(f'删除失败: {p} ({e})')

    print(f'\n删除完成: 成功 {deleted}，失败 {failed}')
    return 0 if failed == 0 else 1


def run_fill_missing(args):
    check_path = args.check or args.stats
    if not check_path:
        print('请指定核对结果文件 --check')
        return 1

    rows = read_check_result(check_path)
    jobs = [r for r in rows if r.get('batches')]
    if not jobs:
        print('核对结果中没有标记为缺失的批次，无需补导')
        return 0

    total_batches = sum(len(j['batches']) for j in jobs)
    print(f'核对结果: {check_path}')
    print(f'待补导: {len(jobs)} 个研究方向，共 {total_batches} 个批次')
    for j in jobs:
        batch_txt = ', '.join(f'{s}-{e}' for s, e in j['batches'])
        print(f'  - {j["name"]} (Count={j["count"]}): {batch_txt}')

    if args.dry_run:
        print('\n[--dry-run] 仅列出缺失批次，不启动浏览器')
        return 0

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        print(f'导出目录不存在: {root}')
        return 1
    save_dir = resolve_save_dir(root)
    print(f'导出任务目录: {save_dir}')

    date_label = args.date_range.strip()
    start_date, end_date = parse_date_label(date_label)

    # 合集：默认从目录名推断，或命令行指定
    collections = args.collections
    database = args.database or '所有数据库'
    if not collections:
        base = os.path.basename(os.path.dirname(save_dir.rstrip('\\/')))
        # 形如 所有数据库_ProQuest Dissertations & Theses Citation Index
        if '_' in base and base.startswith('所有数据库'):
            collections = base.split('_', 1)[1].strip() or None
        if not collections:
            collections = input(
                '请输入合集名 (如 ProQuest Dissertations & Theses Citation Index): '
            ).strip()
    if not collections:
        print('未指定合集，无法补导')
        return 1

    print(f'数据库: {database}')
    print(f'合集: {collections}')
    print(f'检索日期: {start_date} ~ {end_date}')

    # 导入自动化模块（与本脚本同目录）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    try:
        import sciforexcel as sci
    except ImportError as e:
        print('无法导入 sciforexcel.py:', e)
        return 1

    driver = None
    try:
        driver = sci.start_browser_session(save_dir, database, collections)
        sci.bind_browser_ctx(driver, save_dir, database, collections)
        if not sci.is_driver_alive(driver) or not sci.is_wos_logged_in(driver):
            print('登录失败')
            return 1
        n = sci.export_ra_missing_batches(
            driver, save_dir, date_label, start_date, end_date,
            database, collections, jobs)
        print(f'\n补导完成，成功批次: {n}')
        return 0 if n > 0 else 1
    except Exception as e:
        print('补导异常:', e)
        return 1
    finally:
        try:
            sci.quit_chrome_driver(driver)
        except Exception:
            pass


def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description='核对 Research Areas 导出批次，或补导缺失 / 删除多余文件')
    mode = p.add_mutually_exclusive_group()
    mode.add_argument(
        '--fill-missing', action='store_true',
        help='根据核对结果仅补导缺失批次（需浏览器登录 WoS）')
    mode.add_argument(
        '--delete-extra', action='store_true',
        help='删除多余/重复批次文件（按 Count 计算应有批次）')
    p.add_argument(
        '--dry-run', action='store_true',
        help='补导/删除前只打印清单，不实际执行')
    p.add_argument(
        '--check', '-c',
        help='核对结果 xlsx（补导/删除模式；列：Research Areas/Count/结果/多余文件）')
    p.add_argument(
        '--stats', '-s',
        help='统计表路径（核对模式；第1列 Research Areas，第2列 Count）')
    p.add_argument(
        '--root', '-r',
        help='导出根目录（含 ResearchAreas 的任务目录或其上级）')
    p.add_argument(
        '--date-range', '-d',
        help='日期段标签，如 2023-01-01_2023-12-31')
    p.add_argument(
        '--batch-size', type=int, default=BATCH_SIZE,
        help=f'每批条数，默认 {BATCH_SIZE}')
    p.add_argument(
        '--out', '-o',
        help='核对结果输出路径（默认与统计表同目录）')
    p.add_argument(
        '--database',
        default='所有数据库',
        help='检索数据库，默认「所有数据库」')
    p.add_argument(
        '--collections',
        help='合集名；默认尝试从导出目录名推断')
    return p.parse_args(argv)


def prompt_check(args):
    if not args.stats:
        args.stats = input(
            '统计表路径 (Research Areas / Count): ').strip().strip('"')
    if not args.root:
        args.root = input('导出根目录: ').strip().strip('"')
    if not args.date_range:
        args.date_range = input(
            '日期段标签 (如 2023-01-01_2023-12-31): ').strip()
    if not args.out:
        base = os.path.splitext(os.path.abspath(args.stats))[0]
        args.out = f'{base}_核对_{args.date_range}.xlsx'
    return args


def prompt_fill(args):
    if not args.check and not args.stats:
        args.check = input(
            '核对结果 xlsx 路径: ').strip().strip('"')
    elif args.stats and not args.check:
        args.check = args.stats
    if not args.root:
        args.root = input('导出根目录: ').strip().strip('"')
    if not args.date_range:
        args.date_range = input(
            '日期段标签 (如 2023-01-01_2023-12-31): ').strip()
    return args


def prompt_delete(args):
    if not args.check and not args.stats:
        args.check = input(
            '核对结果或统计表 xlsx 路径: ').strip().strip('"')
    if not args.root:
        args.root = input('导出根目录: ').strip().strip('"')
    if not args.date_range:
        args.date_range = input(
            '日期段标签 (如 2023-01-01_2023-12-31): ').strip()
    return args


def main(argv=None):
    args = parse_args(argv)
    if args.fill_missing:
        if not (args.check or args.stats) or not args.root or not args.date_range:
            args = prompt_fill(args)
        return run_fill_missing(args)

    if args.delete_extra:
        if not (args.check or args.stats) or not args.root or not args.date_range:
            args = prompt_delete(args)
        return run_delete_extra(args)

    if not (args.stats and args.root and args.date_range):
        args = prompt_check(args)
    return run_check(args)


if __name__ == '__main__':
    sys.exit(main() or 0)

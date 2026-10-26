"""合并「九类期刊题录」各子文件夹中的期刊 Excel 表。

目录约定：
  {根目录}/{领域}/{刊名}/SCI_*.xls(x)

只读取原表，不修改、不覆盖原文件。合并结果写到独立目录。

默认：
  1. 将同一刊名文件夹内的分批表合并为一张「按刊」表
  2. 再将同一领域下各刊合并为一张「按领域」表
  3. 暂不合并「生物医药」（可用 --skip 调整）

推荐依赖：pandas、xlsxwriter、python-calamine（读写更快；未安装时回退 openpyxl/xlrd）

示例：
  python merge_journal_records.py
  python merge_journal_records.py --no-domain
  python merge_journal_records.py --all
  python merge_journal_records.py --keep-dups --resume
  python merge_journal_records.py --skip 生物医药
  python merge_journal_records.py --skip
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

DEFAULT_SRC = Path(r"D:\2026年工作\SCI数据库下载\九类期刊题录")
DEFAULT_OUT = Path(r"D:\2026年工作\SCI数据库下载\九类期刊题录_合并")
DEFAULT_SKIP = ("生物医药",)

EXCEL_SUFFIXES = {".xlsx", ".xls"}
RANGE_RE = re.compile(r"_(\d+)-(\d+)\.(xlsx|xls)$", re.I)
INVALID_NAME_RE = re.compile(r'[<>:"/\\|?*]')

ID_COL_CANDIDATES = (
    "UT (Unique WOS ID)",
    "UT",
    "Accession Number",
    "DOI",
)

_HAS_CALAMINE = False
try:
    import python_calamine  # noqa: F401

    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False


def read_table(path: Path) -> pd.DataFrame:
    """只读打开 Excel，不写入源文件。"""
    suffix = path.suffix.lower()
    if _HAS_CALAMINE:
        return pd.read_excel(path, engine="calamine")
    engine = "openpyxl" if suffix == ".xlsx" else "xlrd"
    return pd.read_excel(path, engine=engine)


def is_excel(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
        and not path.name.startswith(".")
    )


def file_sort_key(path: Path) -> tuple[int, int, str]:
    match = RANGE_RE.search(path.name)
    if match:
        return (int(match.group(1)), int(match.group(2)), path.name.lower())
    return (10**12, 0, path.name.lower())


def safe_filename(name: str) -> str:
    text = INVALID_NAME_RE.sub("_", name).strip(" .")
    return text or "untitled"


def find_id_col(columns) -> str | None:
    colset = list(columns)
    for name in ID_COL_CANDIDATES:
        if name in colset:
            return name
    lowered = {str(col).strip().lower(): col for col in colset}
    for key in ("ut (unique wos id)", "ut", "accession number", "doi"):
        if key in lowered:
            return lowered[key]
    return None


def collect_excel_files(journal_dir: Path) -> list[Path]:
    files = [p for p in journal_dir.iterdir() if is_excel(p)]
    files.sort(key=file_sort_key)
    return files


def write_excel(df: pd.DataFrame, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.stem + ".tmp.xlsx")
    try:
        df.to_excel(tmp, index=False, engine="xlsxwriter")
        if dest.exists():
            dest.unlink()
        os.replace(tmp, dest)
    except Exception:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
        raise


def drop_duplicates(df: pd.DataFrame) -> tuple[pd.DataFrame, int, str]:
    col = find_id_col(df.columns)
    if col is None:
        return df, 0, ""
    before = len(df)
    key = df[col].astype(str).str.strip()
    valid = key.ne("") & key.str.lower().ne("nan") & key.str.lower().ne("none")
    keep = ~valid | ~key.duplicated(keep="first")
    out = df.loc[keep].reset_index(drop=True)
    return out, before - len(out), col


def merge_frames(
    files: list[Path],
    domain: str,
    journal: str,
) -> tuple[pd.DataFrame, list[str]]:
    frames: list[pd.DataFrame] = []
    errors: list[str] = []
    for path in files:
        try:
            table = read_table(path)
        except Exception as exc:
            errors.append(f"{path}: {exc}")
            continue
        if table.empty:
            continue
        table.insert(0, "来源文件", path.name)
        table.insert(0, "期刊文件夹", journal)
        table.insert(0, "领域", domain)
        frames.append(table)
    if not frames:
        return pd.DataFrame(), errors
    return pd.concat(frames, ignore_index=True, sort=False), errors


def list_domains(src: Path) -> list[Path]:
    return sorted([p for p in src.iterdir() if p.is_dir() and not p.name.startswith(".")])


def list_journals(domain_dir: Path) -> list[Path]:
    return sorted(
        [p for p in domain_dir.iterdir() if p.is_dir() and not p.name.startswith(".")]
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="合并九类期刊题录各刊 Excel 表")
    parser.add_argument("--src", type=Path, default=DEFAULT_SRC, help="题录根目录")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help="合并结果目录")
    parser.add_argument(
        "--no-domain",
        action="store_true",
        help="只按刊合并，不生成各领域总表",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="额外生成九类全部合并表",
    )
    parser.add_argument(
        "--keep-dups",
        action="store_true",
        help="保留重复记录（默认按 UT / Accession Number / DOI 去重）",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="跳过已存在的按刊合并文件",
    )
    parser.add_argument(
        "--skip",
        nargs="*",
        default=list(DEFAULT_SKIP),
        metavar="领域",
        help="跳过的领域名；默认跳过生物医药。传入 --skip 且不跟名称则不跳过任何领域",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    src = args.src
    out = args.out
    if not src.is_dir():
        print(f"源目录不存在: {src}", file=sys.stderr)
        return 1

    skip = {name.strip() for name in args.skip if str(name).strip()}
    print(f"读取引擎: {'calamine' if _HAS_CALAMINE else 'openpyxl/xlrd'}")
    print(f"源目录: {src}")
    print(f"输出目录: {out}")
    if skip:
        print(f"跳过领域: {', '.join(sorted(skip))}")

    summary_rows: list[dict] = []
    all_domain_paths: list[Path] = []

    for domain_dir in list_domains(src):
        domain = domain_dir.name
        if domain in skip:
            print(f"\n=== {domain} ===")
            print("  按设置跳过，不合并")
            summary_rows.append(
                {
                    "领域": domain,
                    "期刊": "",
                    "源文件数": 0,
                    "合并前行数": 0,
                    "去重行数": 0,
                    "合并后行数": 0,
                    "去重列": "",
                    "状态": "跳过领域",
                    "输出": "",
                    "错误": "",
                }
            )
            continue
        domain_out = out / domain
        domain_frames: list[pd.DataFrame] = []
        print(f"\n=== {domain} ===")

        for journal_dir in list_journals(domain_dir):
            journal = journal_dir.name
            files = collect_excel_files(journal_dir)
            dest = domain_out / f"{safe_filename(journal)}.xlsx"
            row = {
                "领域": domain,
                "期刊": journal,
                "源文件数": len(files),
                "合并前行数": 0,
                "去重行数": 0,
                "合并后行数": 0,
                "去重列": "",
                "状态": "",
                "输出": "",
                "错误": "",
            }

            if not files:
                row["状态"] = "无表格"
                summary_rows.append(row)
                print(f"  跳过（无表格）: {journal}")
                continue

            if args.resume and dest.exists() and dest.stat().st_size > 0:
                row["状态"] = "已存在"
                row["输出"] = str(dest)
                summary_rows.append(row)
                print(f"  跳过（已存在）: {journal}")
                if not args.no_domain:
                    try:
                        domain_frames.append(read_table(dest))
                    except Exception as exc:
                        print(f"    读取已有合并表失败: {exc}")
                continue

            merged, errors = merge_frames(files, domain, journal)
            if errors:
                row["错误"] = " | ".join(errors)
                for err in errors:
                    print(f"    读取失败: {err}")
            if merged.empty:
                row["状态"] = "空表"
                summary_rows.append(row)
                print(f"  空表: {journal} ({len(files)} 个文件)")
                continue

            row["合并前行数"] = len(merged)
            dropped = 0
            id_col = ""
            if not args.keep_dups:
                merged, dropped, id_col = drop_duplicates(merged)
            row["去重行数"] = dropped
            row["去重列"] = id_col
            row["合并后行数"] = len(merged)
            write_excel(merged, dest)
            row["状态"] = "完成"
            row["输出"] = str(dest)
            summary_rows.append(row)
            print(
                f"  {journal}: {len(files)} 文件, {row['合并前行数']} -> "
                f"{row['合并后行数']} 行"
                + (f"（去重 {dropped}）" if dropped else "")
            )
            if not args.no_domain:
                domain_frames.append(merged)

        if args.no_domain or not domain_frames:
            continue
        domain_df = pd.concat(domain_frames, ignore_index=True, sort=False)
        dropped = 0
        id_col = ""
        before = len(domain_df)
        if not args.keep_dups:
            domain_df, dropped, id_col = drop_duplicates(domain_df)
        domain_dest = domain_out / f"{safe_filename(domain)}_全部.xlsx"
        write_excel(domain_df, domain_dest)
        all_domain_paths.append(domain_dest)
        summary_rows.append(
            {
                "领域": domain,
                "期刊": f"{domain}_全部",
                "源文件数": sum(
                    1
                    for r in summary_rows
                    if r["领域"] == domain and r["期刊"] != f"{domain}_全部" and r["状态"] == "完成"
                ),
                "合并前行数": before,
                "去重行数": dropped,
                "合并后行数": len(domain_df),
                "去重列": id_col,
                "状态": "完成",
                "输出": str(domain_dest),
                "错误": "",
            }
        )
        print(
            f"  >> {domain} 总表: {before} -> {len(domain_df)} 行 -> {domain_dest.name}"
        )
        del domain_frames, domain_df

    if args.all and all_domain_paths:
        frames = []
        for path in all_domain_paths:
            frames.append(read_table(path))
        all_df = pd.concat(frames, ignore_index=True, sort=False)
        before = len(all_df)
        dropped = 0
        id_col = ""
        if not args.keep_dups:
            all_df, dropped, id_col = drop_duplicates(all_df)
        all_dest = out / "九类期刊_全部.xlsx"
        write_excel(all_df, all_dest)
        summary_rows.append(
            {
                "领域": "全部",
                "期刊": "九类期刊_全部",
                "源文件数": len(all_domain_paths),
                "合并前行数": before,
                "去重行数": dropped,
                "合并后行数": len(all_df),
                "去重列": id_col,
                "状态": "完成",
                "输出": str(all_dest),
                "错误": "",
            }
        )
        print(f"\n九类总表: {before} -> {len(all_df)} 行 -> {all_dest}")

    if summary_rows:
        summary = pd.DataFrame(summary_rows)
        write_excel(summary, out / "合并摘要.xlsx")
        done = sum(1 for r in summary_rows if r["状态"] == "完成")
        skipped = sum(
            1 for r in summary_rows if r["状态"] in {"无表格", "空表", "已存在", "跳过领域"}
        )
        print(f"\n完成 {done} 个输出，跳过 {skipped} 项。摘要: {out / '合并摘要.xlsx'}")
    else:
        print("未找到可合并的期刊表格。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

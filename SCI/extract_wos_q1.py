"""根据 Q1.xlsx 的中科院 1 区期刊列表，从 WoS 核心合集导出表中筛选 1 区记录。

只读取原表，不修改、不覆盖原文件。提取结果按文件名中的日期分文件夹存放。

推荐依赖：pandas、xlsxwriter、python-calamine（读写更快；未安装时回退 openpyxl/xlrd）

示例：
  python extract_wos_q1.py
  python extract_wos_q1.py --workers 4 --resume
  python extract_wos_q1.py --dates 2023-01-02,2023-01-03
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
import warnings
from collections import Counter, defaultdict
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

DEFAULT_Q1 = Path(__file__).resolve().parent / "Q1.xlsx"
DEFAULT_SRC = Path(r"D:\2026年工作\SCI数据库下载\Web of Science 核心合集_All")
DEFAULT_OUT = Path(r"D:\2026年工作\SCI数据库下载\Web of Science 核心合集_All_Q1")

DATE_RE = re.compile(r"SCI_(\d{4}-\d{2}-\d{2})_")
EXCEL_SUFFIXES = {".xlsx", ".xls"}
SOURCE_COL_CANDIDATES = ("Source Title", "source title", "SOURCE TITLE", "SO")

_HAS_CALAMINE = False
try:
    import python_calamine  # noqa: F401

    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False


def normalize_journal(name: object) -> str:
    """统一刊名：全角、大小写、&/AND、标点与多余空格。"""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    text = unicodedata.normalize("NFKC", str(name)).upper().strip()
    if not text or text == "NAN":
        return ""
    text = text.replace("&", " AND ")
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if text.startswith("THE "):
        text = text[4:]
    return text


def load_q1_names(q1_path: Path) -> set[str]:
    df = pd.read_excel(q1_path)
    if "Journal" not in df.columns:
        raise SystemExit(f"Q1 表缺少 Journal 列: {q1_path}")
    if "分区" in df.columns:
        df = df[df["分区"] == 1]
    names = {normalize_journal(x) for x in df["Journal"]}
    names.discard("")
    if not names:
        raise SystemExit(f"未从 Q1 表读到期刊名: {q1_path}")
    return names


def find_source_col(columns) -> str | None:
    colset = list(columns)
    for name in SOURCE_COL_CANDIDATES:
        if name in colset:
            return name
    for col in colset:
        if str(col).strip().lower() == "source title":
            return col
    return None


def read_table(path: Path) -> pd.DataFrame:
    """只读打开 Excel，不写入源文件。"""
    suffix = path.suffix.lower()
    if _HAS_CALAMINE:
        return pd.read_excel(path, engine="calamine")
    engine = "openpyxl" if suffix == ".xlsx" else "xlrd"
    return pd.read_excel(path, engine=engine)


def collect_source_files(src_dir: Path, out_dir: Path) -> dict[str, list[Path]]:
    grouped: dict[str, list[Path]] = defaultdict(list)
    out_resolved = out_dir.resolve()
    for path in src_dir.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in EXCEL_SUFFIXES:
            continue
        try:
            if out_resolved in path.resolve().parents or path.resolve() == out_resolved:
                continue
        except OSError:
            continue
        match = DATE_RE.search(path.name)
        if not match:
            continue
        grouped[match.group(1)].append(path)
    for date in grouped:
        grouped[date].sort()
    return dict(sorted(grouped.items()))


def process_one_date(
    date: str,
    files: list[str],
    q1_names: set[str],
    out_dir: str,
    resume: bool,
) -> dict:
    out_root = Path(out_dir)
    dest_dir = out_root / date
    dest = dest_dir / f"Q1_{date}.xlsx"
    if resume and dest.exists() and dest.stat().st_size > 0:
        return {
            "date": date,
            "n_files": len(files),
            "n_rows": None,
            "n_q1": None,
            "status": "skipped",
            "output": str(dest),
            "errors": [],
            "unmatched": {},
        }

    frames: list[pd.DataFrame] = []
    errors: list[str] = []
    unmatched: Counter[str] = Counter()
    n_rows = 0
    source_col_used = None

    for file_str in files:
        path = Path(file_str)
        try:
            df = read_table(path)
        except Exception as exc:
            errors.append(f"{path.name}: 读取失败 {exc}")
            continue
        n_rows += len(df)
        col = find_source_col(df.columns)
        if col is None:
            errors.append(f"{path.name}: 未找到 Source Title 列")
            continue
        source_col_used = col
        keys = df[col].map(normalize_journal)
        hit = keys.isin(q1_names)
        miss_titles = df.loc[~hit, col].dropna().astype(str).str.strip()
        miss_titles = miss_titles[miss_titles != ""]
        if not miss_titles.empty:
            unmatched.update(miss_titles.value_counts().to_dict())
        part = df.loc[hit].copy()
        if not part.empty:
            frames.append(part)

    if not frames:
        return {
            "date": date,
            "n_files": len(files),
            "n_rows": n_rows,
            "n_q1": 0,
            "status": "empty",
            "output": "",
            "errors": errors,
            "unmatched": dict(unmatched),
        }

    out_df = pd.concat(frames, ignore_index=True, sort=False)
    dest_dir.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.stem + ".tmp.xlsx")
    try:
        out_df.to_excel(tmp, index=False, engine="xlsxwriter")
        if dest.exists():
            dest.unlink()
        os.replace(tmp, dest)
    except Exception:
        if tmp.exists():
            tmp.unlink(missing_ok=True)
        raise

    return {
        "date": date,
        "n_files": len(files),
        "n_rows": n_rows,
        "n_q1": int(len(out_df)),
        "status": "ok",
        "output": str(dest),
        "errors": errors,
        "unmatched": dict(unmatched),
        "source_col": source_col_used,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按中科院 1 区期刊筛选 WoS 导出表（不修改原表）")
    parser.add_argument("--q1", type=Path, default=DEFAULT_Q1, help="1 区期刊表 Q1.xlsx")
    parser.add_argument("--src", type=Path, default=DEFAULT_SRC, help="WoS 原始表格根目录（只读）")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help="提取结果输出目录")
    parser.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 4) // 2))
    parser.add_argument("--resume", action="store_true", help="跳过已生成的日期文件")
    parser.add_argument("--dates", type=str, default="", help="仅处理指定日期，逗号分隔")
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    args = parse_args()
    q1_path = args.q1.resolve()
    src_dir = args.src.resolve()
    out_dir = args.out.resolve()

    if not q1_path.exists():
        print(f"找不到 Q1 表: {q1_path}", file=sys.stderr)
        return 1
    if not src_dir.is_dir():
        print(f"找不到源目录: {src_dir}", file=sys.stderr)
        return 1
    if out_dir == src_dir or src_dir in out_dir.parents:
        # 允许输出在源目录旁，但不允许写进源目录内部，以免扫到结果文件或误改原表
        print(f"输出目录不能位于源目录内部: {out_dir}", file=sys.stderr)
        return 1

    q1_names = load_q1_names(q1_path)
    grouped = collect_source_files(src_dir, out_dir)
    if args.dates:
        wanted = {x.strip() for x in args.dates.split(",") if x.strip()}
        grouped = {k: v for k, v in grouped.items() if k in wanted}
        missing = wanted - set(grouped)
        if missing:
            print("未找到这些日期的源文件:", ", ".join(sorted(missing)))

    if not grouped:
        print("没有可处理的源表格")
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    n_files = sum(len(v) for v in grouped.values())
    print(f"Q1 期刊: {len(q1_names)}")
    print(f"源目录: {src_dir}（只读）")
    print(f"输出目录: {out_dir}")
    print(f"日期数: {len(grouped)}，源文件: {n_files}")
    print(f"Excel 引擎: {'calamine' if _HAS_CALAMINE else 'openpyxl/xlrd'}")
    print(f"进程数: {args.workers}")

    tasks = [
        (date, [str(p) for p in files], q1_names, str(out_dir), bool(args.resume))
        for date, files in grouped.items()
    ]

    results: list[dict] = []
    unmatched_all: Counter[str] = Counter()
    done = 0

    if args.workers <= 1:
        for task in tasks:
            result = process_one_date(*task)
            results.append(result)
            unmatched_all.update(result.get("unmatched") or {})
            done += 1
            print(
                f"[{done}/{len(tasks)}] {result['date']} "
                f"{result['status']} 源行={result['n_rows']} Q1={result['n_q1']}",
                flush=True,
            )
            for err in result.get("errors") or []:
                print("  !", err)
    else:
        with ProcessPoolExecutor(max_workers=args.workers) as pool:
            futures = {pool.submit(process_one_date, *task): task[0] for task in tasks}
            for fut in as_completed(futures):
                date = futures[fut]
                try:
                    result = fut.result()
                except Exception as exc:
                    result = {
                        "date": date,
                        "n_files": len(grouped[date]),
                        "n_rows": None,
                        "n_q1": None,
                        "status": "failed",
                        "output": "",
                        "errors": [str(exc)],
                        "unmatched": {},
                    }
                results.append(result)
                unmatched_all.update(result.get("unmatched") or {})
                done += 1
                print(
                    f"[{done}/{len(tasks)}] {result['date']} "
                    f"{result['status']} 源行={result['n_rows']} Q1={result['n_q1']}",
                    flush=True,
                )
                for err in result.get("errors") or []:
                    print("  !", err)

    results.sort(key=lambda x: x["date"])
    summary = pd.DataFrame(
        [
            {
                "日期": r["date"],
                "源文件数": r["n_files"],
                "源记录数": r["n_rows"],
                "1区记录数": r["n_q1"],
                "状态": r["status"],
                "输出文件": r["output"],
                "错误": "；".join(r.get("errors") or []),
            }
            for r in results
        ]
    )
    summary_path = out_dir / "_提取汇总.xlsx"
    if summary_path.exists() and args.dates:
        old = pd.read_excel(summary_path)
        old = old[~old["日期"].astype(str).isin(summary["日期"].astype(str))]
        summary = pd.concat([old, summary], ignore_index=True)
        summary = summary.sort_values("日期").reset_index(drop=True)
    summary.to_excel(summary_path, index=False, engine="xlsxwriter")

    unmatched_path = None
    if unmatched_all and not args.dates:
        unmatched_df = pd.DataFrame(
            [{"Source Title": k, "出现次数": v} for k, v in unmatched_all.most_common()]
        )
        unmatched_path = out_dir / "_未匹配刊名.xlsx"
        unmatched_df.to_excel(unmatched_path, index=False, engine="xlsxwriter")

    n_ok = sum(1 for r in results if r["status"] == "ok")
    n_q1 = sum(int(r["n_q1"] or 0) for r in results)
    n_fail = sum(1 for r in results if r["status"] == "failed")
    print(f"完成：成功日期 {n_ok}，失败 {n_fail}，1 区记录合计 {n_q1}")
    print(f"汇总: {summary_path}")
    if unmatched_path:
        print(f"未匹配刊名: {unmatched_path}")
    return 0 if n_fail == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())

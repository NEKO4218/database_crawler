"""按 Q1 九类期刊表，从已提取的 Q1 数据中筛选 9 类记录。

只读取 Q1 结果表，不修改原文件。输出按领域分子文件夹、再按日期存表。

示例：
  python extract_wos_q1_domains.py
  python extract_wos_q1_domains.py --workers 8 --resume
  python extract_wos_q1_domains.py --dates 2023-01-01,2023-01-02
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
import warnings
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

DEFAULT_JOURNALS = Path(r"D:\2026年工作\SCI数据库下载\Web of Science 核心合集_All_Q1\Q1_Journals_九类.xlsx")
DEFAULT_SRC = Path(r"D:\2026年工作\SCI数据库下载\Web of Science 核心合集_All_Q1")
DEFAULT_OUT = Path(r"D:\2026年工作\SCI数据库下载\Web of Science 核心合集_All_Q1_九类")

DOMAINS = [
    "集成电路",
    "高端科学仪器",
    "智能制造",
    "高端机器人",
    "人工智能",
    "智能网联新能源车",
    "商业航天",
    "生物医药",
    "生物育种",
]
DATE_RE = re.compile(r"Q1_(\d{4}-\d{2}-\d{2})\.xlsx$", re.I)
SOURCE_COL_CANDIDATES = ("Source Title", "source title", "SOURCE TITLE", "SO")

_HAS_CALAMINE = False
try:
    import python_calamine  # noqa: F401

    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False


def normalize_journal(name: object) -> str:
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


def load_journal_map(path: Path) -> dict[str, str]:
    mapping: dict[str, str] = {}
    xl = pd.ExcelFile(path)
    for domain in DOMAINS:
        if domain not in xl.sheet_names:
            continue
        df = pd.read_excel(xl, sheet_name=domain)
        if "Journal" not in df.columns:
            continue
        for name in df["Journal"]:
            key = normalize_journal(name)
            if key:
                mapping[key] = domain
    if not mapping:
        raise SystemExit(f"未从九类期刊表读到刊名: {path}")
    return mapping


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
    if _HAS_CALAMINE:
        return pd.read_excel(path, engine="calamine")
    return pd.read_excel(path, engine="openpyxl")


def collect_date_files(src_dir: Path, out_dir: Path) -> dict[str, Path]:
    grouped: dict[str, Path] = {}
    out_resolved = out_dir.resolve()
    for path in src_dir.rglob("Q1_*.xlsx"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        try:
            if out_resolved in path.resolve().parents:
                continue
        except OSError:
            continue
        match = DATE_RE.search(path.name)
        if not match:
            continue
        grouped[match.group(1)] = path
    return dict(sorted(grouped.items()))


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


def process_one_date(
    date: str,
    file_str: str,
    journal_map: dict[str, str],
    out_dir: str,
    resume: bool,
) -> dict:
    out_root = Path(out_dir)
    done_flag = out_root / "_done" / f"{date}.ok"
    if resume and done_flag.exists():
        return {
            "date": date,
            "n_rows": None,
            "status": "skipped",
            "errors": [],
            "counts": {d: None for d in DOMAINS},
        }

    errors: list[str] = []
    counts = {d: 0 for d in DOMAINS}
    path = Path(file_str)
    try:
        df = read_table(path)
    except Exception as exc:
        return {
            "date": date,
            "n_rows": None,
            "status": "failed",
            "errors": [f"{path.name}: 读取失败 {exc}"],
            "counts": counts,
        }

    col = find_source_col(df.columns)
    if col is None:
        return {
            "date": date,
            "n_rows": len(df),
            "status": "failed",
            "errors": [f"{path.name}: 未找到 Source Title 列"],
            "counts": counts,
        }

    domains = df[col].map(normalize_journal).map(journal_map)
    n_hit = 0
    for domain in DOMAINS:
        part = df.loc[domains == domain]
        counts[domain] = int(len(part))
        if part.empty:
            continue
        n_hit += len(part)
        dest = out_root / domain / f"{date}.xlsx"
        write_excel(part.copy(), dest)

    done_flag.parent.mkdir(parents=True, exist_ok=True)
    done_flag.write_text(str(n_hit), encoding="utf-8")
    return {
        "date": date,
        "n_rows": int(len(df)),
        "n_hit": int(n_hit),
        "status": "ok" if n_hit else "empty",
        "errors": errors,
        "counts": counts,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按九类期刊筛选 Q1 数据表（不修改原表）")
    parser.add_argument("--journals", type=Path, default=DEFAULT_JOURNALS)
    parser.add_argument("--src", type=Path, default=DEFAULT_SRC)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    parser.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 4) // 2))
    parser.add_argument("--resume", action="store_true")
    parser.add_argument("--dates", type=str, default="")
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    src_dir = args.src.resolve()
    out_dir = args.out.resolve()
    if not src_dir.is_dir():
        print(f"找不到 Q1 目录: {src_dir}", file=sys.stderr)
        return 1
    if out_dir == src_dir or src_dir in out_dir.parents:
        print(f"输出目录不能位于源目录内部: {out_dir}", file=sys.stderr)
        return 1

    journal_map = load_journal_map(args.journals.resolve())
    grouped = collect_date_files(src_dir, out_dir)
    if args.dates:
        wanted = {x.strip() for x in args.dates.split(",") if x.strip()}
        grouped = {k: v for k, v in grouped.items() if k in wanted}

    if not grouped:
        print("没有可处理的 Q1 表格")
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    for domain in DOMAINS:
        (out_dir / domain).mkdir(parents=True, exist_ok=True)

    print(f"九类期刊: {len(journal_map)}")
    print(f"源目录: {src_dir}（只读）")
    print(f"输出目录: {out_dir}")
    print(f"日期数: {len(grouped)}")
    print(f"Excel 引擎: {'calamine' if _HAS_CALAMINE else 'openpyxl'}")
    print(f"进程数: {args.workers}")

    tasks = [
        (date, str(path), journal_map, str(out_dir), bool(args.resume))
        for date, path in grouped.items()
    ]
    results: list[dict] = []
    done = 0

    if args.workers <= 1:
        for task in tasks:
            result = process_one_date(*task)
            results.append(result)
            done += 1
            print(
                f"[{done}/{len(tasks)}] {result['date']} {result['status']} "
                f"源行={result.get('n_rows')} 命中={result.get('n_hit')}",
                flush=True,
            )
            for err in result.get("errors") or []:
                print("  !", err, flush=True)
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
                        "n_rows": None,
                        "n_hit": 0,
                        "status": "failed",
                        "errors": [str(exc)],
                        "counts": {d: 0 for d in DOMAINS},
                    }
                results.append(result)
                done += 1
                print(
                    f"[{done}/{len(tasks)}] {result['date']} {result['status']} "
                    f"源行={result.get('n_rows')} 命中={result.get('n_hit')}",
                    flush=True,
                )
                for err in result.get("errors") or []:
                    print("  !", err, flush=True)

    results.sort(key=lambda x: x["date"])
    rows = []
    for r in results:
        rec = {
            "日期": r["date"],
            "源记录数": r.get("n_rows"),
            "命中合计": r.get("n_hit"),
            "状态": r["status"],
            "错误": "；".join(r.get("errors") or []),
        }
        counts = r.get("counts") or {}
        for domain in DOMAINS:
            rec[domain] = counts.get(domain)
        rows.append(rec)
    summary = pd.DataFrame(rows)
    totals = {"日期": "合计", "源记录数": summary["源记录数"].sum(min_count=1), "命中合计": summary["命中合计"].sum(min_count=1), "状态": "", "错误": ""}
    for domain in DOMAINS:
        totals[domain] = summary[domain].sum(min_count=1)
    summary = pd.concat([summary, pd.DataFrame([totals])], ignore_index=True)
    summary_path = out_dir / "_提取汇总.xlsx"
    summary.to_excel(summary_path, index=False, engine="xlsxwriter")

    n_fail = sum(1 for r in results if r["status"] == "failed")
    n_hit = sum(int(r.get("n_hit") or 0) for r in results)
    print(f"完成：失败 {n_fail}，九类记录合计 {n_hit}")
    print(f"汇总: {summary_path}")
    return 0 if n_fail == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())

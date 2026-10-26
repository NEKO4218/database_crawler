"""从已提取的 Q1 表中汇总 WoS Categories。

只读取 Q1 结果表，不修改原文件。多类别按分号拆分后去重，
在 Q1 目录下输出一张类别表。

示例：
  python extract_wos_categories.py
"""

from __future__ import annotations

import argparse
import sys
import warnings
from collections import Counter
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=FutureWarning)

DEFAULT_Q1_DIR = Path(r"D:\2026年工作\SCI数据库下载\Web of Science 核心合集_All_Q1")
OUTPUT_NAME = "WoS_Categories.xlsx"
CATEGORY_COL_CANDIDATES = ("WoS Categories", "Web of Science Categories", "WC")

_HAS_CALAMINE = False
try:
    import python_calamine  # noqa: F401

    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False


def find_category_col(columns) -> str | None:
    colset = list(columns)
    for name in CATEGORY_COL_CANDIDATES:
        if name in colset:
            return name
    for col in colset:
        if str(col).strip().lower() in {"wos categories", "web of science categories"}:
            return col
    return None


def read_categories(path: Path) -> pd.Series:
    kwargs = {"engine": "calamine"} if _HAS_CALAMINE else {"engine": "openpyxl"}
    try:
        df = pd.read_excel(path, usecols=["WoS Categories"], **kwargs)
        return df["WoS Categories"]
    except Exception:
        df = pd.read_excel(path, **kwargs)
        col = find_category_col(df.columns)
        if col is None:
            raise KeyError("未找到 WoS Categories 列")
        return df[col]


def split_categories(value: object) -> list[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return []
    parts = []
    for item in text.split(";"):
        name = " ".join(item.split()).strip()
        if name:
            parts.append(name)
    return parts


def collect_q1_files(q1_dir: Path) -> list[Path]:
    files = [
        p
        for p in q1_dir.rglob("Q1_*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    return sorted(files)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从 Q1 表提取并拆分 WoS Categories")
    parser.add_argument("--q1-dir", type=Path, default=DEFAULT_Q1_DIR, help="Q1 结果根目录")
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="输出表格路径，默认写到 Q1 目录下的 WoS_Categories.xlsx",
    )
    return parser.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    q1_dir = args.q1_dir.resolve()
    out_path = (args.out or (q1_dir / OUTPUT_NAME)).resolve()

    if not q1_dir.is_dir():
        print(f"找不到 Q1 目录: {q1_dir}", file=sys.stderr)
        return 1

    files = collect_q1_files(q1_dir)
    if not files:
        print(f"未找到 Q1_*.xlsx: {q1_dir}", file=sys.stderr)
        return 1

    print(f"Q1 目录: {q1_dir}（只读）")
    print(f"待处理表格: {len(files)}")
    print(f"输出: {out_path}")

    counts: Counter[str] = Counter()
    n_rows = 0
    n_split = 0
    errors: list[str] = []

    for i, path in enumerate(files, 1):
        try:
            series = read_categories(path)
        except Exception as exc:
            errors.append(f"{path.name}: 读取失败 {exc}")
            continue
        n_rows += len(series)
        for value in series.tolist():
            cats = split_categories(value)
            if len(cats) > 1:
                n_split += 1
            counts.update(cats)
        if i % 200 == 0 or i == len(files):
            print(f"[{i}/{len(files)}] 已累计类别 {len(counts)}", flush=True)

    if not counts:
        print("没有提取到任何类别", file=sys.stderr)
        return 1

    table = pd.DataFrame(
        [{"WoS Categories": name, "出现次数": n} for name, n in counts.most_common()]
    )
    table.insert(0, "序号", range(1, len(table) + 1))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    table.to_excel(out_path, index=False, engine="xlsxwriter")

    print(f"源记录: {n_rows}")
    print(f"含多类已拆分的记录: {n_split}")
    print(f"去重后类别数: {len(table)}")
    print(f"已保存: {out_path}")
    for err in errors:
        print(" !", err)
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())

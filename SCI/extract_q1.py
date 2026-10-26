"""从 2025 中科院期刊分区表中提取 1 区期刊，另存为 Q1.xlsx。"""

from pathlib import Path

import pandas as pd

SRC = Path(__file__).resolve().parent / "2025中科院期刊分区表excel完整版.xlsx"
DST = Path(__file__).resolve().parent / "Q1.xlsx"


def extract_q1(src: Path = SRC, dst: Path = DST) -> pd.DataFrame:
    df = pd.read_excel(src, header=1, usecols="F:I")
    q1 = df[df["分区"] == 1].copy().reset_index(drop=True)
    q1.to_excel(dst, sheet_name="Q1", index=False)
    return q1


if __name__ == "__main__":
    q1 = extract_q1()
    print(f"已提取 1 区期刊 {len(q1)} 种，保存至: {DST}")

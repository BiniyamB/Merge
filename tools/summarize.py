"""Summarize the structure of an Excel report: sheets, ranges, headers,
per-column fill stats (including blank columns), and sample rows."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd
from openpyxl.utils import get_column_letter

from merger import _detect_engine


def main(path: str, max_samples: int = 5):
    data = open(path, "rb").read()
    engine = _detect_engine(data)
    sheets = pd.read_excel(
        __import__("io").BytesIO(data),
        sheet_name=None,
        header=None,
        dtype=object,
        keep_default_na=False,
        engine=engine,
    )
    print(f"FILE: {path}  (engine={engine})")
    for sname, grid in sheets.items():
        print(f"\nSHEET: {sname!r}  shape={grid.shape}")
        if grid.empty:
            print("  (empty)")
            continue
        print("  First 3 rows:")
        for i in range(min(3, len(grid))):
            vals = [str(v)[:28] for v in grid.iloc[i].tolist()]
            print(f"    row {i + 1}: {vals}")
        # header = first row containing ACQUIRER-ish? just show row 1 header guess
        print("  Per-column fill stats (non-empty counts over all rows):")
        for c in range(grid.shape[1]):
            col = grid.iloc[:, c]
            filled = sum(1 for v in col if not (v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and v.strip() == "")))
            samples = [str(v)[:30] for v in col if not (v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and v.strip() == ""))][:max_samples]
            print(f"    {get_column_letter(c + 1)}: filled={filled} samples={samples}")


if __name__ == "__main__":
    main(sys.argv[1])

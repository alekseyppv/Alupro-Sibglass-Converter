from __future__ import annotations

import pandas as pd
from openpyxl import load_workbook


class ExcelRepository:
    def read_lines(self, path: str) -> list[str]:
        df = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
        lines: list[str] = []
        for row in df.fillna("").values.tolist():
            joined = " ".join(str(x) for x in row if str(x).strip())
            if joined.strip():
                lines.append(joined)
        return lines

    def open_workbook(self, path: str):
        return load_workbook(path)

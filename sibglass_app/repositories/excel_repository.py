from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook


class ExcelRepository:
    def read_lines(self, path: str) -> list[str]:
        suffix = Path(path).suffix.lower()
        if suffix == ".xlsx":
            return self._read_xlsx_lines(path)
        if suffix == ".xls":
            return self._read_xls_lines(path)
        raise ValueError("Поддерживаются только файлы .xlsx и .xls")

    def open_workbook(self, path: str):
        suffix = Path(path).suffix.lower()
        if suffix != ".xlsx":
            raise ValueError(
                "Файл заявки должен быть в формате .xlsx. Для записи в .xls сохраните шаблон как .xlsx и выберите его."
            )
        return load_workbook(path)

    @staticmethod
    def _read_xlsx_lines(path: str) -> list[str]:
        workbook = load_workbook(path, data_only=True)
        sheet = workbook.active
        lines: list[str] = []
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            values = [str(cell.value).strip() for cell in row if cell.value is not None and str(cell.value).strip()]
            if values:
                lines.append(" ".join(values))
        return lines

    @staticmethod
    def _read_xls_lines(path: str) -> list[str]:
        try:
            import xlrd  # type: ignore
        except ImportError as exc:
            raise ValueError(
                "Для чтения файлов .xls установите зависимость xlrd>=2.0.1: pip install xlrd"
            ) from exc

        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        lines: list[str] = []
        for row_idx in range(sheet.nrows):
            row_values: list[str] = []
            for col_idx in range(sheet.ncols):
                value = sheet.cell_value(row_idx, col_idx)
                text = str(value).strip()
                if text:
                    row_values.append(text)
            if row_values:
                lines.append(" ".join(row_values))
        return lines

from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from sibglass_app.models.order_item import OrderItem
from sibglass_app.utils.excel_utils import find_cell_by_value


class SibglassWriterService:
    def write(self, workbook, customer: str, address: str, items: list[OrderItem]) -> None:
        sheet = workbook.active
        self._fill_requisites(sheet, customer, address)
        start_row = self._find_table_start(sheet)
        self._write_items(sheet, start_row, items)

    @classmethod
    def _fill_requisites(cls, sheet: Worksheet, customer: str, address: str) -> None:
        customer_cell = find_cell_by_value(sheet, "Заказчик")
        address_cell = find_cell_by_value(sheet, "Адрес доставки")
        if customer_cell:
            cls._write_text_right_of_label(sheet, customer_cell.row, customer_cell.column, customer)
        if address_cell:
            cls._write_text_right_of_label(sheet, address_cell.row, address_cell.column, address)

    @classmethod
    def _write_text_right_of_label(cls, sheet: Worksheet, row: int, label_col: int, text: str) -> None:
        col = label_col + 1
        max_search = max(sheet.max_column + 20, col + 20)

        while col <= max_search:
            merged_range = cls._find_merged_range(sheet, row, col)
            if merged_range is None:
                sheet.cell(row=row, column=col, value=text)
                return

            anchor_row, anchor_col = merged_range.min_row, merged_range.min_col
            if anchor_row == row and anchor_col > label_col:
                sheet.cell(row=anchor_row, column=anchor_col, value=text)
                return

            col = merged_range.max_col + 1

        # Fallback: если всё справа занято merge-ячейками, пишем в следующую доступную обычную
        sheet.cell(row=row, column=label_col + 1, value=text)

    @staticmethod
    def _find_merged_range(sheet: Worksheet, row: int, col: int):
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                return merged_range
        return None

    @staticmethod
    def _find_table_start(sheet: Worksheet) -> int:
        for row_idx in range(1, sheet.max_row + 1):
            if str(sheet.cell(row_idx, 1).value).strip() == "№":
                return row_idx + 1
        return 15

    @staticmethod
    def _write_items(sheet: Worksheet, start_row: int, items: list[OrderItem]) -> None:
        row = start_row
        for item in items:
            sheet.cell(row=row, column=1, value=item.index)
            sheet.cell(row=row, column=2, value="")
            sheet.cell(row=row, column=3, value=item.formula)
            sheet.cell(row=row, column=4, value=item.width)
            sheet.cell(row=row, column=5, value=item.height)
            sheet.cell(row=row, column=6, value=item.count)
            sheet.cell(row=row, column=7, value=round(item.area, 4))
            sheet.cell(row=row, column=8, value=round(item.total_area, 4))
            row += 1

        cleanup_limit = max(row + 1, sheet.max_row)
        for clear_row in range(row, cleanup_limit + 1):
            if sheet.cell(clear_row, 1).value is None:
                continue
            for col in range(1, 9):
                sheet.cell(clear_row, col, value=None)

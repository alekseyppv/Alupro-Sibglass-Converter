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

    @staticmethod
    def _fill_requisites(sheet: Worksheet, customer: str, address: str) -> None:
        customer_cell = find_cell_by_value(sheet, "Заказчик")
        address_cell = find_cell_by_value(sheet, "Адрес доставки")
        if customer_cell:
            sheet.cell(row=customer_cell.row, column=customer_cell.column + 1, value=customer)
        if address_cell:
            sheet.cell(row=address_cell.row, column=address_cell.column + 1, value=address)

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

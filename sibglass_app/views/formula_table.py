from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem

from sibglass_app.models.formula_item import FormulaRowState


class FormulaTableWidget(QTableWidget):
    HEADERS = ["Исходная формула", "Итоговая формула"]

    def __init__(self, parent=None) -> None:
        super().__init__(0, len(self.HEADERS), parent)
        self.setHorizontalHeaderLabels(self.HEADERS)
        self.itemChanged.connect(self._on_item_changed)
        self._updating = False

    def set_rows(self, rows: list[FormulaRowState]) -> None:
        self._updating = True
        self.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            source_item = QTableWidgetItem(row.source_formula)
            source_item.setFlags(source_item.flags() & ~Qt.ItemIsEditable)
            target_item = QTableWidgetItem(row.resolved_formula)
            self.setItem(row_idx, 0, source_item)
            self.setItem(row_idx, 1, target_item)
            self._apply_highlight(row_idx, row.modified)
        self._updating = False

    def collect_rows(self) -> list[FormulaRowState]:
        rows: list[FormulaRowState] = []
        for row in range(self.rowCount()):
            source = self.item(row, 0).text().strip()
            resolved = self.item(row, 1).text().strip()
            modified = self.item(row, 1).background().color() == QColor("#fff59d")
            rows.append(FormulaRowState(source_formula=source, resolved_formula=resolved, modified=modified))
        return rows

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._updating or item.column() != 1:
            return
        self._apply_highlight(item.row(), True)

    def _apply_highlight(self, row: int, modified: bool) -> None:
        color = QColor("#fff59d") if modified else QColor("white")
        if self.item(row, 1):
            self.item(row, 1).setBackground(color)

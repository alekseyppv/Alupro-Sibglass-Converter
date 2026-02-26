from __future__ import annotations

from PySide6.QtWidgets import QDialog, QDialogButtonBox, QLineEdit, QVBoxLayout, QLabel


class ManualInputDialog(QDialog):
    def __init__(self, title: str, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self._line_edit = QLineEdit(self)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Введите значение:", self))
        layout.addWidget(self._line_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @property
    def value(self) -> str:
        return self._line_edit.text().strip()

from __future__ import annotations

from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)

from sibglass_app.views.formula_table import FormulaTableWidget


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Создание заявки СибГласс из AluPro")
        self.resize(1100, 760)

        root = QWidget(self)
        self.setCentralWidget(root)
        main_layout = QVBoxLayout(root)

        self.alupro_line = QLineEdit(self)
        self.sibglass_line = QLineEdit(self)
        self.select_alupro_btn = QPushButton("Выбрать файл", self)
        self.select_sibglass_btn = QPushButton("Выбрать файл", self)

        file_grid = QGridLayout()
        file_grid.addWidget(QLabel("Файл AluPro"), 0, 0)
        file_grid.addWidget(self.select_alupro_btn, 1, 0)
        file_grid.addWidget(self.alupro_line, 1, 1)
        file_grid.addWidget(QLabel("Файл заявки СибГласс"), 0, 2)
        file_grid.addWidget(self.select_sibglass_btn, 1, 2)
        file_grid.addWidget(self.sibglass_line, 1, 3)
        main_layout.addLayout(file_grid)

        form_layout = QFormLayout()
        self.customer_line = QLineEdit(self)
        self.customer_line.setText("ИП Колодинов С.С.")
        self.address_line = QLineEdit(self)
        form_layout.addRow("Заказчик", self.customer_line)
        form_layout.addRow("Адрес", self.address_line)
        main_layout.addLayout(form_layout)

        self.zak_outer = QCheckBox("Зак", self)
        self.zak_middle = QCheckBox("Зак", self)
        self.zak_inner = QCheckBox("Зак", self)
        self.argon = QCheckBox("Арг", self)

        self.outer_combo = QComboBox(self)
        self.middle_combo = QComboBox(self)
        self.inner_combo = QComboBox(self)
        self.spacer_combo = QComboBox(self)

        self.manual_outer_btn = QPushButton("Ручной ввод", self)
        self.manual_middle_btn = QPushButton("Ручной ввод", self)
        self.manual_inner_btn = QPushButton("Ручной ввод", self)
        self.manual_spacer_btn = QPushButton("Ручной ввод", self)

        options_box = QGroupBox("Комплектация", self)
        options_grid = QGridLayout(options_box)
        self._add_option_row(options_grid, 0, "Стекло наружное", self.zak_outer, self.outer_combo, self.manual_outer_btn)
        self._add_option_row(options_grid, 1, "Стекло среднее", self.zak_middle, self.middle_combo, self.manual_middle_btn)
        self._add_option_row(options_grid, 2, "Стекло внутреннее", self.zak_inner, self.inner_combo, self.manual_inner_btn)
        self._add_option_row(options_grid, 3, "Рамка", self.argon, self.spacer_combo, self.manual_spacer_btn)
        main_layout.addWidget(options_box)

        self.formula_table = FormulaTableWidget(self)
        main_layout.addWidget(QLabel("Найденные формулы", self))
        main_layout.addWidget(self.formula_table)

        bottom_row = QHBoxLayout()
        self.open_glass_btn = QPushButton("Открыть список стекол", self)
        self.save_btn = QPushButton("Сохранить заявку", self)
        bottom_row.addWidget(self.open_glass_btn)
        bottom_row.addStretch(1)
        bottom_row.addWidget(self.save_btn)
        main_layout.addLayout(bottom_row)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        main_layout.addWidget(self.progress_bar)

    def _add_option_row(self, layout: QGridLayout, row: int, title: str, checkbox: QCheckBox, combo: QComboBox, button: QPushButton) -> None:
        layout.addWidget(QLabel(title, self), row, 0)
        layout.addWidget(checkbox, row, 1)
        layout.addWidget(combo, row, 2)
        layout.addWidget(button, row, 3)

    def pick_file(self, caption: str, initial_path: str) -> str:
        path, _ = QFileDialog.getOpenFileName(self, caption, initial_path, "Excel (*.xlsx *.xls)")
        return path

    def show_error(self, message: str) -> None:
        QMessageBox.critical(self, "Ошибка", message)

    def show_warning(self, message: str) -> None:
        QMessageBox.warning(self, "Внимание", message)

    def ask_restore(self) -> bool:
        result = QMessageBox.question(
            self,
            "Восстановление",
            "Найдено автосохранение. Восстановить состояние?",
            QMessageBox.Yes | QMessageBox.No,
        )
        return result == QMessageBox.Yes

    def ask_manual_input(self, title: str, handler: Callable[[], str]) -> str:
        return handler()

    def set_busy(self, busy: bool) -> None:
        for widget in [
            self.select_alupro_btn,
            self.select_sibglass_btn,
            self.save_btn,
            self.manual_outer_btn,
            self.manual_middle_btn,
            self.manual_inner_btn,
            self.manual_spacer_btn,
            self.open_glass_btn,
        ]:
            widget.setDisabled(busy)
        self.setCursor(Qt.WaitCursor if busy else Qt.ArrowCursor)

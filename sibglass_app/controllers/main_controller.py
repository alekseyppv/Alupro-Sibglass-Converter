from __future__ import annotations

import logging
import subprocess
import sys
import traceback

from PySide6.QtCore import QTimer

from sibglass_app.config.paths import GLASS_FILE
from sibglass_app.config.settings import SettingsManager
from sibglass_app.models.formula_item import FormulaRowState
from sibglass_app.models.order_item import OrderItem
from sibglass_app.services.alupro_parser import AluProParserService
from sibglass_app.services.autosave_service import AutosaveService
from sibglass_app.services.formula_builder import FormulaBuilderService
from sibglass_app.services.glass_catalog_service import GlassCatalogService
from sibglass_app.services.sibglass_writer import SibglassWriterService
from sibglass_app.services.validation_service import ValidationService
from sibglass_app.utils.text_utils import is_numeric_formula
from sibglass_app.views.dialogs import ManualInputDialog
from sibglass_app.views.main_window import MainWindow

logger = logging.getLogger(__name__)


class MainController:
    def __init__(
        self,
        window: MainWindow,
        settings_manager: SettingsManager,
        validation_service: ValidationService,
        parser_service: AluProParserService,
        writer_service: SibglassWriterService,
        formula_builder: FormulaBuilderService,
        glass_catalog_service: GlassCatalogService,
        autosave_service: AutosaveService,
        excel_repository,
    ) -> None:
        self.window = window
        self.settings_manager = settings_manager
        self.validation_service = validation_service
        self.parser_service = parser_service
        self.writer_service = writer_service
        self.formula_builder = formula_builder
        self.glass_catalog_service = glass_catalog_service
        self.autosave_service = autosave_service
        self.excel_repository = excel_repository

        self.settings = self.settings_manager.load()
        self.catalog = None
        self._glass_mtime: float | None = None

        self._bind()
        self._load_catalog()
        self._restore_autosave_if_needed()
        self._apply_settings()
        self._start_glass_file_watcher()

    def _bind(self) -> None:
        self.window.select_alupro_btn.clicked.connect(self.on_pick_alupro)
        self.window.select_sibglass_btn.clicked.connect(self.on_pick_sibglass)
        self.window.save_btn.clicked.connect(self.on_generate)
        self.window.open_glass_btn.clicked.connect(self.on_open_glass_file)

        self.window.manual_outer_btn.clicked.connect(lambda: self.on_manual_add("outer_glass", "Стекло наружное"))
        self.window.manual_middle_btn.clicked.connect(lambda: self.on_manual_add("middle_glass", "Стекло среднее"))
        self.window.manual_inner_btn.clicked.connect(lambda: self.on_manual_add("inner_glass", "Стекло внутреннее"))
        self.window.manual_spacer_btn.clicked.connect(lambda: self.on_manual_add("spacers", "Рамки"))

        for line in [self.window.alupro_line, self.window.sibglass_line, self.window.customer_line, self.window.address_line]:
            line.textChanged.connect(self.on_any_change)

        for box in [self.window.zak_outer, self.window.zak_middle, self.window.zak_inner, self.window.argon]:
            box.stateChanged.connect(self.on_any_change)

        for combo in [self.window.outer_combo, self.window.middle_combo, self.window.inner_combo, self.window.spacer_combo]:
            combo.currentTextChanged.connect(self.on_any_change)

        self.window.formula_table.itemChanged.connect(lambda *_: self.on_any_change())

    def _apply_settings(self) -> None:
        self.window.alupro_line.setText(self.settings.last_alupro_path)
        self.window.sibglass_line.setText(self.settings.last_sibglass_path)

    def _load_catalog(self) -> None:
        catalog, exists = self.glass_catalog_service.load_or_empty()
        self.catalog = catalog
        if not exists:
            self.window.show_warning("Файл glass.txt не найден. Будет создан при первом сохранении.")
            self.glass_catalog_service.save(self.catalog)
        self._refresh_catalog_ui()
        self._remember_glass_mtime()

    def _remember_glass_mtime(self) -> None:
        self._glass_mtime = GLASS_FILE.stat().st_mtime if GLASS_FILE.exists() else None

    def _refresh_catalog_ui(self) -> None:
        current_outer = self.window.outer_combo.currentText()
        current_middle = self.window.middle_combo.currentText()
        current_inner = self.window.inner_combo.currentText()
        current_spacer = self.window.spacer_combo.currentText()

        self.window.outer_combo.clear()
        self.window.outer_combo.addItems(self.catalog.outer_glass)

        self.window.middle_combo.clear()
        self.window.middle_combo.addItems(self.catalog.middle_glass)

        self.window.inner_combo.clear()
        self.window.inner_combo.addItems(self.catalog.inner_glass)

        self.window.spacer_combo.clear()
        self.window.spacer_combo.addItems(self.catalog.spacers)

        self._select_if_exists(self.window.outer_combo, current_outer)
        self._select_if_exists(self.window.middle_combo, current_middle)
        self._select_if_exists(self.window.inner_combo, current_inner)
        self._select_if_exists(self.window.spacer_combo, current_spacer)

    @staticmethod
    def _select_if_exists(combo, value: str) -> None:
        if not value:
            return
        idx = combo.findText(value)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def _restore_autosave_if_needed(self) -> None:
        payload = self.autosave_service.load_state()
        if not payload:
            return
        if not self.window.ask_restore():
            return

        self.window.alupro_line.setText(payload.get("alupro", ""))
        self.window.sibglass_line.setText(payload.get("sibglass", ""))
        self.window.customer_line.setText(payload.get("customer", self.window.customer_line.text()))
        self.window.address_line.setText(payload.get("address", ""))
        self.window.zak_outer.setChecked(payload.get("zak_outer", False))
        self.window.zak_middle.setChecked(payload.get("zak_middle", False))
        self.window.zak_inner.setChecked(payload.get("zak_inner", False))
        self.window.argon.setChecked(payload.get("argon", False))

        self._select_if_exists(self.window.outer_combo, payload.get("outer", ""))
        self._select_if_exists(self.window.middle_combo, payload.get("middle", ""))
        self._select_if_exists(self.window.inner_combo, payload.get("inner", ""))
        self._select_if_exists(self.window.spacer_combo, payload.get("spacer", ""))

    def _start_glass_file_watcher(self) -> None:
        self._watch_timer = QTimer(self.window)
        self._watch_timer.setInterval(1200)
        self._watch_timer.timeout.connect(self._reload_catalog_if_changed)
        self._watch_timer.start()

    def _reload_catalog_if_changed(self) -> None:
        if not GLASS_FILE.exists():
            return
        mtime = GLASS_FILE.stat().st_mtime
        if self._glass_mtime is not None and mtime <= self._glass_mtime:
            return
        try:
            self.catalog = self.glass_catalog_service.load_or_empty()[0]
            self._refresh_catalog_ui()
            self._glass_mtime = mtime
        except Exception:
            logger.exception("Не удалось обновить справочник glass.txt")

    def on_pick_alupro(self) -> None:
        path = self.window.pick_file("Выберите файл AluPro", self.settings.last_alupro_path)
        if not path:
            return
        try:
            self._validate_file(path, marker="Заполнения")
        except Exception:
            return

        self.window.alupro_line.setText(path)
        self.settings.last_alupro_path = path
        self.settings_manager.save(self.settings)
        self._load_formulas()

    def on_pick_sibglass(self) -> None:
        path = self.window.pick_file("Выберите файл заявки СибГласс", self.settings.last_sibglass_path)
        if not path:
            return
        try:
            self._validate_file(path, marker="ЗАЯВКА НА РАСЧЕТ СТЕКЛОПАКЕТОВ")
        except Exception:
            return

        self.window.sibglass_line.setText(path)
        self.settings.last_sibglass_path = path
        self.settings_manager.save(self.settings)

    def _validate_file(self, path: str, marker: str) -> None:
        try:
            self.validation_service.validate_extension(path)
            self.validation_service.validate_contains(path, marker)
        except Exception as exc:
            logger.exception("Ошибка проверки файла")
            self.window.show_error(str(exc))
            raise

    def _load_formulas(self) -> None:
        try:
            items = self.parser_service.parse(self.window.alupro_line.text())
            unique = sorted({item.formula.strip() for item in items if item.formula.strip()})
            rows = [FormulaRowState(source_formula=f) for f in unique]
            for row in rows:
                if is_numeric_formula(row.source_formula):
                    row.resolved_formula = self._autobuild(row.source_formula)
            self.window.formula_table.set_rows(rows)
            if not rows:
                self.window.show_warning("В блоке 'Заполнения' не найдены строки формул.")
        except Exception:
            logger.exception("Ошибка парсинга AluPro")
            self.window.show_error("Не удалось разобрать файл AluPro. Подробности в errors.log")

    def _autobuild(self, formula: str) -> str:
        return self.formula_builder.build(
            source_formula=formula,
            outer_glass=self.window.outer_combo.currentText(),
            middle_glass=self.window.middle_combo.currentText(),
            inner_glass=self.window.inner_combo.currentText(),
            spacer=self.window.spacer_combo.currentText(),
            zak_outer=self.window.zak_outer.isChecked(),
            zak_middle=self.window.zak_middle.isChecked(),
            zak_inner=self.window.zak_inner.isChecked(),
            argon=self.window.argon.isChecked(),
        )

    def on_generate(self) -> None:
        try:
            self.window.set_busy(True)
            self.window.progress_bar.setValue(10)

            alupro_items = self.parser_service.parse(self.window.alupro_line.text())
            formula_map = {
                row.source_formula: row.resolved_formula
                for row in self.window.formula_table.collect_rows()
                if row.resolved_formula.strip()
            }
            orders: list[OrderItem] = []
            idx = 1
            for item in alupro_items:
                resolved = formula_map.get(item.formula, "")
                if not resolved:
                    continue
                orders.append(OrderItem(index=idx, formula=resolved, width=item.width, height=item.height, count=item.count))
                idx += 1

            wb = self.excel_repository.open_workbook(self.window.sibglass_line.text())
            self.window.progress_bar.setValue(50)
            self.writer_service.write(
                wb,
                customer=self.window.customer_line.text().strip(),
                address=self.window.address_line.text().strip(),
                items=orders,
            )
            wb.save(self.window.sibglass_line.text())
            self.window.progress_bar.setValue(100)
            self.autosave_service.clear()
        except Exception as exc:
            logger.error("Ошибка генерации: %s", exc)
            logger.error(traceback.format_exc())
            self.window.show_error("Ошибка при сохранении заявки. Подробности в errors.log")
        finally:
            self.window.set_busy(False)

    def on_manual_add(self, section_attr: str, title: str) -> None:
        dialog = ManualInputDialog(title, parent=self.window)
        if dialog.exec() != dialog.Accepted:
            return

        value = dialog.value
        if not value:
            return

        try:
            self.catalog = self.glass_catalog_service.add_value(self.catalog, section_attr, value)
            self.glass_catalog_service.save(self.catalog)
            self._refresh_catalog_ui()
            combo_by_section = {
                "outer_glass": self.window.outer_combo,
                "middle_glass": self.window.middle_combo,
                "inner_glass": self.window.inner_combo,
                "spacers": self.window.spacer_combo,
            }
            self._select_if_exists(combo_by_section[section_attr], value)
            self._remember_glass_mtime()
        except Exception:
            logger.exception("Не удалось сохранить ручной ввод в glass.txt")
            self.window.show_error("Не удалось сохранить значение в glass.txt")

    def on_open_glass_file(self) -> None:
        if not GLASS_FILE.exists():
            self.glass_catalog_service.save(self.catalog)
            self._remember_glass_mtime()
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["notepad.exe", str(GLASS_FILE)])
            else:
                subprocess.Popen(["xdg-open", str(GLASS_FILE)])
        except Exception:
            logger.exception("Не удалось открыть glass.txt")
            self.window.show_error("Не удалось открыть glass.txt")

    def on_any_change(self) -> None:
        payload = {
            "alupro": self.window.alupro_line.text().strip(),
            "sibglass": self.window.sibglass_line.text().strip(),
            "customer": self.window.customer_line.text().strip(),
            "address": self.window.address_line.text().strip(),
            "zak_outer": self.window.zak_outer.isChecked(),
            "zak_middle": self.window.zak_middle.isChecked(),
            "zak_inner": self.window.zak_inner.isChecked(),
            "argon": self.window.argon.isChecked(),
            "outer": self.window.outer_combo.currentText(),
            "middle": self.window.middle_combo.currentText(),
            "inner": self.window.inner_combo.currentText(),
            "spacer": self.window.spacer_combo.currentText(),
        }
        self.autosave_service.save_state(payload)

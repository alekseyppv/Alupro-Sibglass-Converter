from __future__ import annotations

import logging
import subprocess
import sys
import traceback
from dataclasses import asdict

from sibglass_app.config.settings import AppSettings, SettingsManager
from sibglass_app.models.formula_item import FormulaRowState
from sibglass_app.models.order_item import OrderItem
from sibglass_app.services.alupro_parser import AluProParserService
from sibglass_app.services.autosave_service import AutosaveService
from sibglass_app.services.formula_builder import FormulaBuilderService
from sibglass_app.services.glass_catalog_service import GlassCatalogService
from sibglass_app.services.sibglass_writer import SibglassWriterService
from sibglass_app.services.validation_service import ValidationService
from sibglass_app.views.dialogs import ManualInputDialog
from sibglass_app.views.main_window import MainWindow
from sibglass_app.utils.text_utils import is_numeric_formula

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

        self._bind()
        self._load_catalog()
        self._restore_autosave_if_needed()
        self._apply_settings()

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
            self.window.show_warning("Файл glass.txt не найден. Доступен только ручной ввод.")
        self._refresh_catalog_ui()

    def _refresh_catalog_ui(self) -> None:
        self.window.outer_combo.clear()
        self.window.outer_combo.addItems(self.catalog.outer_glass)

        self.window.middle_combo.clear()
        self.window.middle_combo.addItems(self.catalog.middle_glass)

        self.window.inner_combo.clear()
        self.window.inner_combo.addItems(self.catalog.inner_glass)

        self.window.spacer_combo.clear()
        self.window.spacer_combo.addItems(self.catalog.spacers)

    def _restore_autosave_if_needed(self) -> None:
        payload = self.autosave_service.load_state()
        if not payload:
            return
        if not self.window.ask_restore():
            return
        self.window.alupro_line.setText(payload.get("alupro", ""))
        self.window.sibglass_line.setText(payload.get("sibglass", ""))
        self.window.customer_line.setText(payload.get("customer", ""))
        self.window.address_line.setText(payload.get("address", ""))
        self.window.zak_outer.setChecked(payload.get("zak_outer", False))
        self.window.zak_middle.setChecked(payload.get("zak_middle", False))
        self.window.zak_inner.setChecked(payload.get("zak_inner", False))
        self.window.argon.setChecked(payload.get("argon", False))

    def on_pick_alupro(self) -> None:
        path = self.window.pick_file("Выберите файл AluPro", self.settings.last_alupro_path)
        if not path:
            return
        self._validate_file(path, marker="Заполнения")
        self.window.alupro_line.setText(path)
        self.settings.last_alupro_path = path
        self.settings_manager.save(self.settings)
        self._load_formulas()

    def on_pick_sibglass(self) -> None:
        path = self.window.pick_file("Выберите файл заявки СибГласс", self.settings.last_sibglass_path)
        if not path:
            return
        self._validate_file(path, marker="ЗАЯВКА НА РАСЧЕТ СТЕКЛОПАКЕТОВ")
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
            unique = sorted({item.formula for item in items})
            rows = [FormulaRowState(source_formula=f) for f in unique]
            for row in rows:
                if is_numeric_formula(row.source_formula):
                    row.resolved_formula = self._autobuild(row.source_formula)
            self.window.formula_table.set_rows(rows)
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
        self.catalog = self.glass_catalog_service.add_value(self.catalog, section_attr, value)
        self.glass_catalog_service.save(self.catalog)
        self._refresh_catalog_ui()

    def on_open_glass_file(self) -> None:
        from sibglass_app.config.paths import GLASS_FILE

        if not GLASS_FILE.exists():
            self.glass_catalog_service.save(self.catalog)
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

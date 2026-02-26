from __future__ import annotations

from PySide6.QtWidgets import QApplication

from sibglass_app.config.settings import SettingsManager
from sibglass_app.controllers import MainController
from sibglass_app.repositories.excel_repository import ExcelRepository
from sibglass_app.repositories.glass_file_repository import GlassFileRepository
from sibglass_app.services.alupro_parser import AluProParserService
from sibglass_app.services.autosave_service import AutosaveService
from sibglass_app.services.formula_builder import FormulaBuilderService
from sibglass_app.services.glass_catalog_service import GlassCatalogService
from sibglass_app.services.sibglass_writer import SibglassWriterService
from sibglass_app.services.validation_service import ValidationService
from sibglass_app.utils.logger import configure_logging
from sibglass_app.views.main_window import MainWindow


class SibglassApplication:
    def __init__(self) -> None:
        configure_logging()

        self.qt_app = QApplication([])

        excel_repository = ExcelRepository()
        glass_repository = GlassFileRepository()

        window = MainWindow()
        self.controller = MainController(
            window=window,
            settings_manager=SettingsManager(),
            validation_service=ValidationService(excel_repository),
            parser_service=AluProParserService(excel_repository),
            writer_service=SibglassWriterService(),
            formula_builder=FormulaBuilderService(),
            glass_catalog_service=GlassCatalogService(glass_repository),
            autosave_service=AutosaveService(),
            excel_repository=excel_repository,
        )
        self.window = window

    def run(self) -> int:
        self.window.show()
        return self.qt_app.exec()

from __future__ import annotations

from pathlib import Path

from sibglass_app.repositories.excel_repository import ExcelRepository


class ValidationService:
    def __init__(self, excel_repo: ExcelRepository) -> None:
        self._excel_repo = excel_repo

    def validate_extension(self, path: str) -> None:
        if Path(path).suffix.lower() not in {".xlsx", ".xls"}:
            raise ValueError("Поддерживаются только файлы .xlsx и .xls")

    def validate_contains(self, path: str, marker: str) -> None:
        lines = self._excel_repo.read_lines(path)
        marker_lower = marker.lower()
        if not any(marker_lower in line.lower() for line in lines):
            raise ValueError(f"Файл {Path(path).name} не содержит обязательный маркер: {marker}")

from __future__ import annotations

import re

from sibglass_app.models.formula_item import FormulaItem
from sibglass_app.repositories.excel_repository import ExcelRepository


class AluProParserService:
    def __init__(self, excel_repository: ExcelRepository) -> None:
        self._excel_repository = excel_repository

    def parse(self, path: str) -> list[FormulaItem]:
        rows = self._excel_repository.read_rows(path)

        # Приоритет: табличный парсинг по заголовкам колонок
        table_items = self._parse_by_table_headers(rows)
        if table_items:
            return table_items

        # Фолбэк: старый режим по блоку "Заполнения" ... "Сумма:"
        block = self._extract_fillings_block(rows)
        parsed: list[FormulaItem] = []
        for row in block:
            item = self._parse_row_fallback(row)
            if item is not None:
                parsed.append(item)
        return parsed

    def _parse_by_table_headers(self, rows: list[list[str]]) -> list[FormulaItem]:
        header_idx = -1
        formula_col = width_col = height_col = count_col = -1

        for idx, row in enumerate(rows):
            normalized = [c.strip().lower() for c in row]
            if not normalized:
                continue

            c_formula = self._find_col(normalized, ["наименование"])
            c_width = self._find_col(normalized, ["ширина"])
            c_height = self._find_col(normalized, ["высота"])
            c_count = self._find_col(normalized, ["кол-во", "кол во", "количество"])
            if min(c_formula, c_width, c_height, c_count) >= 0:
                header_idx = idx
                formula_col, width_col, height_col, count_col = c_formula, c_width, c_height, c_count
                break

        if header_idx < 0:
            return []

        items: list[FormulaItem] = []
        for row in rows[header_idx + 1 :]:
            joined = " ".join(cell.strip().lower() for cell in row if cell)
            if "сумма:" in joined:
                break

            formula = self._safe_get(row, formula_col).split(",", 1)[0].strip()
            if not self._is_formula_candidate(formula):
                continue

            width = self._parse_int(self._safe_get(row, width_col))
            height = self._parse_int(self._safe_get(row, height_col))
            count = self._parse_int(self._safe_get(row, count_col), default=1)
            items.append(FormulaItem(formula=formula, width=width, height=height, count=max(count, 1)))

        return items

    @staticmethod
    def _safe_get(row: list[str], index: int) -> str:
        if index < 0 or index >= len(row):
            return ""
        return row[index].strip()

    @staticmethod
    def _find_col(cells: list[str], tokens: list[str]) -> int:
        for idx, cell in enumerate(cells):
            if any(token in cell for token in tokens):
                return idx
        return -1

    @staticmethod
    def _parse_int(text: str, default: int = 0) -> int:
        match = re.search(r"\d+", text.replace(",", "."))
        if not match:
            return default
        return int(match.group(0))

    @staticmethod
    def _extract_fillings_block(rows: list[list[str]]) -> list[list[str]]:
        started = False
        block: list[list[str]] = []
        for row in rows:
            joined = " ".join(cell for cell in row if cell).lower()
            if "заполнения" in joined:
                started = True
                continue
            if started and "сумма:" in joined:
                break
            if started:
                block.append(row)
        return block

    def _parse_row_fallback(self, row: list[str]) -> FormulaItem | None:
        cells = [cell.strip() for cell in row if cell and cell.strip()]
        if not cells:
            return None

        formula = self._extract_formula(cells)
        if not formula:
            return None

        size_source = " ".join(cells[1:])
        count_source = " ".join(cells[2:])
        width, height = self._extract_size(size_source)
        count = self._extract_count(count_source)
        return FormulaItem(formula=formula, width=width, height=height, count=count)

    @staticmethod
    def _extract_formula(cells: list[str]) -> str:
        first = cells[0]
        candidate = first.split(",", 1)[0].strip()
        if AluProParserService._is_formula_candidate(candidate):
            return candidate

        for cell in cells:
            value = cell.split(",", 1)[0].strip()
            if AluProParserService._is_formula_candidate(value):
                return value
        return ""

    @staticmethod
    def _is_formula_candidate(value: str) -> bool:
        lowered = value.lower()
        if not value:
            return False
        if "заполнения" in lowered or "сумма:" in lowered:
            return False
        if "glass" == lowered or "площад" in lowered or "ширина" in lowered or "высота" in lowered:
            return False
        return bool(re.search(r"\d", value))

    @staticmethod
    def _extract_size(source: str) -> tuple[int, int]:
        pair_match = re.search(r"(\d{2,5})\s*[xXхХ+]\s*(\d{2,5})", source)
        if pair_match:
            return int(pair_match.group(1)), int(pair_match.group(2))

        numbers = [int(n) for n in re.findall(r"\d+", source)]
        if len(numbers) >= 2:
            return numbers[0], numbers[1]
        return 0, 0

    @staticmethod
    def _extract_count(source: str) -> int:
        match = re.search(r"\b(\d+)\b", source)
        if not match:
            return 1
        value = int(match.group(1))
        return value if value > 0 else 1

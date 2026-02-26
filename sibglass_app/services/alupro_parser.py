from __future__ import annotations

import re

from sibglass_app.models.formula_item import FormulaItem
from sibglass_app.repositories.excel_repository import ExcelRepository


class AluProParserService:
    def __init__(self, excel_repository: ExcelRepository) -> None:
        self._excel_repository = excel_repository

    def parse(self, path: str) -> list[FormulaItem]:
        rows = self._excel_repository.read_rows(path)
        block = self._extract_fillings_block(rows)

        parsed: list[FormulaItem] = []
        for row in block:
            item = self._parse_row(row)
            if item is not None:
                parsed.append(item)
        return parsed

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

    def _parse_row(self, row: list[str]) -> FormulaItem | None:
        cells = [cell.strip() for cell in row if cell and cell.strip()]
        if not cells:
            return None

<<<<<<< codex/develop-industrial-desktop-application-in-python-kjbxuq
        formula = self._extract_formula(cells)
        if not formula:
            return None

        size_source = " ".join(cells[1:])
        count_source = " ".join(cells[2:])
=======
        if "," in cells[0]:
            chunks = [part.strip() for part in cells[0].split(",") if part.strip()]
            formula = chunks[0] if chunks else ""
            size_source = " ".join(cells[1:] + chunks[1:])
            count_source = " ".join(cells[2:] + chunks[2:])
        else:
            formula = cells[0]
            size_source = " ".join(cells[1:])
            count_source = " ".join(cells[2:])

        if not formula or "заполнения" in formula.lower() or "сумма:" in formula.lower():
            return None

>>>>>>> main
        width, height = self._extract_size(size_source)
        count = self._extract_count(count_source)
        return FormulaItem(formula=formula, width=width, height=height, count=count)

    @staticmethod
<<<<<<< codex/develop-industrial-desktop-application-in-python-kjbxuq
    def _extract_formula(cells: list[str]) -> str:
        # Частый случай: всё в первой ячейке "10-16-8, 1200x800, 2"
        first = cells[0]
        candidate = first.split(",", 1)[0].strip()
        if AluProParserService._is_formula_candidate(candidate):
            return candidate

        # Если данные разнесены по колонкам, ищем первое подходящее значение
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
        # В формулах всегда есть число (в т.ч. нестандартные формулы с текстом)
        return bool(re.search(r"\d", value))

    @staticmethod
=======
>>>>>>> main
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

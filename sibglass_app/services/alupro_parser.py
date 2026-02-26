from __future__ import annotations

import re

from sibglass_app.models.formula_item import FormulaItem
from sibglass_app.repositories.excel_repository import ExcelRepository


class AluProParserService:
    def __init__(self, excel_repository: ExcelRepository) -> None:
        self._excel_repository = excel_repository

    def parse(self, path: str) -> list[FormulaItem]:
        lines = self._excel_repository.read_lines(path)
        started = False
        block: list[str] = []
        for line in lines:
            if "заполнения" in line.lower():
                started = True
                continue
            if started and "сумма:" in line.lower():
                break
            if started:
                block.append(line)

        parsed: list[FormulaItem] = []
        regex = re.compile(r"^(?P<formula>[^,]+),\s*(?P<size>\d+\+\d+(?:\+\d+)?mm)?\s*,\s*(?P<count>\d+)\s*,?", re.IGNORECASE)

        for line in block:
            match = regex.search(line)
            if not match:
                continue
            formula = match.group("formula").strip()
            size = match.group("size") or "0+0mm"
            count = int(match.group("count"))
            width, height = self._extract_size(size)
            parsed.append(FormulaItem(formula=formula, width=width, height=height, count=count))

        return parsed

    @staticmethod
    def _extract_size(size: str) -> tuple[int, int]:
        numbers = [int(x) for x in re.findall(r"\d+", size)]
        if len(numbers) >= 2:
            return numbers[0], numbers[1]
        return 0, 0

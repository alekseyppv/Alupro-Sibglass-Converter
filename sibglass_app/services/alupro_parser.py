from __future__ import annotations

import re

from sibglass_app.models.formula_item import FormulaItem
from sibglass_app.repositories.excel_repository import ExcelRepository


class AluProParserService:
    def __init__(self, excel_repository: ExcelRepository) -> None:
        self._excel_repository = excel_repository

    def parse(self, path: str) -> list[FormulaItem]:
        lines = self._excel_repository.read_lines(path)
        block = self._extract_fillings_block(lines)

        parsed: list[FormulaItem] = []
        for line in block:
            item = self._parse_line(line)
            if item is not None:
                parsed.append(item)
        return parsed

    @staticmethod
    def _extract_fillings_block(lines: list[str]) -> list[str]:
        started = False
        block: list[str] = []
        for line in lines:
            lowered = line.lower()
            if "заполнения" in lowered:
                started = True
                continue
            if started and "сумма:" in lowered:
                break
            if started:
                block.append(line)
        return block

    def _parse_line(self, line: str) -> FormulaItem | None:
        if "," not in line:
            return None

        chunks = [part.strip() for part in line.split(",")]
        formula = chunks[0].strip()
        if not formula:
            return None

        # Не отбрасываем нестандартные формулы: показываем все, что нашли в блоке
        width, height = self._extract_size_from_chunks(chunks[1:])
        count = self._extract_count_from_chunks(chunks[2:])

        return FormulaItem(
            formula=formula,
            width=width,
            height=height,
            count=count,
        )

    @staticmethod
    def _extract_size_from_chunks(chunks: list[str]) -> tuple[int, int]:
        joined = " ".join(chunks)

        # Частые форматы: 1200x800, 1200х800, 1200+800mm
        pair_match = re.search(r"(\d{2,5})\s*[xXхХ+]\s*(\d{2,5})", joined)
        if pair_match:
            return int(pair_match.group(1)), int(pair_match.group(2))

        numbers = [int(n) for n in re.findall(r"\d+", joined)]
        if len(numbers) >= 2:
            return numbers[0], numbers[1]
        return 0, 0

    @staticmethod
    def _extract_count_from_chunks(chunks: list[str]) -> int:
        if not chunks:
            return 1

        candidate = chunks[0].strip() if chunks[0].strip() else ""
        if candidate:
            num = re.search(r"\d+", candidate)
            if num:
                value = int(num.group())
                return value if value > 0 else 1

        joined = " ".join(chunks)
        numbers = [int(n) for n in re.findall(r"\d+", joined)]
        if numbers:
            value = numbers[-1]
            return value if value > 0 else 1
        return 1

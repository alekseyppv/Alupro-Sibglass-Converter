from __future__ import annotations

import re


def normalize_formula(formula: str) -> str:
    return re.sub(r"\s+", "", formula)


def is_numeric_formula(formula: str) -> bool:
    normalized = normalize_formula(formula)
    return bool(re.fullmatch(r"\d+(?:-\d+){0,4}", normalized)) and len(normalized.split("-")) in (1, 3, 5)


def extract_thicknesses(formula: str) -> list[int]:
    normalized = normalize_formula(formula)
    return [int(part) for part in normalized.split("-") if part.isdigit()]

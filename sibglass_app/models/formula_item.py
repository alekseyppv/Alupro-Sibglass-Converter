from __future__ import annotations

from dataclasses import dataclass


@dataclass
class FormulaItem:
    formula: str
    width: int
    height: int
    count: int


@dataclass
class FormulaRowState:
    source_formula: str
    resolved_formula: str = ""
    modified: bool = False

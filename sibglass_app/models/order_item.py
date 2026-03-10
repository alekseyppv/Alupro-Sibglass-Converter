from __future__ import annotations

from dataclasses import dataclass


@dataclass
class OrderItem:
    index: int
    formula: str
    width: int
    height: int
    count: int

    @property
    def area(self) -> float:
        return self.width * self.height / 1_000_000

    @property
    def total_area(self) -> float:
        return self.area * self.count

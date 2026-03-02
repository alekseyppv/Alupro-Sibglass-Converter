from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class GlassCatalog:
    outer_glass: list[str] = field(default_factory=list)
    middle_glass: list[str] = field(default_factory=list)
    inner_glass: list[str] = field(default_factory=list)
    spacers: list[str] = field(default_factory=list)


SECTION_MAP = {
    "Стекло наружное": "outer_glass",
    "Стекло среднее": "middle_glass",
    "Стекло внутреннее": "inner_glass",
    "Рамки": "spacers",
}

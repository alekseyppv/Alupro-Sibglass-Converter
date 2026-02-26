from __future__ import annotations

from sibglass_app.utils.text_utils import extract_thicknesses, is_numeric_formula


class FormulaBuilderService:
    def build(
        self,
        source_formula: str,
        outer_glass: str,
        middle_glass: str,
        inner_glass: str,
        spacer: str,
        zak_outer: bool,
        zak_middle: bool,
        zak_inner: bool,
        argon: bool,
    ) -> str:
        if not is_numeric_formula(source_formula):
            return ""

        thicknesses = extract_thicknesses(source_formula)
        glass_values = [outer_glass, middle_glass, inner_glass]
        zak_flags = [zak_outer, zak_middle, zak_inner]

        if len(thicknesses) == 1:
            return self._glass_part(thicknesses[0], glass_values[0], zak_flags[0])

        if len(thicknesses) == 3:
            left = self._glass_part(thicknesses[0], glass_values[0], zak_flags[0])
            spacer_part = self._spacer_part(thicknesses[1], spacer, argon)
            right = self._glass_part(thicknesses[2], glass_values[2], zak_flags[2])
            return f"{left}-{spacer_part}-{right}"

        if len(thicknesses) == 5:
            left = self._glass_part(thicknesses[0], glass_values[0], zak_flags[0])
            spacer_left = self._spacer_part(thicknesses[1], spacer, argon)
            middle = self._glass_part(thicknesses[2], glass_values[1], zak_flags[1])
            spacer_right = self._spacer_part(thicknesses[3], spacer, argon)
            right = self._glass_part(thicknesses[4], glass_values[2], zak_flags[2])
            return f"{left}-{spacer_left}-{middle}-{spacer_right}-{right}"

        return ""

    @staticmethod
    def _glass_part(thickness: int, name: str, zak: bool) -> str:
        suffix = "SGTemp " if zak else ""
        return f"{thickness}{suffix}{name}".strip()

    @staticmethod
    def _spacer_part(thickness: int, name: str, argon: bool) -> str:
        suffix = "Ar" if argon else ""
        return f"{thickness}{name}{suffix}".strip()

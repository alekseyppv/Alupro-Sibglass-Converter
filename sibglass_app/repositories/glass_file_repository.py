from __future__ import annotations

from sibglass_app.config.paths import GLASS_FILE
from sibglass_app.models.glass_catalog import GlassCatalog, SECTION_MAP


class GlassFileRepository:
    def load(self) -> GlassCatalog:
        if not GLASS_FILE.exists():
            raise FileNotFoundError(GLASS_FILE)

        catalog = GlassCatalog()
        current_section: str | None = None

        for raw in GLASS_FILE.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line:
                continue
            header = line.rstrip(":")
            if header in SECTION_MAP:
                current_section = SECTION_MAP[header]
                continue
            if current_section:
                items = getattr(catalog, current_section)
                if line not in items:
                    items.append(line)
        return catalog

    def save(self, catalog: GlassCatalog) -> None:
        sections = [
            ("Стекло наружное", catalog.outer_glass),
            ("Стекло среднее", catalog.middle_glass),
            ("Стекло внутреннее", catalog.inner_glass),
            ("Рамки", catalog.spacers),
        ]
        chunks: list[str] = []
        for title, values in sections:
            chunks.append(f"{title}:")
            chunks.extend(values)
            chunks.append("")
        GLASS_FILE.write_text("\n".join(chunks).strip() + "\n", encoding="utf-8")

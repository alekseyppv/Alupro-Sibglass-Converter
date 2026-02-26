from __future__ import annotations

from sibglass_app.models.glass_catalog import GlassCatalog
from sibglass_app.repositories.glass_file_repository import GlassFileRepository


class GlassCatalogService:
    def __init__(self, repository: GlassFileRepository) -> None:
        self._repository = repository

    def load_or_empty(self) -> tuple[GlassCatalog, bool]:
        try:
            return self._repository.load(), True
        except FileNotFoundError:
            return GlassCatalog(), False

    def add_value(self, catalog: GlassCatalog, section_attr: str, value: str) -> GlassCatalog:
        cleaned = value.strip()
        if not cleaned:
            return catalog
        target = getattr(catalog, section_attr)
        if cleaned not in target:
            target.append(cleaned)
        return catalog

    def save(self, catalog: GlassCatalog) -> None:
        self._repository.save(catalog)

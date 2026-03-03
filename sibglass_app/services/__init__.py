"""Service package.

Важно: не импортируем подмодули здесь, чтобы в frozen-сборках
не падать из-за преждевременной загрузки всех сервисов сразу.
"""

__all__ = [
    "alupro_parser",
    "autosave_service",
    "formula_builder",
    "glass_catalog_service",
    "sibglass_writer",
    "validation_service",
]

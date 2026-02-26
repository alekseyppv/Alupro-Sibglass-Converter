from __future__ import annotations

import json
from typing import Any

from sibglass_app.config.paths import AUTOSAVE_FILE


class AutosaveService:
    def save_state(self, payload: dict[str, Any]) -> None:
        AUTOSAVE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_state(self) -> dict[str, Any] | None:
        if not AUTOSAVE_FILE.exists():
            return None
        try:
            return json.loads(AUTOSAVE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return None

    def clear(self) -> None:
        if AUTOSAVE_FILE.exists():
            AUTOSAVE_FILE.unlink()

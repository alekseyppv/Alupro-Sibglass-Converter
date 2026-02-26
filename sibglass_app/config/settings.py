from __future__ import annotations

import json
from dataclasses import dataclass, asdict

from sibglass_app.config.paths import SETTINGS_FILE


@dataclass
class AppSettings:
    last_alupro_path: str = ""
    last_sibglass_path: str = ""


class SettingsManager:
    def load(self) -> AppSettings:
        if not SETTINGS_FILE.exists():
            return AppSettings()
        try:
            payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            return AppSettings(**payload)
        except Exception:
            return AppSettings()

    def save(self, settings: AppSettings) -> None:
        SETTINGS_FILE.write_text(
            json.dumps(asdict(settings), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

from __future__ import annotations

import sys

from sibglass_app.app import SibglassApplication


def main() -> int:
    app = SibglassApplication()
    return app.run()


if __name__ == "__main__":
    raise SystemExit(main())

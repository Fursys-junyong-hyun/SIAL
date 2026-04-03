from __future__ import annotations

import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from outsourced_inventory_confirmation.main_window import MainWindow, build_application
else:
    from .main_window import MainWindow, build_application


def main() -> None:
    app = build_application()
    window = MainWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()

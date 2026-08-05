"""Entry point for O Fechador."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> None:
    app_path = Path(__file__).resolve().parent / "src" / "app.py"
    raise SystemExit(
        subprocess.call(
            [sys.executable, "-m", "streamlit", "run", str(app_path), *sys.argv[1:]]
        )
    )


if __name__ == "__main__":
    main()

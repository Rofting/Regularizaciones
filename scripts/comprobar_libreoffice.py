"""Comprueba que LibreOffice está instalado y cumple la versión mínima.

Uso:
    python scripts/comprobar_libreoffice.py

Lo usa el CI antes de la suite y sirve para revisar una instalación nueva.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "core"))

from office_recalculation import (  # noqa: E402
    MINIMUM_LIBREOFFICE_VERSION,
    LibreOfficeRecalculator,
    RecalculationError,
    version_text,
)


def main() -> int:
    recalculator = LibreOfficeRecalculator()
    try:
        executable = recalculator._find_executable()
        version = recalculator.check_minimum_version()
    except RecalculationError as error:
        print(error, file=sys.stderr)
        return 1
    print(f"LibreOffice {version_text(version)} en {executable} "
          f"(mínima {version_text(MINIMUM_LIBREOFFICE_VERSION)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Punto de entrada del ejecutable (PyInstaller no admite importaciones relativas en el script principal)."""

import sys

from cronograma_neon.__main__ import main

if __name__ == "__main__":
    sys.exit(main())

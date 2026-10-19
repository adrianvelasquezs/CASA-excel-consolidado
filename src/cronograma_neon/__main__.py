"""Punto de entrada: sin argumentos abre la interfaz; con argumentos usa la línea de comandos."""

import os
import sys


def main(argv=None):
    # En el ejecutable de ventana (sin consola) stdout/stderr pueden no existir.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")  # noqa: SIM115
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")  # noqa: SIM115
    argv = sys.argv[1:] if argv is None else argv
    argv = [a for a in argv if not a.startswith("-psn_")]  # argumento que agrega macOS al abrir desde Finder
    if argv:
        from .cli import main as cli

        return cli(argv)
    from .interfaz.app import ejecutar

    ejecutar()
    return 0


if __name__ == "__main__":
    sys.exit(main())

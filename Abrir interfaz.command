#!/bin/bash
# Doble clic en Mac para abrir la interfaz del generador.
cd "$(dirname "$0")" || exit 1
if ! command -v uv >/dev/null 2>&1; then
  echo "Falta instalar uv: https://docs.astral.sh/uv/getting-started/installation/"
  read -r -p "Presione Enter para cerrar..."
  exit 1
fi
uv run interfaz.py

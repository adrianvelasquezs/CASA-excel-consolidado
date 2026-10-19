# -*- mode: python ; coding: utf-8 -*-
# Empaquetado con PyInstaller:
#   macOS   -> dist/Cronograma Neon.app
#   Windows -> dist/Cronograma Neon.exe (un solo archivo)
# Uso (desde la raíz del proyecto): uv run --group build pyinstaller --noconfirm --clean empaquetado/cronograma_neon.spec

import subprocess
import sys
from pathlib import Path

RAIZ = Path(SPECPATH).parent
PAQUETE = RAIZ / "src" / "cronograma_neon"
sys.path.insert(0, str(RAIZ / "src"))
from cronograma_neon import __version__  # noqa: E402

NOMBRE = "Cronograma Neon"

datas = [
    (str(PAQUETE / "google_sheets" / "apps_script"), "cronograma_neon/google_sheets/apps_script"),
    (str(PAQUETE / "guias"), "cronograma_neon/guias"),
]
# Plantillas de ejemplo incluidas como punto de partida (el usuario puede elegir otras).
datas += [(str(p), "cronograma_neon/plantillas_incluidas")
          for p in sorted((RAIZ / "raw").glob("*.xlsx")) if not p.name.startswith("~$")]

a = Analysis(
    [str(RAIZ / "empaquetado" / "lanzador.py")],
    pathex=[str(RAIZ / "src")],
    datas=datas,
    excludes=["pytest", "_pytest", "setuptools", "pip"],
    noarchive=False,
)
pyz = PYZ(a.pure)

if sys.platform == "darwin":
    # App universal (Intel + Apple Silicon) si el Python usado lo es; si no, la arquitectura del equipo.
    try:
        arquitecturas = subprocess.run(["lipo", "-archs", sys.executable], capture_output=True, text=True).stdout.split()
    except OSError:
        arquitecturas = []
    arquitectura = "universal2" if {"x86_64", "arm64"} <= set(arquitecturas) else None

    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=NOMBRE, console=False,
              target_arch=arquitectura, argv_emulation=False, upx=False)
    coll = COLLECT(exe, a.binaries, a.datas, name=NOMBRE, upx=False)
    app = BUNDLE(
        coll,
        name=f"{NOMBRE}.app",
        bundle_identifier="co.edu.uniandes.casa.cronograma-neon",
        version=__version__,
        info_plist={
            "CFBundleDisplayName": "Cronograma Neón",
            "CFBundleShortVersionString": __version__,
            "NSHighResolutionCapable": True,
            "NSRequiresAquaSystemAppearance": False,
        },
    )
else:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=NOMBRE, console=False, upx=False)

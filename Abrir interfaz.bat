@echo off
rem Doble clic en Windows para abrir la interfaz del generador.
cd /d "%~dp0"
where uv >nul 2>nul || (echo Falta instalar uv: https://docs.astral.sh/uv/getting-started/installation/ & pause & exit /b 1)
uv run interfaz.py
if errorlevel 1 pause

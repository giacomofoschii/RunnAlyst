@echo off
cd /d "%~dp0\.."
poetry run python -m runnalyst.cli
pause
@echo off
cd /d "%~dp0"
py -3 --version >nul 2>nul && (start "" pyw -3 walidator.py) || (start "" pythonw walidator.py)

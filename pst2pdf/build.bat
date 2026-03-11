@echo off
REM Build a self-contained pst2pdf.exe using PyInstaller.
REM Run this script from the pst2pdf\ directory.

echo Installing dependencies...
pip install -r requirements.txt pyinstaller

echo Building binary...
pyinstaller ^
  --onefile ^
  --name pst2pdf ^
  --add-data "*.py;." ^
  pst2pdf.py

echo.
echo Binary written to: dist\pst2pdf.exe
echo Run with: dist\pst2pdf.exe input.pst output_dir\ --manifest

echo.
echo Building GUI binary...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name pst2pdf-gui ^
  --collect-data customtkinter ^
  --add-data "*.py;." ^
  gui.py

echo.
echo GUI binary written to: dist\pst2pdf-gui.exe
echo Run with: dist\pst2pdf-gui.exe

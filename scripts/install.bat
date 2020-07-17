@echo off
REM
REM Programmbibliotheken installieren
REM
call setenv.bat

venv\Scripts\activate

pip install requests
pip install openpyxl

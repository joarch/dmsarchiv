@echo off
REM
REM Start des Programmes
REM
call setenv.bat

set PYTHONPATH=%PROGRAM_DIR%\src

venv\Scripts\activate

python dmsarchiv

pause

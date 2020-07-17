@echo off
REM
REM Update prüfen und ausführen
REM
call setenv.bat

md "%TEMP_DIR%" 2>NUL

IF NOT DEFINED GIT_HOME (
echo "Das automatische Update kann nicht durchgeführt werden. Bitte GIT_HOME konfigurieren."
goto ENDE
)

REM Prüfe auf Programmänderungen (Update)
cd %PROGRAM_DIR%
%GIT_HOME%\bin\git fetch
%GIT_HOME%\bin\git log --all --oneline -n1 > ../update_2.log
cd ..
copy update_2.log %TEMP_DIR% >NUL
del update_2.log >NUL
fc %TEMP_DIR%\update_1.log %TEMP_DIR%\update_2.log >NUL
if errorlevel 1 goto UPDATE

echo "Kein Update notwendig."
goto :ENDE

:UPDATE
REM Update Source Dateien
echo "**********************************************************"
echo "Es steht ein neues Programm Update zur Verfügung"
echo "---------------------------------------------------------"
echo "Details:"
type %TEMP_DIR%\update_2.log
echo "---------------------------------------------------------"
echo "Das Update wird jetzt eingespielt ..."
echo "---------------------------------------------------------"
cd %PROGRAM_DIR%
%GIT_HOME%\bin\git pull
%GIT_HOME%\bin\git log  --all --oneline -n1 > ../update_1.log
cd ..
copy update_1.log %TEMP_DIR%
del update_1.log

copy %PROGRAM_DIR%\scripts\update.bat . >NUL 2>NUL
copy %PROGRAM_DIR%\scripts\start.bat . >NUL 2>NUL

REM notwendige Programmbibliotheken installieren
venv\Scripts\pip install requests >NUL 2>NUL
venv\Scripts\pip install openpyxl >NUL 2>NUL

echo "---------------------------------------------------------"
echo "Update fertig"
echo "**********************************************************"

:ENDE
pause

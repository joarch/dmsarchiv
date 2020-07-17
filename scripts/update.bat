@echo off
REM
REM Update prüfen und ausführen
REM
call setenv.bat

IF NOT DEFINED GIT_HOME (
echo "Das automatische Update kann nicht durchgeführt werden. Bitte GIT_HOME konfigurieren."
goto ENDE
)

REM Prüfe auf Programmänderungen (Update)
cd %PROGRAM_DIR%
%GIT_HOME%\bin\git fetch
%GIT_HOME%\bin\git log --all --oneline -n1 > ../update_2.log
cd ..
fc update_1.log update_2.log
if errorlevel 1 goto UPDATE

echo "Kein Update notwendig."
goto :ENDE

:UPDATE
REM Update Source Dateien
echo "**********************************************************"
echo "Es steht ein neues Programm Update zur Verfügung"
echo "---------------------------------------------------------"
echo "Details:"
type update_2.log
echo "---------------------------------------------------------"
echo "Das Update wird jetzt eingespielt ..."
echo "---------------------------------------------------------"
cd %PROGRAM_DIR%
%GIT_HOME%\bin\git pull
%GIT_HOME%\bin\git log  --all --oneline -n1 > ../update_1.log
cd ..

copy %PROGRAM_DIR%\scripts\update.bat .
copy %PROGRAM_DIR%\scripts\start.bat .

REM notwendige Programmbibliotheken installieren
venv\Scripts\activate
pip install requests
pip install openpyxl

echo "---------------------------------------------------------"
echo "Update fertig"
echo "**********************************************************"

:ENDE
pause

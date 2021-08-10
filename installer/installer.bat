@echo off
::sets the newline character
set n=^&echo.

cd %~dp0
call :check_Permissions

if %errorlevel% EQU 0 (
call :check_python
if %errorlevel% EQU 0 (
	call :install_python
	)
if %errorlevel% EQU 0 (
	call :install_pip
	call :setup_pip
	) else ( call :setup_pip )
call :moveCelex
call :createShortCut
call :addToPath

echo Tutto e' stato configurato correttamente, premi qualsiasi tasto per chiudere la console :D, la finestra si chiuderà in 10 secondi...
timeout /T 10 >nul
)else (
echo Premi qualsiasi tasto per uscire, la finestra si chiuderà in 10 secondi...
timeout /T 10 >nul
)
del get-pip.py >nul 2>nul
del python-installer.exe >nul 2>nul
del create_shortcut.ps1 >nul 2>nul
(goto) 2>nul & del "%~f0"


::Checks if the console has been opened with admin privileges

:check_Permissions
    echo L'installazione richiede permessi di amministrazione. Controllando i permessi...%n%

    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Successo: Permessi di amministratore confermati.%n%
	EXIT /B 0
    ) else (
        echo Permessi d'amministratore non esistenti. Riavvia la console con permessi d'amministratore [ Start - cmd - tasto destro - esegui come amminsitratore ]%n%
	EXIT /B 1
    )

::checks if python is already installed on the system

:check_python
echo Controllando se Python e' gia' installato%n%
python -c "import os, sys; print(os.path.dirname(sys.executable))" >nul 2>&1
if %errorlevel% NEQ 0 (
echo Python e' gia' installato :O%n%
EXIT /B 1
) else (
echo Python non e' installato :(%n%
EXIT /B 0
)


::installs python

:install_python
echo Installando python, ci vorranno un paio di minuti al massimo ;)%n%
python-installer.exe /quiet PrependPath=1 Include_test=0
if %errorlevel% EQU 0 (
echo Python installato con successo :D%n%
exit /b 0)
else (
echo Ci sono stati problemi con l'installazione di python, riprova :(%n%
exit /b 1
)

::check if pip is installed, and if it is, upgrade it to the latest version

:check_pip
echo Controllando se pip è già installato %n%
pip help >nul 2>&1
if %errorlevel% NEQ 0 (
echo pip non e' installato.%n%
exit /b 0
) else (
echo Aggiornando pip alla versione piu' recente...%n%
call :upgrade_pip
exit /b 1
)

::installs pip

:install_pip
echo Sto installando pip per te :D
echo/
python get-pip.py >nul 2>&1
EXIT /B 0

::upgrades pip to the latest version

:upgrade_pip
python -m pip install --upgrade pip >nul 2>&1
echo/
EXIT /B 0

::setups pip for excel usage

:setup_pip
pip install xlwt >nul 2>&1
pip install openpyxl >nul 2>&1
pip install pysimplegui >nul 2>&1
pip install xlrd >nul 2>&1
pip install pandas >nul 2>&1
EXIT /B 0

:moveCelex
::Moves celex.py to C:/Program Files/Celex/
mkdir "C:/Program Files/Celex" >nul 2>&1
move /Y "%~dp0..\Celex.py" "C:/Program Files/Celex" >nul 2>&1
move /Y "%~dp0..\icon.ico" "C:/Program Files/Celex" >nul 2>&1
exit /B 0

:addToPath
::Add program to path so it can be executed from anywhere
echo C:\Program Files\Celex > %PATH%

:createShortCut
::Calls a powershell to create a shortcut
powershell.exe -ExecutionPolicy Bypass -Command "%~dp0create_shortcut.ps1"
exit /B 0

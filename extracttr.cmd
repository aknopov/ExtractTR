@echo off

SETLOCAL ENABLEEXTENSIONS

IF _%1 == _ GOTO Run
IF %1 == pack GOTO Pack
IF %1 == install GOTO Install
IF %1 == -h GOTO Help
echo Command '%1' is not supported
GOTO End

:Run
CALL venv\Scripts\activate
venv\Scripts\python main.pyw
CALL venv\Scripts\deactivate
GOTO End

:Pack
IF _%2 == _ GOTO Help
SET DEST_DIR=%2\ExtractTR
mkdir %DEST_DIR%
FOR %%I IN (*.py main.pyw requirements.txt res\extract.ico README.md %~f0) DO copy %%I %DEST_DIR% > nul
IF %ERRORLEVEL% NEQ 0 (echo Failed to copy files & GOTO End)
GOTO End

:Install
python --version > nul
IF %ERRORLEVEL% NEQ 0 winget install -e --id Python.Python.3.13
python -m ensurepip --upgrade
python -m venv venv
CALL venv\Scripts\activate
pip install -r requirements.txt
CALL venv\Scripts\deactivate
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('.\ExtractTR.lnk');$s.TargetPath='%~f0';$s.WorkingDirectory='%~dp0';$s.IconLocation='%~dp0\extract.ico';$s.WindowStyle=7;$s.Save()"
GOTO End

:Help
echo Usage: "%~n0 [pack <dir> | install | run]"
echo     pack - creates deployment package from the codebase in directory (developer action)
echo     install - installs on the target machine required components
echo     run - runs the program
GOTO End

:End
ENDLOCAL
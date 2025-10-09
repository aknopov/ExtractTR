@echo off

IF _%1 == _ echo Missing target directory! & exit

SETLOCAL ENABLEEXTENSIONS

SET DEST_DIR=%1\ExtractTR
mkdir %DEST_DIR%\res
FOR %%I IN (%~dp0\..\*.py %~dp0\..\main.pyw %~dp0\..\requirements.txt %~dp0\..\README.md %~dp0\install.cmd %~dp0\run.cmd) DO copy %%I %DEST_DIR% > nul
copy %~dp0\..\res\extract.ico %DEST_DIR%\res > nul
IF %ERRORLEVEL% NEQ 0 (echo Failed to copy some files)

ENDLOCAL
@echo off

python --version > nul
IF %ERRORLEVEL% NEQ 0 winget install -e --id Python.Python.3.13

:: Path is not updated yet, assuming location
set PYTHON=%USERPROFILE%\AppData\Local\Programs\Python\Python313\python.exe
%PYTHON% -m ensurepip --upgrade
%PYTHON% -m venv venv

CALL venv\Scripts\activate
pip install -r requirements.txt
CALL venv\Scripts\deactivate

powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('.\ExtractTR.lnk');$s.TargetPath='%~dp0\run.cmd';$s.WorkingDirectory='%~dp0';$s.IconLocation='%~dp0\res\extract.ico';$s.Description='ExtractTR';$s.WindowStyle=7;$s.Save()"

pause

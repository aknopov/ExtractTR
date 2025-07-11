@echo off

python --version > nul
IF %ERRORLEVEL% NEQ 0 winget install -e --id Python.Python.3.13

python -m ensurepip --upgrade
python -m venv venv

CALL venv\Scripts\activate
pip install -r requirements.txt
CALL venv\Scripts\deactivate

powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('.\ExtractTR.lnk');$s.TargetPath='%~dp0\run.cmd';$s.WorkingDirectory='%~dp0';$s.IconLocation='%~dp0\extract.ico';$s.WindowStyle=7;$s.Save()"


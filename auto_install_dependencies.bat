@echo off

rem Get the directory path of the batch script
cd /d %~dp0

rem Install packages from requirement.txt
python.exe -m pip install -r requirement.txt

REM Check the value of %ERRORLEVEL%
if %ERRORLEVEL% neq 0 (
    echo An error occurred. Dependency installation failed.
) else (
    echo All dependencies installed successfully. Now you can execute xero_bank_reconciler.py
)

pause
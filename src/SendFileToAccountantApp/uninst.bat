@echo off
set "KEY=HKEY_CLASSES_ROOT\*\shell\SendExpenseToAccountant"

REG QUERY "%KEY%" >nul 2>nul
if %ERRORLEVEL% equ 0 (
    REG DELETE "%KEY%" /f
    echo Registry key %KEY% has been deleted.
) else (
    echo Registry key not found. Nothing to uninstall.
)
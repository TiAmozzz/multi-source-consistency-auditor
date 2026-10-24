@echo off
setlocal EnableExtensions
title File Picker Consistency Auditor

set "PYTHONIOENCODING=utf-8"
set "PYTHONDONTWRITEBYTECODE=1"
set "PY="

if defined MSCA_PYTHON if exist "%MSCA_PYTHON%" set "PY=%MSCA_PYTHON%"
if not defined PY if exist "C:\Users\Lenovo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" set "PY=C:\Users\Lenovo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not defined PY where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"

if not defined PY goto :missing_runtime
"%PY%" "%~dp0file_picker.py"
if errorlevel 1 goto :run_failed
goto :end

:missing_runtime
echo Python was not found. Install Python 3.10 or later.
pause

:run_failed
echo.
echo The file picker failed to start. Review the error message above.
pause

:end
endlocal

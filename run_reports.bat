@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

call :resolve_python

if not defined PYEXE (
  echo Python not found. Attempting to install via winget...
  where winget >nul 2>&1
  if errorlevel 1 (
    echo winget not available. Please install Python 3 and rerun this script.
    goto :end
  )
  winget install -e --id Python.Python.3.12
  call :resolve_python
  if not defined PYEXE (
    echo Python still not found. Open a new terminal and rerun.
    goto :end
  )
)

if not exist ".venv\\Scripts\\python.exe" (
  echo Creating virtual environment...
  %PYEXE% -m venv .venv
)

echo Upgrading pip...
".venv\\Scripts\\python.exe" -m pip install --upgrade pip
echo Installing Python requirements...
".venv\\Scripts\\python.exe" -m pip install -r Scripts\\requirements.txt

set "PW_DIR=%LOCALAPPDATA%\\ms-playwright"
if not exist "%PW_DIR%\\chromium-*" (
  echo Installing Playwright Chromium...
  ".venv\\Scripts\\python.exe" -m playwright install chromium
)

set "WORKFLOW_OK=0"
set "MODE=%~1"
echo Selected mode: %MODE%
if /I "%MODE%"=="ui" (
  echo Launching desktop UI...
  ".venv\\Scripts\\python.exe" Scripts\\automation_ui.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="epic" (
  echo Running EPIC report...
  ".venv\\Scripts\\python.exe" Scripts\\epic_report.py
  if errorlevel 1 goto :run_failed
  echo Running data consolidation...
  ".venv\\Scripts\\python.exe" Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  echo Running New Biz year tabs...
  ".venv\\Scripts\\python.exe" Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  echo Running Written Business YTD vs PYTD tab...
  ".venv\\Scripts\\python.exe" Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="bignition" (
  echo Running Bignition report...
  ".venv\\Scripts\\python.exe" Scripts\\main.py
  if errorlevel 1 goto :run_failed
  echo Running data consolidation...
  ".venv\\Scripts\\python.exe" Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  echo Running New Biz year tabs...
  ".venv\\Scripts\\python.exe" Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  echo Running Written Business YTD vs PYTD tab...
  ".venv\\Scripts\\python.exe" Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="all" (
  echo Running Bignition report...
  ".venv\\Scripts\\python.exe" Scripts\\main.py
  if errorlevel 1 goto :run_failed
  echo Running EPIC report...
  ".venv\\Scripts\\python.exe" Scripts\\epic_report.py
  if errorlevel 1 goto :run_failed
  echo Running data consolidation...
  ".venv\\Scripts\\python.exe" Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  echo Running New Biz year tabs...
  ".venv\\Scripts\\python.exe" Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  echo Running Written Business YTD vs PYTD tab...
  ".venv\\Scripts\\python.exe" Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="writtenbiz" (
  echo Running Written Business YTD vs PYTD tab only...
  ".venv\\Scripts\\python.exe" Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="newbiz" (
  echo Running New Biz year tabs only...
  ".venv\\Scripts\\python.exe" Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
) else (
  echo Running full pipeline: Bignition + EPIC + Consolidation + New Biz Tabs + Written Business...
  ".venv\\Scripts\\python.exe" Scripts\\main.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" Scripts\\epic_report.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
)

set "WORKFLOW_OK=1"
goto :end

:run_failed
echo.
echo ERROR: One of the report steps failed. Review logs above.

:end
if "%WORKFLOW_OK%"=="1" echo Workflow is Complete
echo.
echo Press any key to close...
pause >nul
goto :eof

:resolve_python
set "PYEXE="
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
  set "PYEXE=py -3"
  exit /b 0
)
python -c "import sys" >nul 2>&1
if not errorlevel 1 (
  set "PYEXE=python"
  exit /b 0
)
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
  set "PYEXE=%LocalAppData%\Programs\Python\Python312\python.exe"
  exit /b 0
)
if exist "%ProgramFiles%\Python312\python.exe" (
  set "PYEXE=%ProgramFiles%\Python312\python.exe"
  exit /b 0
)
exit /b 0

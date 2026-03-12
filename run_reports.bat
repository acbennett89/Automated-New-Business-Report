@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "LOCAL_PYTHON_INSTALLER=%~dp0python-manager-26.0.msix"
set "PYTHONUNBUFFERED=1"
call :resolve_python

if not defined PYEXE (
  if exist "%LOCAL_PYTHON_INSTALLER%" (
    echo Python not found. Attempting install from bundled MSIX...
    call :install_python_from_msix "%LOCAL_PYTHON_INSTALLER%"
    call :resolve_python
  )
  if not defined PYEXE (
    echo Python not found. Attempting to install via winget...
    where winget >nul 2>&1
    if errorlevel 1 (
      echo winget not available. Please install Python and rerun this script.
      goto :end
    )
    call :install_python_with_winget
    call :resolve_python
  )
  if not defined PYEXE (
    echo Python still not found. Finish any other installer or reboot, then rerun.
    goto :end
  )
)

if not exist ".venv\\Scripts\\python.exe" (
  echo Creating virtual environment...
  %PYEXE% -m venv .venv
)

echo Upgrading pip...
".venv\\Scripts\\python.exe" -u -m pip install --upgrade pip
echo Installing Python requirements...
".venv\\Scripts\\python.exe" -u -m pip install -r Scripts\\requirements.txt

set "PW_DIR=%LOCALAPPDATA%\\ms-playwright"
if not exist "%PW_DIR%\\chromium-*" (
  echo Installing Playwright Chromium...
  ".venv\\Scripts\\python.exe" -u -m playwright install chromium
)

set "WORKFLOW_OK=0"
set "MODE=%~1"
echo Selected mode: %MODE%
if /I "%MODE%"=="ui" (
  echo Launching desktop UI...
  ".venv\\Scripts\\python.exe" -u Scripts\\automation_ui.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="epic" (
  echo Running EPIC report...
  ".venv\\Scripts\\python.exe" -u Scripts\\epic_report.py
  if errorlevel 1 goto :run_failed
  echo Running data consolidation...
  ".venv\\Scripts\\python.exe" -u Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  echo Running New Biz year tabs...
  ".venv\\Scripts\\python.exe" -u Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  echo Running Written Business YTD vs PYTD tab...
  ".venv\\Scripts\\python.exe" -u Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="bignition" (
  echo Running Bignition report...
  ".venv\\Scripts\\python.exe" -u Scripts\\main.py
  if errorlevel 1 goto :run_failed
  echo Running data consolidation...
  ".venv\\Scripts\\python.exe" -u Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  echo Running New Biz year tabs...
  ".venv\\Scripts\\python.exe" -u Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  echo Running Written Business YTD vs PYTD tab...
  ".venv\\Scripts\\python.exe" -u Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="all" (
  echo Running Bignition report...
  ".venv\\Scripts\\python.exe" -u Scripts\\main.py
  if errorlevel 1 goto :run_failed
  echo Running EPIC report...
  ".venv\\Scripts\\python.exe" -u Scripts\\epic_report.py
  if errorlevel 1 goto :run_failed
  echo Running data consolidation...
  ".venv\\Scripts\\python.exe" -u Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  echo Running New Biz year tabs...
  ".venv\\Scripts\\python.exe" -u Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  echo Running Written Business YTD vs PYTD tab...
  ".venv\\Scripts\\python.exe" -u Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="writtenbiz" (
  echo Running Written Business YTD vs PYTD tab only...
  ".venv\\Scripts\\python.exe" -u Scripts\\written_business_ytd.py
  if errorlevel 1 goto :run_failed
) else if /I "%MODE%"=="newbiz" (
  echo Running New Biz year tabs only...
  ".venv\\Scripts\\python.exe" -u Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
) else (
  echo Running full pipeline: Bignition + EPIC + Consolidation + New Biz Tabs + Written Business...
  ".venv\\Scripts\\python.exe" -u Scripts\\main.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" -u Scripts\\epic_report.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" -u Scripts\\data_consolidation.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" -u Scripts\\new_biz_tabs.py
  if errorlevel 1 goto :run_failed
  ".venv\\Scripts\\python.exe" -u Scripts\\written_business_ytd.py
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

:install_python_from_msix
set "MSIX_PATH=%~f1"
echo Installing Python from %~nx1...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Add-AppxPackage -LiteralPath '%MSIX_PATH%' -ForceApplicationShutdown -ErrorAction Stop ^| Out-Null; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
set "MSIX_RC=%ERRORLEVEL%"
call :wait_for_python 60
if defined PYEXE exit /b 0
if not "%MSIX_RC%"=="0" (
  echo Bundled MSIX install returned exit code %MSIX_RC%.
)
exit /b %MSIX_RC%

:install_python_with_winget
set "WINGET_RC=1"
for /L %%I in (1,1,4) do (
  echo winget install attempt %%I/4...
  winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
  set "WINGET_RC=!ERRORLEVEL!"
  call :wait_for_python 120
  if defined PYEXE exit /b 0
  if "!WINGET_RC!"=="1618" (
    echo Another installation is in progress. Waiting 30 seconds before retry...
    timeout /t 30 /nobreak >nul
  ) else (
    echo winget returned exit code !WINGET_RC!.
    exit /b !WINGET_RC!
  )
)
echo ERROR: winget remained blocked by another installation.
exit /b %WINGET_RC%

:wait_for_python
set "PY_WAIT_SECONDS=%~1"
echo Waiting for Python install to finish...
for /L %%I in (1,1,!PY_WAIT_SECONDS!) do (
  call :resolve_python
  if defined PYEXE exit /b 0
  timeout /t 1 /nobreak >nul
)
exit /b 0

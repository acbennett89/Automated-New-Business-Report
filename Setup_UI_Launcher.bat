@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

set "SETUP_OK=0"
set "LAUNCH_UI=1"
set "LOCAL_PYTHON_INSTALLER=%~dp0python-manager-26.0.msix"
if /I "%~1"=="--no-launch" set "LAUNCH_UI=0"

echo [Setup] Preparing New Biz Report Automation UI launcher...

call :resolve_python

if not defined PYEXE (
  if exist "%LOCAL_PYTHON_INSTALLER%" (
    echo [Setup] Python not found. Attempting install from bundled MSIX...
    call :install_python_from_msix "%LOCAL_PYTHON_INSTALLER%"
    call :resolve_python
  )
  if not defined PYEXE (
    echo [Setup] Python not found. Attempting install via winget...
    where winget >nul 2>&1
    if errorlevel 1 (
      echo [Setup] ERROR: winget is not available. Install Python and rerun this script.
      goto :failed
    )
    call :install_python_with_winget
    call :resolve_python
  )
  if not defined PYEXE (
    echo [Setup] ERROR: Python still not found. Finish any other installer or reboot, then rerun.
    goto :failed
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo [Setup] Creating virtual environment...
  %PYEXE% -m venv .venv
  if errorlevel 1 (
    echo [Setup] ERROR: Failed to create virtual environment.
    goto :failed
  )
) else (
  echo [Setup] Virtual environment already exists.
)

echo [Setup] Upgrading pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
  echo [Setup] ERROR: pip upgrade failed.
  goto :failed
)

if not exist "Scripts\requirements.txt" (
  echo [Setup] ERROR: Missing requirements file at Scripts\requirements.txt.
  goto :failed
)

echo [Setup] Installing Python requirements...
".venv\Scripts\python.exe" -m pip install -r Scripts\requirements.txt
if errorlevel 1 (
  echo [Setup] ERROR: requirements install failed.
  goto :failed
)

set "PW_DIR=%LOCALAPPDATA%\ms-playwright"
if not exist "%PW_DIR%\chromium-*" (
  echo [Setup] Installing Playwright Chromium...
  ".venv\Scripts\python.exe" -m playwright install chromium
  if errorlevel 1 (
    echo [Setup] ERROR: Playwright Chromium install failed.
    goto :failed
  )
) else (
  echo [Setup] Playwright Chromium already installed.
)

set "SETUP_OK=1"
echo [Setup] UI launcher setup is complete.

if "%LAUNCH_UI%"=="0" goto :end

if exist ".venv\Scripts\pythonw.exe" if exist "Scripts\automation_ui.py" (
  echo [Setup] Launching desktop UI...
  start "" ".venv\Scripts\pythonw.exe" "Scripts\automation_ui.py"
  goto :end
)

if exist ".venv\Scripts\python.exe" if exist "Scripts\automation_ui.py" (
  echo [Setup] Launching desktop UI...
  start "" ".venv\Scripts\python.exe" "Scripts\automation_ui.py"
  goto :end
)

echo [Setup] ERROR: automation_ui.py was not found after setup.
set "SETUP_OK=0"
goto :failed

:failed
echo.
echo [Setup] Setup did not complete successfully.
echo Press any key to close...
pause >nul

:end
if "%SETUP_OK%"=="1" exit /b 0
exit /b 1

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
echo [Setup] Installing Python from %~nx1...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Add-AppxPackage -LiteralPath '%MSIX_PATH%' -ForceApplicationShutdown -ErrorAction Stop ^| Out-Null; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
set "MSIX_RC=%ERRORLEVEL%"
call :wait_for_python 60
if defined PYEXE exit /b 0
if not "%MSIX_RC%"=="0" (
  echo [Setup] Bundled MSIX install returned exit code %MSIX_RC%.
)
exit /b %MSIX_RC%

:install_python_with_winget
set "WINGET_RC=1"
for /L %%I in (1,1,4) do (
  echo [Setup] winget install attempt %%I/4...
  winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
  set "WINGET_RC=!ERRORLEVEL!"
  call :wait_for_python 120
  if defined PYEXE exit /b 0
  if "!WINGET_RC!"=="1618" (
    echo [Setup] Another installation is in progress. Waiting 30 seconds before retry...
    timeout /t 30 /nobreak >nul
  ) else (
    echo [Setup] winget returned exit code !WINGET_RC!.
    exit /b !WINGET_RC!
  )
)
echo [Setup] ERROR: winget remained blocked by another installation.
exit /b %WINGET_RC%

:wait_for_python
set "PY_WAIT_SECONDS=%~1"
echo [Setup] Waiting for Python install to finish...
for /L %%I in (1,1,!PY_WAIT_SECONDS!) do (
  call :resolve_python
  if defined PYEXE exit /b 0
  timeout /t 1 /nobreak >nul
)
exit /b 0

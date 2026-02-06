@echo off
setlocal enableextensions

set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
set "CONDA_CMD="

where conda >nul 2>nul
if not errorlevel 1 set "CONDA_CMD=conda"

if not defined CONDA_CMD (
  for %%I in (
    "%USERPROFILE%\miniconda3\condabin\conda.bat"
    "%USERPROFILE%\anaconda3\condabin\conda.bat"
    "%ProgramData%\miniconda3\condabin\conda.bat"
    "%ProgramData%\anaconda3\condabin\conda.bat"
  ) do (
    if exist "%%~I" (
      set "CONDA_CMD=%%~I"
      goto :conda_found
    )
  )
)

:conda_found
if not defined CONDA_CMD (
  echo [ERROR] conda is not found. Install Miniconda/Anaconda or run from Anaconda Prompt.
  pause
  exit /b 1
)

pushd "%REPO_ROOT%"
call "%CONDA_CMD%" run -n pseudo-semantic-bridge streamlit run web_app.py
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [ERROR] Failed to start Streamlit. Check whether the environment 'pseudo-semantic-bridge' exists.
  pause
  exit /b %EXIT_CODE%
)

exit /b 0

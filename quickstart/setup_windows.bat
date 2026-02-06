@echo off
setlocal enableextensions

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%.."

where conda >nul 2>nul
if errorlevel 1 (
  echo [ERROR] conda is not found in PATH.
  echo Please open "Anaconda Prompt" and rerun this file, or add conda to PATH.
  popd
  pause
  exit /b 1
)

conda env create -f environment.yml
if errorlevel 1 (
  echo.
  echo [ERROR] Failed to create the conda environment.
  popd
  pause
  exit /b 1
)

popd
echo Setup done.
pause

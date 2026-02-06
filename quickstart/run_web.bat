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

conda run -n pseudo-semantic-bridge streamlit run web_app.py
if errorlevel 1 (
  echo.
  echo [ERROR] Failed to start the Streamlit server.
  echo Try running "conda env list" to confirm the environment exists.
  popd
  pause
  exit /b 1
)

popd

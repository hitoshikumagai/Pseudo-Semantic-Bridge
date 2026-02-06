@echo off
setlocal enableextensions

set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
set "CONDA_CMD="
set "ENV_NAME=pseudo-semantic-bridge"

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
set "ENV_LIST_FILE=%TEMP%\psb_conda_env_list_%RANDOM%%RANDOM%.txt"
call :run_conda env list > "%ENV_LIST_FILE%" 2>nul
if errorlevel 1 (
  echo [ERROR] Failed to query conda environments.
  if exist "%ENV_LIST_FILE%" del /q "%ENV_LIST_FILE%" >nul 2>nul
  popd
  pause
  exit /b 1
)

findstr /R /B /C:"%ENV_NAME% " "%ENV_LIST_FILE%" >nul
set "ENV_EXISTS=%ERRORLEVEL%"
if exist "%ENV_LIST_FILE%" del /q "%ENV_LIST_FILE%" >nul 2>nul

if not "%ENV_EXISTS%"=="0" (
  echo [INFO] Creating conda environment "%ENV_NAME%" from environment.yml...
  call :run_conda env create -f environment.yml
  if errorlevel 1 (
    echo.
    echo [ERROR] Failed to create conda environment.
    popd
    pause
    exit /b 1
  )
)

call :run_conda run -n %ENV_NAME% streamlit run web_app.py
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [ERROR] Failed to start Streamlit. Check whether the environment 'pseudo-semantic-bridge' exists.
  pause
  exit /b %EXIT_CODE%
)

exit /b 0

:run_conda
if /I "%CONDA_CMD%"=="conda" (
  call conda %*
) else (
  call "%CONDA_CMD%" %*
)
exit /b %ERRORLEVEL%

@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%.."

call conda env create -f environment.yml
call conda activate pseudo-semantic-bridge

popd
echo Setup done.
pause

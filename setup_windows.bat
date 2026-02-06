@echo off
setlocal

call conda env create -f environment.yml
call conda activate pseudo-semantic-bridge

echo Setup done.
pause

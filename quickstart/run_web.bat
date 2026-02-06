@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%.."

call conda activate pseudo-semantic-bridge
streamlit run web_app.py

popd

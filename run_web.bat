@echo off
setlocal

call conda activate pseudo-semantic-bridge
streamlit run web_app.py

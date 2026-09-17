@echo off
cd /d "%~dp0"
echo ===================================================
echo Starting Local Report Merger Streamlit Application...
echo Browser will automatically open at http://localhost:8501
echo Press Ctrl+C in this window to stop the server anytime.
echo ===================================================
"pos-report-merger\.venv\Scripts\python.exe" -m streamlit run streamlit_app.py
pause

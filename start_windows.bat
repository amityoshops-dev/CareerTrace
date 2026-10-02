@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\activate.bat py -3 -m venv .venv
call .venv\Scripts\activate.bat
if exist requirements.txt pip install -q -r requirements.txt
pip install -q fastapi uvicorn python-dotenv requests pypdf python-docx openpyxl pywebview langchain-google-genai langgraph
if not exist data mkdir data
python desktop.py
pause

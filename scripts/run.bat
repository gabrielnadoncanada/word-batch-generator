@echo off
REM scripts\run.bat [limit]
call .venv\Scripts\activate.bat
python generate_and_pdf.py

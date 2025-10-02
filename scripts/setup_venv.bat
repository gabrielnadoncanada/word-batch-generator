@echo off
REM scripts\setup_venv.bat
set PYTHON=python
echo Creating venv...
%PYTHON% -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
echo Done. Activate later with: call .venv\Scripts\activate.bat

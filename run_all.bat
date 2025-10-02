@echo off
REM run_all.bat - End-to-end setup & run on Windows (CMD)
setlocal ENABLEDELAYEDEXPANSION

echo === Word Batch Generator: Setup & Run ===

REM 1) Ensure Python is available
where python >nul 2>nul
IF ERRORLEVEL 1 (
  echo [ERROR] Python n'est pas detecte dans PATH. Installe Python 3.x et relance.
  exit /b 1
)

REM 2) Create venv if missing
if not exist .venv (
  echo [INFO] Creation de l'environnement virtuel .venv
  python -m venv .venv
  if ERRORLEVEL 1 (
    echo [ERROR] Echec creation venv.
    exit /b 1
  )
)

REM 3) Activate venv
call .venv\Scripts\activate.bat
if ERRORLEVEL 1 (
  echo [ERROR] Echec activation venv.
  exit /b 1
)

REM 4) Upgrade pip and install deps
python -m pip install --upgrade pip
if ERRORLEVEL 1 (
  echo [ERROR] Echec upgrade pip.
  exit /b 1
)

pip install -r requirements.txt
if ERRORLEVEL 1 (
  echo [ERROR] Echec installation dependances (requirements.txt).
  exit /b 1
)

REM 5) Run generation (limit optional first arg, default 20)
echo [INFO] Generation selon le CSV (toutes les lignes valides)
python generate_and_pdf.py
set RET=%ERRORLEVEL%
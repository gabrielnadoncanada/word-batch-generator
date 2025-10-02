# run_all.ps1 - End-to-end setup & run on Windows (PowerShell)

Write-Host "=== Word Batch Generator: Setup & Run ==="

# 1) Check Python
$pythonPath = (Get-Command $PythonExe -ErrorAction SilentlyContinue)
if (-not $pythonPath) {
  Write-Error "[ERROR] Python n'est pas detecte dans PATH. Installe Python 3.x."
  exit 1
}

# 2) Create venv if missing
if (-not (Test-Path ".\.venv")) {
  Write-Host "[INFO] Creation venv .venv"
  & $PythonExe -m venv .venv
  if ($LASTEXITCODE -ne 0) { exit 1 }
}

# 3) Activate venv
. .\.venv\Scripts\Activate.ps1

# 4) Install deps
python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { exit 1 }

pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { exit 1 }

# 5) Run generation
Write-Host "[INFO] Generation selon le CSV (toutes les lignes valides)..."
python generate_and_pdf.py
$ret = $LASTEXITCODE

if ($ret -eq 0) {
  Write-Host "[OK] Generation et conversion PDF terminees."
} elseif ($ret -eq 2) {
  Write-Warning "docx2pdf a besoin de Microsoft Word. Alternative: LibreOffice headless."
  Write-Host "soffice --headless --convert-to pdf --outdir out/pdf out/docx/*.docx"
} else {
  Write-Warning "Script Python a retourne le code $ret."
}

Write-Host "Sorties: out/docx/ et out/pdf/"
exit $ret

# scripts/setup_venv.ps1
param(
  [string]$PythonExe = "python"
)
Write-Host ">> Création de l'environnement virtuel .venv"
& $PythonExe -m venv .venv
Write-Host ">> Activation de l'environnement"
& .\.venv\Scripts\Activate.ps1
Write-Host ">> Upgrade pip"
pip install --upgrade pip
Write-Host ">> Installation des dépendances"
pip install -r requirements.txt
Write-Host "OK. Pour activer plus tard: .\.venv\Scripts\Activate.ps1"

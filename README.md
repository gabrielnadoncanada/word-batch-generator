# Word Batch Generator (DOCX -> PDF)

Automatise la génération de **N documents Word** à partir d'un **modèle** puis la **conversion en PDF**, avec aucune fusion : un PDF par entrée.

- **OS ciblé** : Windows (docx2pdf utilise Microsoft Word via COM).
- **Alternative** : LibreOffice headless si Word n'est pas installé.
- **Placeholder attendu dans le modèle** : `{{VENDEUR}}`

## Structure du projet

```
word-batch-generator/
  data/
    entrepreneurs.csv         # liste des entreprises avec email (colonnes 'nom', 'email')
  out/
    docx/                     # DOCX générés
    pdf/                      # PDF générés
  scripts/
    setup_venv.ps1            # création/installation venv (PowerShell)
    run.ps1                   # exécution (PowerShell)
    setup_venv.bat            # variante .bat
    run.bat                   # variante .bat
  templates/
    modele.docx               # modèle Word (placeholder {{VENDEUR}})
  generate_and_pdf.py         # script principal
  requirements.txt
  .gitignore
  README.md
```

## Prérequis

- **Python 3.10+** (idéalement)
- **Microsoft Word** installé (pour la conversion PDF via `docx2pdf`)
- **PowerShell** ou **CMD**

> Si tu n'as pas Word, regarde **Option C** ci-dessous (LibreOffice).

## Installation rapide (Windows, PowerShell)

Depuis le dossier `word-batch-generator` :

```powershell
# 1) Crée le venv et installe les dépendances
.\scripts\setup_venv.ps1

# 2) (optionnel) Activer le venv dans cette session si pas déjà actif
.\.venv\Scripts\Activate.ps1

# 3) Exécuter la génération (selon le nombre de lignes du CSV)
.\scripts
un.ps1 -
```

### Variante CMD (.bat)

```bat
REM 1) Setup
scripts\setup_venv.bat

REM 2) Run (20 documents par défaut)
scripts
un.bat 20
```

## Personnalisation

### 1) Éditer la liste des entreprises

Modifie `data/entrepreneurs.csv` (colonne `nom`). Le script lit **jusqu'à 20** lignes par défaut (paramètre `--limit`).

### 2) Adapter le modèle Word

Ouvre `templates/modele.docx` et assure-toi que le placeholder **exact** existe : `{{VENDEUR}}`.
- Si Word fragmente le placeholder sur plusieurs *runs* (style appliqué au milieu), le script dispose d'un **mode secours** qui remplace le texte du paragraphe (avec petite perte de style local).

### 3) Changer le placeholder

Si ton modèle utilise un autre token, modifie la constante `PLACEHOLDER` dans `generate_and_pdf.py`.

### 4) Fusionner tous les PDF

La fusion est désactivée (`MERGE_ENABLED = False`) : un PDF par entrée. (`out/pdf/bundle_entrepreneurs.pdf`).

## Options d'exécution

```powershell
# Limiter le nombre de documents générés (ex: 10)
.\scripts
un.ps1 -Limit 10

# Ou en direct (venv actif)
python generate_and_pdf.py 10
```

## Option C — Sans Microsoft Word (LibreOffice headless)

Si Word n'est pas disponible, tu peux convertir avec LibreOffice :
1. Installe **LibreOffice**.
2. Après génération des `.docx`, lance :
   ```bash
   soffice --headless --convert-to pdf --outdir out/pdf out/docx/*.docx
   ```

> Remarque : la fidélité visuelle peut légèrement différer selon les polices/mises en page.

## Dépannage

- **docx2pdf error / COM** : Assure-toi que Microsoft Word est bien installé et ouvert au moins une fois.
- **Polices manquantes** : Installe les polices utilisées par `modele.docx` pour éviter les substitutions.
- **Placeholder non remplacé** : Vérifie l'orthographe exacte `{{VENDEUR}}` et qu'il ne soit pas scindé par du style. Le mode secours tente de le corriger.
- **PDF fusionné manquant** : Mets `MERGE_ENABLED = True` (par défaut).

## Sécurité & traçabilité

- Sorties nommées comme `01_Entreprise_Alpha.docx/.pdf` (ordre stable).
- Pas de données sensibles écrites en clair hors de `data/`.
- `.gitignore` exclut `out/` et `.venv/`.

---

**Tu peux remplacer `templates/modele.docx` par ton vrai modèle et `data/entrepreneurs.csv` par ta liste, sans toucher au code.**


## Format CSV (mis à jour)

Le fichier `data/entrepreneurs.csv` doit contenir :
```csv
nom,email
Entreprise Alpha,alpha@example.com
Entreprise Beta,beta@example.com
...
```

- La colonne **nom** est obligatoire.
- La colonne **email** est optionnelle et sert **uniquement** à concaténer l'adresse dans le **nom du fichier PDF** (pas dans le contenu du document).
- Exemple de sortie : `01_Entreprise_Alpha_alpha_at_example.com.pdf`

# Générateur de Documents Word - Version Refactorisée

## Structure du Projet

```
word-batch-generator/
├── config.py                 # Configuration centralisée
├── main.py                   # Point d'entrée principal
├── document_generator.py     # Générateur de documents Word
├── email_sender.py          # Gestionnaire d'envoi d'emails
├── outlook_utils.py         # Utilitaires Outlook
├── file_utils.py            # Utilitaires de gestion des fichiers
├── validators.py            # Validateurs de données
├── logger_config.py         # Configuration du logging
├── tests/                   # Tests unitaires
│   ├── __init__.py
│   ├── test_validators.py
│   └── test_document_generator.py
├── templates/               # Modèles Word
│   └── modele.docx
├── data/                    # Données CSV
│   └── entrepreneurs.csv
└── out/                     # Fichiers de sortie
    ├── docx/
    ├── pdf/
    └── logs/
```

## Améliorations Apportées

### 1. **Séparation des Responsabilités**

- **`config.py`** : Configuration centralisée
- **`document_generator.py`** : Logique de génération de documents
- **`email_sender.py`** : Gestion de l'envoi d'emails
- **`outlook_utils.py`** : Intégration avec Outlook
- **`file_utils.py`** : Utilitaires de gestion des fichiers
- **`validators.py`** : Validation des données

### 2. **Programmation Orientée Objet**

- Classes `DocumentGenerator`, `EmailSender`, `WordBatchGenerator`
- Encapsulation des fonctionnalités
- Réutilisabilité du code

### 3. **Gestion d'Erreurs Améliorée**

- Validation des données d'entrée
- Gestion des erreurs spécifiques
- Logging structuré et informatif

### 4. **Configuration Centralisée**

- Tous les paramètres dans `config.py`
- Facilite la maintenance et les modifications

### 5. **Tests Unitaires**

- Tests pour les validateurs
- Tests pour le générateur de documents
- Couverture des fonctionnalités critiques

## Utilisation

### Exécution Simple

```bash
python main.py
```

### Exécution avec Tests

```bash
# Exécuter tous les tests
python -m pytest tests/

# Exécuter un test spécifique
python -m pytest tests/test_validators.py
```

## Configuration

Modifiez `config.py` pour ajuster :

- Chemins des fichiers
- Paramètres d'email
- Configuration de signature
- Paramètres de retry
- Niveau de logging

## Dépendances

Les dépendances restent les mêmes :

- `python-docx`
- `docx2pdf`
- `pywin32` (Windows uniquement)

Voir `requirements.txt` pour la liste complète.

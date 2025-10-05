# Configuration SMTP - Guide d'installation

## Vue d'ensemble

Ce guide vous explique comment configurer le système d'envoi d'emails via SMTP pour remplacer l'intégration Outlook qui peut être peu fiable.

## Avantages du SMTP

- ✅ Plus fiable que l'intégration Outlook COM
- ✅ Fonctionne avec n'importe quel fournisseur d'email
- ✅ Pas de dépendance à Microsoft Outlook
- ✅ Configuration simple et directe
- ✅ Support TLS/SSL sécurisé

## Configuration

### 1. Modifier config.py

Ouvrez `config.py` et configurez les paramètres SMTP :

```python
# Configuration SMTP (remplace Outlook)
USE_SMTP = True  # Utiliser SMTP au lieu d'Outlook
SMTP_SERVER = "smtp.gmail.com"  # Changez selon votre fournisseur
SMTP_PORT = 587  # 587 pour TLS, 465 pour SSL
SMTP_USERNAME = "votre@email.com"  # Votre email
SMTP_PASSWORD = "votre_mot_de_passe"  # Votre mot de passe
SMTP_USE_TLS = True  # True pour TLS, False pour SSL
SMTP_USE_SSL = False  # True pour SSL, False pour TLS
```

### 2. Configuration par fournisseur

#### Gmail

```python
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USE_TLS = True
SMTP_USE_SSL = False
# Utilisez un mot de passe d'application, pas votre mot de passe normal
```

#### Outlook/Hotmail

```python
SMTP_SERVER = "smtp-mail.outlook.com"
SMTP_PORT = 587
SMTP_USE_TLS = True
SMTP_USE_SSL = False
```

#### Yahoo

```python
SMTP_SERVER = "smtp.mail.yahoo.com"
SMTP_PORT = 587
SMTP_USE_TLS = True
SMTP_USE_SSL = False
```

#### Autres fournisseurs

Consultez la documentation de votre fournisseur d'email pour les paramètres SMTP.

### 3. Sécurité

**Important** : Pour Gmail, vous devez utiliser un "mot de passe d'application" :

1. Activez l'authentification à 2 facteurs sur votre compte Google
2. Allez dans "Sécurité" > "Mots de passe d'application"
3. Générez un mot de passe d'application pour "Mail"
4. Utilisez ce mot de passe dans `SMTP_PASSWORD`

## Test de la configuration

### 1. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 2. Tester la connexion

```bash
python test_smtp.py
```

Ce script va :

- Vérifier votre configuration
- Tester la connexion SMTP
- Afficher les erreurs s'il y en a

### 3. Test d'envoi (optionnel)

Décommentez la ligne `test_send_email()` dans `test_smtp.py` pour tester l'envoi réel d'un email.

## Utilisation

Le système fonctionne automatiquement. Quand `USE_SMTP = True` dans `config.py`, l'application utilisera SMTP au lieu d'Outlook.

### Retour à Outlook

Si vous voulez revenir à Outlook temporairement :

```python
USE_SMTP = False
```

## Dépannage

### Erreur d'authentification

- Vérifiez votre nom d'utilisateur et mot de passe
- Pour Gmail, utilisez un mot de passe d'application
- Vérifiez que l'authentification à 2 facteurs est activée

### Erreur de connexion

- Vérifiez les paramètres de serveur et port
- Vérifiez votre connexion internet
- Certains réseaux bloquent les ports SMTP

### Emails non reçus

- Vérifiez le dossier spam
- Vérifiez que l'adresse d'expéditeur est correcte
- Vérifiez les paramètres CC/BCC

## Support

En cas de problème, vérifiez les logs dans `out/logs/mail.log` pour plus de détails.

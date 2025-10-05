# Configuration des Variables d'Environnement

## Vue d'ensemble

Ce projet utilise des variables d'environnement pour sécuriser les informations sensibles comme les mots de passe SMTP. Cela évite de commiter des données sensibles sur GitHub.

## Configuration

### 1. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 2. Configurer les variables d'environnement

1. **Copiez le fichier template :**
   ```bash
   copy env.example .env
   ```

2. **Éditez le fichier `.env`** avec vos vraies valeurs :
   ```env
   # Configuration SMTP - Variables sensibles
   SMTP_SERVER=secure.emailsrvr.com
   SMTP_PORT=465
   SMTP_USERNAME=votre@email.com
   SMTP_PASSWORD=votre_mot_de_passe
   FROM_ACCOUNT=votre@email.com
   ```

### 3. Variables disponibles

| Variable | Description | Exemple |
|----------|-------------|---------|
| `SMTP_SERVER` | Serveur SMTP | `secure.emailsrvr.com` |
| `SMTP_PORT` | Port SMTP | `465` |
| `SMTP_USERNAME` | Nom d'utilisateur SMTP | `votre@email.com` |
| `SMTP_PASSWORD` | Mot de passe SMTP | `votre_mot_de_passe` |
| `FROM_ACCOUNT` | Adresse email d'expéditeur | `votre@email.com` |
| `CC` | Copie carbone (optionnel) | `cc@example.com` |
| `BCC` | Copie carbone cachée (optionnel) | `bcc@example.com` |
| `SMTP_USE_TLS` | Utiliser TLS | `false` |
| `SMTP_USE_SSL` | Utiliser SSL | `true` |

## Sécurité

### ✅ **Ce qui est sécurisé :**
- Le fichier `.env` est dans `.gitignore`
- Les mots de passe ne sont plus dans le code
- Les informations sensibles ne sont pas commitées

### ⚠️ **Important :**
- **Ne jamais commiter le fichier `.env`**
- **Ne jamais partager le fichier `.env`**
- **Utiliser `env.example` comme template public**

## Dépannage

### Le fichier .env n'est pas lu
- Vérifiez que le fichier `.env` existe dans le répertoire racine
- Vérifiez que `python-dotenv` est installé : `pip install python-dotenv`

### Variables non trouvées
- Vérifiez l'orthographe des variables dans `.env`
- Vérifiez qu'il n'y a pas d'espaces autour du `=`
- Redémarrez l'application après modification

### Exemple de fichier .env complet
```env
# Configuration SMTP - Variables sensibles
SMTP_SERVER=secure.emailsrvr.com
SMTP_PORT=465
SMTP_USERNAME=gabriel@dilamco.com
SMTP_PASSWORD=Gab!2025
FROM_ACCOUNT=gabriel@dilamco.com
CC=
BCC=
SMTP_USE_TLS=false
SMTP_USE_SSL=true
```

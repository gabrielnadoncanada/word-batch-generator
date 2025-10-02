# -*- coding: utf-8 -*-
"""
Utilitaires pour l'intégration avec Microsoft Outlook
"""
import logging
import os
import time
import traceback
from pathlib import Path
from typing import Optional, Tuple, List
import pythoncom
import win32com.client as win32
from win32com.client import constants

from file_utils import read_text_smart


def detect_default_signature_name() -> str:
    """Détecte le nom de la signature par défaut d'Outlook."""
    try:
        import winreg
    except Exception:
        return ""
    
    for ver in ("16.0", "15.0", "14.0"):
        try:
            key_path = fr"Software\Microsoft\Office\{ver}\Common\MailSettings"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                name, _ = winreg.QueryValueEx(k, "NewSignature")
                if name:
                    return name
        except FileNotFoundError:
            continue
        except Exception:
            continue
    return ""


def load_system_signature_html() -> str:
    """Charge la signature système d'Outlook."""
    sig_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
    if not sig_dir.exists():
        return ""
    
    name = detect_default_signature_name()
    candidate = None
    
    if name:
        p = sig_dir / f"{name}.htm"
        if p.exists():
            candidate = p
        else:
            p2 = sig_dir / f"{name}.html"
            if p2.exists():
                candidate = p2
    
    if not candidate:
        files = list(sig_dir.glob("*.htm")) or list(sig_dir.glob("*.html"))
        if files:
            candidate = files[0]
    
    return read_text_smart(candidate) if candidate and candidate.exists() else ""


def debug_list_outlook_accounts():
    """Affiche la liste des comptes Outlook disponibles."""
    try:
        outlook = win32.gencache.EnsureDispatch("Outlook.Application")
        session = outlook.Session
        logging.info("Comptes Outlook détectés:")
        for acct in session.Accounts:
            logging.info(f" - DisplayName={getattr(acct, 'DisplayName', '')} | SMTP={getattr(acct, 'SmtpAddress', '')}")
    except Exception as e:
        logging.error("Échec debug_list_outlook_accounts()")
        logging.error(traceback.format_exc())


def ensure_mapi_ready():
    """Force une session MAPI opérationnelle."""
    pythoncom.CoInitialize()

    outlook = win32.gencache.EnsureDispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    
    try:
        ns.Logon("", "", False, False)
    except Exception:
        logging.warning("ns.Logon a échoué (probablement déjà logué), on continue.")

    try:
        _ = ns.GetDefaultFolder(6)  # 6=olFolderInbox
    except Exception:
        logging.warning("GetDefaultFolder(Inbox) a échoué.")
    
    return outlook, ns


def signatures_dir() -> Path:
    """Retourne le répertoire des signatures Outlook."""
    return Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"


def list_signatures() -> List[str]:
    """Liste toutes les signatures disponibles."""
    d = signatures_dir()
    if not d.exists():
        return []
    
    names = []
    for p in d.glob("*.htm"):
        names.append(p.stem)
    for p in d.glob("*.html"):
        if p.stem not in names:
            names.append(p.stem)
    
    return sorted(names)


def load_project_signature() -> str:
    """Charge la signature du projet."""
    try:
        from config import PROJECT_SIGNATURE_FILE
        if PROJECT_SIGNATURE_FILE.exists():
            with open(PROJECT_SIGNATURE_FILE, "r", encoding="utf-8") as f:
                return f.read()
    except Exception as e:
        logging.warning(f"Erreur chargement signature projet: {e}")
    
    return ""


def load_email_template() -> str:
    """Charge le template d'email du projet."""
    try:
        from config import EMAIL_TEMPLATE_FILE, FALLBACK_BODY_HTML_TEMPLATE
        if EMAIL_TEMPLATE_FILE.exists():
            with open(EMAIL_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                return f.read()
        else:
            logging.warning(f"Template d'email introuvable: {EMAIL_TEMPLATE_FILE}")
            return FALLBACK_BODY_HTML_TEMPLATE
    except Exception as e:
        logging.warning(f"Erreur chargement template email: {e}")
        return FALLBACK_BODY_HTML_TEMPLATE


def load_signature_html_by_name(name: str) -> Tuple[str, str]:
    """Charge une signature par son nom."""
    if not name:
        return "", ""
    
    d = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
    for ext in ("htm", "html"):
        path = d / f"{name}.{ext}"
        if path.exists():
            return read_text_smart(path), path.stem
    return "", ""


def attach_and_inline_signature_images(mail, signature_name: str, signature_html: str) -> str:
    """Attache les images de la signature et remplace les src par cid."""
    if not signature_name or not signature_html:
        return signature_html
    
    base = signatures_dir() / f"{signature_name}_files"
    if not base.exists():
        return signature_html

    import re

    def attach_and_replace(match):
        src = match.group(1)
        candidate = base / Path(src).name
        if not candidate.exists():
            candidate = signatures_dir() / src
            if not candidate.exists():
                return match.group(0)

        try:
            att = mail.Attachments.Add(str(candidate))
            cid = Path(candidate).name.replace(".", "_")
            att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
            return f'src="cid:{cid}"'
        except Exception:
            return match.group(0)

    new_html = re.sub(r'src=["\']([^"\']+)["\']', attach_and_replace, signature_html)
    return new_html


def send_email_via_outlook(
    to_email: str,
    subject: str,
    html_body: str,
    attachments: List[Path],
    from_account: Optional[str] = None,
    cc: str = "",
    bcc: str = "",
    use_signature: bool = True,
    signature_name: Optional[str] = None,
    use_email_template: bool = False,
    strict_embed_images: bool = False,
    max_retries: int = 5,
    delay_seconds: float = 2.0,
) -> None:
    """Envoie un email via Outlook avec gestion des erreurs et retry."""
    logging.info(f"[MAIL] Préparation -> to={to_email} | from={from_account} | subj={subject}")

    pythoncom.CoInitialize()

    try:
        outlook, ns = ensure_mapi_ready()
    except Exception as e:
        logging.error("[MAIL] ensure_mapi_ready failed: %s", e)
        logging.error(traceback.format_exc())
        raise

    # Créer l'item avec retry
    mail = None
    for attempt in range(1, max_retries + 1):
        try:
            mail = outlook.CreateItem(constants.olMailItem)
            logging.debug("[MAIL] CreateItem OK (attempt %d)", attempt)
            break
        except Exception as e:
            msg = str(e)
            logging.warning("[MAIL] CreateItem échoué (attempt %d): %s", attempt, msg)
            logging.debug(traceback.format_exc())
            if "Call was rejected by callee" in msg or "-2147418111" in msg:
                time.sleep(delay_seconds)
                continue
            raise
    
    if mail is None:
        raise RuntimeError("Outlook indisponible: CreateItem a échoué après retries.")

    # Forcer le compte d'envoi si demandé
    if from_account:
        try:
            matched = False
            for acct in ns.Session.Accounts:
                smtp = getattr(acct, "SmtpAddress", "")
                if smtp and smtp.lower() == from_account.lower():
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, acct))
                    matched = True
                    logging.info(f"[MAIL] SendUsingAccount = {smtp}")
                    break
            if not matched:
                logging.warning(f"[MAIL] Compte '{from_account}' non trouvé dans Outlook.")
        except Exception as e:
            logging.warning(f"[MAIL] Impossible de forcer le compte '{from_account}': {e}")

    # Configurer les destinataires et le contenu
    mail.To = to_email
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = subject

    # Gestion du template d'email
    if use_email_template:
        try:
            from config import USE_EMAIL_TEMPLATE
            if USE_EMAIL_TEMPLATE:
                template_html = load_email_template()
                if template_html:
                    # Remplacer les variables dans le template
                    from datetime import datetime
                    template_html = template_html.replace("{{DATE_SOUMISSION}}", datetime.now().strftime("%d/%m/%Y"))
                    final_html = template_html
                    logging.debug("[MAIL] Template d'email chargé")
                else:
                    final_html = html_body or ""
            else:
                final_html = html_body or ""
        except Exception as e:
            logging.warning(f"[MAIL] Erreur chargement template: {e}")
            final_html = html_body or ""
    else:
        final_html = html_body or ""

    # Gestion de la signature
    try:
        sig_html = ""
        effective_name = ""
        
        # Priorité: signature du projet, puis signature Outlook spécifique, puis signature système
        from config import USE_PROJECT_SIGNATURE
        if USE_PROJECT_SIGNATURE:
            sig_html = load_project_signature()
            effective_name = "project_signature"
        elif signature_name:
            sig_html, effective_name = load_signature_html_by_name(signature_name)
        elif use_signature:
            sig_html = load_system_signature_html()

        if sig_html:
            if strict_embed_images:
                sig_html = attach_and_inline_signature_images(mail, effective_name or signature_name, sig_html)
            final_html += sig_html
            logging.debug("[MAIL] Signature ajoutée: %s", (effective_name or signature_name or "system_default"))
        else:
            logging.debug("[MAIL] Aucune signature chargée")
    except Exception as e:
        logging.warning(f"[MAIL] Chargement signature a échoué: {e}")

    mail.HTMLBody = final_html

    # Ajouter les pièces jointes
    for att in attachments or []:
        try:
            mail.Attachments.Add(str(att))
            logging.debug(f"[MAIL] Attachment ajouté: {att}")
        except Exception as e:
            logging.warning(f"[MAIL] Impossible d'attacher {att}: {e}")

    # Sauvegarder avant envoi
    try:
        mail.Save()
        logging.debug("[MAIL] Draft sauvegardé")
    except Exception:
        logging.debug("[MAIL] mail.Save() a échoué (non bloquant)")

    # Envoyer avec retry
    for attempt in range(1, max_retries + 1):
        try:
            mail.Send()
            logging.info(f"[MAIL] Send OK -> {to_email}")
            return
        except Exception as e:
            msg = str(e)
            logging.warning("[MAIL] Send échoué (attempt %d): %s", attempt, msg)
            logging.debug(traceback.format_exc())
            if "Call was rejected by callee" in msg or "-2147418111" in msg:
                time.sleep(delay_seconds)
                continue
            raise

    raise RuntimeError(f"[MAIL] Échec envoi à {to_email} après {max_retries} tentatives.")

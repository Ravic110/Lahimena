"""
Authentication handler — gestion des comptes, mots de passe et sessions.

Stockage : users.json dans le répertoire racine du projet.
Sécurité  : hash SHA-256 + sel individuel par utilisateur.
Expiration: mot de passe expire après 90 jours (3 mois).
Accès     : suspension manuelle + expiration de durée d'accès.
"""

import hashlib
import json
import os
import secrets
from datetime import datetime, timedelta

# ── Constantes ────────────────────────────────────────────────────────────────

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
USERS_FILE = os.path.join(_BASE_DIR, "users.json")

PASSWORD_EXPIRY_DAYS = 90  # 3 mois

ROLES = {
    "admin":     "Administrateur — accès complet",
    "agent":     "Agent — accès clients, cotations, factures",
    "comptable": "Comptable — accès comptabilité uniquement",
}

# ── I/O ───────────────────────────────────────────────────────────────────────

def _load_users() -> list:
    if not os.path.exists(USERS_FILE):
        return []
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_users(users: list) -> bool:
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ── Cryptographie ─────────────────────────────────────────────────────────────

def _hash_password(password: str, salt: str) -> str:
    combined = (password + salt).encode("utf-8")
    return hashlib.sha256(combined).hexdigest()


def _generate_salt() -> str:
    return secrets.token_hex(16)


# ── Statut d'accès ────────────────────────────────────────────────────────────

def is_password_expired(user: dict) -> bool:
    """Retourne True si le mot de passe a expiré (>90 jours)."""
    try:
        changed = user.get("password_changed_at") or user.get("created_at", "")
        if not changed:
            return True
        dt = datetime.strptime(changed, "%Y-%m-%d %H:%M:%S")
        return (datetime.now() - dt).days >= PASSWORD_EXPIRY_DAYS
    except Exception:
        return False


def is_suspended(user: dict) -> bool:
    """Retourne True si le compte est suspendu manuellement."""
    return bool(user.get("suspended", False))


def is_access_expired(user: dict) -> bool:
    """Retourne True si la durée d'accès accordée est dépassée."""
    expires = user.get("access_expires_at", "")
    if not expires:
        return False
    try:
        dt = datetime.strptime(expires, "%Y-%m-%d")
        return datetime.now().date() > dt.date()
    except Exception:
        return False


def access_status(user: dict) -> str:
    """Retourne 'suspended' | 'access_expired' | 'pw_expired' | 'active'."""
    if is_suspended(user):
        return "suspended"
    if is_access_expired(user):
        return "access_expired"
    if is_password_expired(user):
        return "pw_expired"
    return "active"


# ── Utilisateurs ──────────────────────────────────────────────────────────────

def has_users() -> bool:
    return len(_load_users()) > 0


def get_users() -> list:
    users = _load_users()
    return [
        {
            "username":            u["username"],
            "role":                u["role"],
            "created_at":          u.get("created_at", ""),
            "password_changed_at": u.get("password_changed_at", ""),
            "suspended":           u.get("suspended", False),
            "access_expires_at":   u.get("access_expires_at", ""),
            "is_expired":          is_password_expired(u),
            "status":              access_status(u),
        }
        for u in users
    ]


def create_user(username: str, password: str, role: str,
                access_expires_at: str = "") -> tuple[bool, str]:
    """
    Crée un nouveau compte utilisateur.
    access_expires_at : date au format YYYY-MM-DD, vide = accès illimité.
    """
    username = username.strip()
    if not username:
        return False, "Le nom d'utilisateur est obligatoire."
    if len(username) < 3:
        return False, "Le nom d'utilisateur doit avoir au moins 3 caractères."
    if not password:
        return False, "Le mot de passe est obligatoire."
    if len(password) < 6:
        return False, "Le mot de passe doit avoir au moins 6 caractères."
    if role not in ROLES:
        return False, f"Rôle invalide. Valeurs acceptées : {', '.join(ROLES)}."

    users = _load_users()
    if any(u["username"].lower() == username.lower() for u in users):
        return False, f"L'utilisateur « {username} » existe déjà."

    salt = _generate_salt()
    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    users.append({
        "username":            username,
        "password_hash":       _hash_password(password, salt),
        "salt":                salt,
        "role":                role,
        "created_at":          now,
        "password_changed_at": now,
        "suspended":           False,
        "access_expires_at":   access_expires_at,
    })
    if _save_users(users):
        from utils.activity_log import log_activity
        log_activity("create_user", f"Compte créé : {username} ({role})")
        return True, ""
    return False, "Erreur lors de la sauvegarde du fichier utilisateurs."


def delete_user(username: str) -> tuple[bool, str]:
    users = _load_users()
    new_users = [u for u in users if u["username"].lower() != username.lower()]
    if len(new_users) == len(users):
        return False, f"Utilisateur « {username} » introuvable."
    if _save_users(new_users):
        from utils.activity_log import log_activity
        log_activity("delete_user", f"Compte supprimé : {username}")
        return True, ""
    return False, "Erreur lors de la sauvegarde."


def change_password(username: str, new_password: str) -> tuple[bool, str]:
    if len(new_password) < 6:
        return False, "Le mot de passe doit avoir au moins 6 caractères."
    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            salt = _generate_salt()
            u["password_hash"]       = _hash_password(new_password, salt)
            u["salt"]                = salt
            u["password_changed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if _save_users(users):
                from utils.activity_log import log_activity
                log_activity("change_password",
                             f"Mot de passe modifié pour : {username}",
                             username=username, role=u.get("role", ""))
                return True, ""
            return False, "Erreur lors de la sauvegarde."
    return False, f"Utilisateur « {username} » introuvable."


def suspend_user(username: str) -> tuple[bool, str]:
    """Suspend l'accès d'un utilisateur."""
    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            u["suspended"] = True
            if _save_users(users):
                from utils.activity_log import log_activity
                log_activity("suspend_user", f"Compte suspendu : {username}")
                return True, ""
            return False, "Erreur lors de la sauvegarde."
    return False, f"Utilisateur « {username} » introuvable."


def reactivate_user(username: str) -> tuple[bool, str]:
    """Réactive l'accès d'un utilisateur suspendu."""
    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            u["suspended"] = False
            if _save_users(users):
                from utils.activity_log import log_activity
                log_activity("reactivate_user", f"Compte réactivé : {username}")
                return True, ""
            return False, "Erreur lors de la sauvegarde."
    return False, f"Utilisateur « {username} » introuvable."


def set_access_expiry(username: str, expires_at: str) -> tuple[bool, str]:
    """
    Définit la date d'expiration d'accès (YYYY-MM-DD).
    Passer une chaîne vide pour supprimer la limite.
    """
    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            u["access_expires_at"] = expires_at
            if _save_users(users):
                from utils.activity_log import log_activity
                detail = f"Expiration fixée au {expires_at}" if expires_at \
                         else "Limite d'accès supprimée"
                log_activity("set_expiry", f"{detail} pour : {username}")
                return True, ""
            return False, "Erreur lors de la sauvegarde."
    return False, f"Utilisateur « {username} » introuvable."


# ── Authentification ──────────────────────────────────────────────────────────

def authenticate(username: str, password: str) -> tuple[bool, dict | None, str]:
    """
    Vérifie les identifiants.

    Returns:
        (True,  user_dict, "")               → succès
        (True,  user_dict, "expired")        → succès mais mot de passe expiré
        (False, None,      "suspended")      → compte suspendu
        (False, None,      "access_expired") → durée d'accès dépassée
        (False, None,      message)          → identifiants incorrects
    """
    username = username.strip()
    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            expected = _hash_password(password, u["salt"])
            if expected != u["password_hash"]:
                return False, None, "Mot de passe incorrect."

            if is_suspended(u):
                return False, None, "suspended"
            if is_access_expired(u):
                return False, None, "access_expired"

            user_info = {
                "username":            u["username"],
                "role":                u["role"],
                "created_at":          u.get("created_at", ""),
                "password_changed_at": u.get("password_changed_at", ""),
            }
            from utils.activity_log import log_activity
            if is_password_expired(u):
                log_activity("login", "Connexion (mot de passe expiré)",
                             username=user_info["username"], role=user_info["role"])
                return True, user_info, "expired"
            log_activity("login", "Connexion réussie",
                         username=user_info["username"], role=user_info["role"])
            return True, user_info, ""
    from utils.activity_log import log_activity
    log_activity("login_failed", f"Tentative échouée pour : {username}",
                 username=username, role="")
    return False, None, "Nom d'utilisateur introuvable."


# ── Session courante ──────────────────────────────────────────────────────────

_current_user: dict | None = None


def set_current_user(user: dict | None):
    global _current_user
    _current_user = user


def get_current_user() -> dict | None:
    return _current_user


def current_role() -> str:
    u = _current_user
    return u["role"] if u else ""


def is_admin() -> bool:
    return current_role() == "admin"


def logout():
    u = _current_user
    if u:
        from utils.activity_log import log_activity
        log_activity("logout", "Déconnexion",
                     username=u.get("username", ""), role=u.get("role", ""))
    set_current_user(None)

"""
Authentication handler — gestion des comptes, mots de passe et sessions.

Stockage : users.json dans le répertoire racine du projet.
Sécurité  : hash SHA-256 + sel individuel par utilisateur.
Expiration: mot de passe expire après 90 jours (3 mois).
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
    "admin": "Administrateur — accès complet",
    "agent": "Agent — accès clients, cotations, factures",
}

# ── I/O ───────────────────────────────────────────────────────────────────────

def _load_users() -> list:
    """Charge la liste des utilisateurs depuis users.json."""
    if not os.path.exists(USERS_FILE):
        return []
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_users(users: list) -> bool:
    """Sauvegarde la liste des utilisateurs dans users.json."""
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ── Cryptographie ─────────────────────────────────────────────────────────────

def _hash_password(password: str, salt: str) -> str:
    """Retourne le hash SHA-256 de password+salt."""
    combined = (password + salt).encode("utf-8")
    return hashlib.sha256(combined).hexdigest()


def _generate_salt() -> str:
    """Génère un sel aléatoire de 32 caractères hex."""
    return secrets.token_hex(16)


# ── Utilisateurs ──────────────────────────────────────────────────────────────

def has_users() -> bool:
    """Retourne True si au moins un compte existe."""
    return len(_load_users()) > 0


def get_users() -> list:
    """Retourne la liste des utilisateurs (sans les mots de passe)."""
    users = _load_users()
    return [
        {
            "username": u["username"],
            "role": u["role"],
            "created_at": u.get("created_at", ""),
            "password_changed_at": u.get("password_changed_at", ""),
            "is_expired": is_password_expired(u),
        }
        for u in users
    ]


def create_user(username: str, password: str, role: str) -> tuple[bool, str]:
    """
    Crée un nouveau compte utilisateur.

    Returns:
        (True, "") en cas de succès
        (False, message_erreur) en cas d'échec
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
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    users.append(
        {
            "username": username,
            "password_hash": _hash_password(password, salt),
            "salt": salt,
            "role": role,
            "created_at": now,
            "password_changed_at": now,
        }
    )
    if _save_users(users):
        return True, ""
    return False, "Erreur lors de la sauvegarde du fichier utilisateurs."


def delete_user(username: str) -> tuple[bool, str]:
    """Supprime un compte utilisateur."""
    users = _load_users()
    new_users = [u for u in users if u["username"].lower() != username.lower()]
    if len(new_users) == len(users):
        return False, f"Utilisateur « {username} » introuvable."
    if _save_users(new_users):
        return True, ""
    return False, "Erreur lors de la sauvegarde."


def change_password(username: str, new_password: str) -> tuple[bool, str]:
    """Change le mot de passe d'un utilisateur et remet le compteur d'expiration à zéro."""
    if len(new_password) < 6:
        return False, "Le mot de passe doit avoir au moins 6 caractères."

    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            salt = _generate_salt()
            u["password_hash"] = _hash_password(new_password, salt)
            u["salt"] = salt
            u["password_changed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if _save_users(users):
                return True, ""
            return False, "Erreur lors de la sauvegarde."
    return False, f"Utilisateur « {username} » introuvable."


# ── Authentification ──────────────────────────────────────────────────────────

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


def authenticate(username: str, password: str) -> tuple[bool, dict | None, str]:
    """
    Vérifie les identifiants.

    Returns:
        (True,  user_dict, "")          → succès
        (True,  user_dict, "expired")   → succès mais mot de passe expiré
        (False, None,      message)     → échec
    """
    username = username.strip()
    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            expected = _hash_password(password, u["salt"])
            if expected != u["password_hash"]:
                return False, None, "Mot de passe incorrect."
            user_info = {
                "username": u["username"],
                "role": u["role"],
                "created_at": u.get("created_at", ""),
                "password_changed_at": u.get("password_changed_at", ""),
            }
            if is_password_expired(u):
                return True, user_info, "expired"
            return True, user_info, ""
    return False, None, "Nom d'utilisateur introuvable."


# ── Session courante ──────────────────────────────────────────────────────────

_current_user: dict | None = None


def set_current_user(user: dict | None):
    global _current_user
    _current_user = user


def get_current_user() -> dict | None:
    return _current_user


def current_role() -> str:
    """Retourne le rôle de l'utilisateur connecté, ou '' si non connecté."""
    u = _current_user
    return u["role"] if u else ""


def is_admin() -> bool:
    return current_role() == "admin"


def logout():
    set_current_user(None)

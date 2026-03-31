"""
Authentication handler — gestion des comptes, mots de passe et sessions.

Stockage : users.json dans le répertoire racine du projet.
Sécurité  : PBKDF2-HMAC-SHA256 (260 000 itérations) + sel individuel par utilisateur.
            Migration automatique de l'ancien hash SHA-256 lors de la connexion.
Expiration: mot de passe expire après 90 jours (3 mois).
Accès     : suspension manuelle + expiration de durée d'accès.
Rate limit: verrouillage temporaire après 5 échecs en 10 minutes (15 min de délai).
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

# Rate limiting : verrouillage après N échecs en WINDOW minutes
_LOCKOUT_MAX_FAILURES = 5
_LOCKOUT_WINDOW_MIN   = 10
_LOCKOUT_DURATION_MIN = 15
# {username_lower: [(timestamp, ...), ...]}
_failed_attempts: dict = {}

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
            return _valid_user_entries(json.load(f))
    except Exception:
        return []


def _save_users(users: list) -> bool:
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def _valid_user_entries(users: list) -> list[dict]:
    valid = []
    for user in users:
        if not isinstance(user, dict):
            continue
        username = str(user.get("username", "")).strip()
        role = str(user.get("role", "")).strip()
        if not username or not role:
            continue
        normalized = dict(user)
        normalized["username"] = username
        normalized["role"] = role
        valid.append(normalized)
    return valid


# ── Rate limiting ─────────────────────────────────────────────────────────────

def _record_failed_attempt(username: str) -> None:
    key = username.lower()
    now = datetime.now()
    attempts = _failed_attempts.get(key, [])
    attempts.append(now)
    # Conserver uniquement les tentatives dans la fenêtre de temps
    cutoff = now - timedelta(minutes=_LOCKOUT_WINDOW_MIN)
    _failed_attempts[key] = [t for t in attempts if t > cutoff]


def _clear_failed_attempts(username: str) -> None:
    _failed_attempts.pop(username.lower(), None)


def check_lockout(username: str) -> tuple[bool, int]:
    """
    Vérifie si un compte est verrouillé.

    Returns:
        (is_locked: bool, seconds_remaining: int)
    """
    key = username.lower()
    now = datetime.now()
    cutoff = now - timedelta(minutes=_LOCKOUT_WINDOW_MIN)
    recent = [t for t in _failed_attempts.get(key, []) if t > cutoff]

    if len(recent) >= _LOCKOUT_MAX_FAILURES:
        # Verrouillage à partir de la Nième tentative
        lock_start = sorted(recent)[_LOCKOUT_MAX_FAILURES - 1]
        lock_end = lock_start + timedelta(minutes=_LOCKOUT_DURATION_MIN)
        if now < lock_end:
            remaining = int((lock_end - now).total_seconds())
            return True, remaining

    return False, 0


# ── Cryptographie ─────────────────────────────────────────────────────────────

# Version 1 : SHA-256 simple (legacy, lecture seule)
def _hash_password_v1(password: str, salt: str) -> str:
    combined = (password + salt).encode("utf-8")
    return hashlib.sha256(combined).hexdigest()


# Version 2 : PBKDF2-HMAC-SHA256 (260 000 itérations)
def _hash_password_v2(password: str, salt: str) -> str:
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        260_000,
    )
    return dk.hex()


def _hash_password(password: str, salt: str) -> str:
    """Hache un mot de passe avec l'algorithme courant (v2 = PBKDF2)."""
    return _hash_password_v2(password, salt)


def _generate_salt() -> str:
    return secrets.token_hex(32)  # 32 octets = 256 bits


# ── Statut d'accès ────────────────────────────────────────────────────────────

def is_password_expired(user: dict) -> bool:
    """Retourne True si le mot de passe a expiré (>90 jours).
    En cas de timestamp invalide/corrompu, considère le mot de passe comme expiré
    (fail-secure : force le changement plutôt que d'autoriser l'accès)."""
    try:
        changed = user.get("password_changed_at") or user.get("created_at", "")
        if not changed:
            return True
        dt = datetime.strptime(changed, "%Y-%m-%d %H:%M:%S")
        return (datetime.now() - dt).days >= PASSWORD_EXPIRY_DAYS
    except Exception:
        return True  # fail-secure : timestamp corrompu → forcer le changement


def password_expires_at(user: dict) -> str:
    """Retourne la date/heure d'expiration du mot de passe."""
    try:
        changed = user.get("password_changed_at") or user.get("created_at", "")
        if not changed:
            return ""
        dt = datetime.strptime(changed, "%Y-%m-%d %H:%M:%S") + timedelta(days=PASSWORD_EXPIRY_DAYS)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ""


def password_days_left(user: dict) -> int | None:
    """Retourne le nombre de jours restants avant expiration du mot de passe."""
    try:
        expires_at = password_expires_at(user)
        if not expires_at:
            return None
        dt = datetime.strptime(expires_at, "%Y-%m-%d %H:%M:%S")
        delta = dt - datetime.now()
        return delta.days
    except Exception:
        return None


def is_password_expiring_soon(user: dict, warning_days: int = 7) -> bool:
    """Retourne True si le mot de passe expire bientôt."""
    days_left = password_days_left(user)
    return days_left is not None and 0 <= days_left < warning_days


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
    return len(_valid_user_entries(_load_users())) > 0


def _admin_count(users: list) -> int:
    return sum(1 for u in users if u.get("role") == "admin")


def _current_username() -> str:
    global _current_user
    return (_current_user or {}).get("username", "")


def get_users() -> list:
    users = _valid_user_entries(_load_users())
    return [
        {
            "username":            u["username"],
            "role":                u["role"],
            "created_at":          u.get("created_at", ""),
            "password_changed_at": u.get("password_changed_at", ""),
            "suspended":           u.get("suspended", False),
            "access_expires_at":   u.get("access_expires_at", ""),
            "is_expired":          is_password_expired(u),
            "password_expires_at": password_expires_at(u),
            "password_days_left":  password_days_left(u),
            "password_expiring_soon": is_password_expiring_soon(u),
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
        "hash_version":        2,
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
    current_username = _current_username()
    if current_username and current_username.lower() == username.lower():
        return False, "Vous ne pouvez pas supprimer votre propre compte."

    target = next((u for u in users if u["username"].lower() == username.lower()), None)
    if not target:
        return False, f"Utilisateur « {username} » introuvable."
    if target.get("role") == "admin" and _admin_count(users) <= 1:
        return False, "Impossible de supprimer le dernier administrateur."

    new_users = [u for u in users if u["username"].lower() != username.lower()]
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
            u["hash_version"]        = 2
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
    current_username = _current_username()
    if current_username and current_username.lower() == username.lower():
        return False, "Vous ne pouvez pas suspendre votre propre compte."
    for u in users:
        if u["username"].lower() == username.lower():
            if u.get("role") == "admin" and _admin_count(users) <= 1:
                return False, "Impossible de suspendre le dernier administrateur."
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


def update_user_role(username: str, new_role: str) -> tuple[bool, str]:
    """Modifie le rôle d'un utilisateur existant."""
    if new_role not in ROLES:
        return False, f"Rôle invalide. Valeurs acceptées : {', '.join(ROLES)}."

    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            old_role = u.get("role", "")
            if old_role == "admin" and new_role != "admin" and _admin_count(users) <= 1:
                return False, "Impossible de rétrograder le dernier administrateur."
            if old_role == new_role:
                return True, ""
            u["role"] = new_role
            if _save_users(users):
                from utils.activity_log import log_activity
                log_activity("change_user_role", f"Rôle modifié : {username} ({old_role} → {new_role})")
                return True, ""
            return False, "Erreur lors de la sauvegarde."
    return False, f"Utilisateur « {username} » introuvable."


def duplicate_user(source_username: str, new_username: str, password: str) -> tuple[bool, str]:
    """Duplique un compte en reprenant rôle et date d'accès, avec un nouveau mot de passe."""
    users = _load_users()
    source = next((u for u in users if u["username"].lower() == source_username.lower()), None)
    if not source:
        return False, f"Utilisateur source « {source_username} » introuvable."

    success, err = create_user(
        new_username,
        password,
        source.get("role", "agent"),
        access_expires_at=source.get("access_expires_at", ""),
    )
    if not success:
        return False, err

    from utils.activity_log import log_activity
    log_activity("duplicate_user", f"Compte dupliqué : {source_username} → {new_username}")
    return True, ""


# ── Authentification ──────────────────────────────────────────────────────────

def authenticate(username: str, password: str) -> tuple[bool, dict | None, str]:
    """
    Vérifie les identifiants.

    Returns:
        (True,  user_dict, "")               → succès
        (True,  user_dict, "expired")        → succès mais mot de passe expiré
        (False, None,      "suspended")      → compte suspendu
        (False, None,      "access_expired") → durée d'accès dépassée
        (False, None,      "locked")         → verrouillage temporaire (brute force)
        (False, None,      message)          → identifiants incorrects
    """
    username = username.strip()

    # ── Vérification du verrouillage ──────────────────────────────────────────
    is_locked, secs = check_lockout(username)
    if is_locked:
        mins = (secs + 59) // 60
        from utils.activity_log import log_activity
        log_activity("login_blocked", f"Tentative bloquée (verrouillage) pour : {username}",
                     username=username, role="")
        return False, None, f"locked:{mins}"

    users = _load_users()
    for u in users:
        if u["username"].lower() == username.lower():
            # ── Vérification du hash (v1 SHA-256 ou v2 PBKDF2) ───────────────
            hash_version = u.get("hash_version", 1)
            if hash_version == 1:
                expected = _hash_password_v1(password, u["salt"])
            else:
                expected = _hash_password_v2(password, u["salt"])

            if expected != u["password_hash"]:
                _record_failed_attempt(username)
                from utils.activity_log import log_activity
                log_activity("login_failed", f"Tentative échouée pour : {username}",
                             username=username, role="")
                return False, None, "Mot de passe incorrect."

            # ── Succès : effacer le compteur d'échecs ─────────────────────────
            _clear_failed_attempts(username)

            # ── Migration automatique SHA-256 → PBKDF2 ────────────────────────
            if hash_version == 1:
                new_salt = _generate_salt()
                u["salt"]         = new_salt
                u["password_hash"] = _hash_password_v2(password, new_salt)
                u["hash_version"]  = 2
                _save_users(users)

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

    _record_failed_attempt(username)
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

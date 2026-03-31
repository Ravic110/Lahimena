"""
Historique d'activité utilisateur — Lahimena Tours.

Stockage : activity_log.json (données courantes) + activity_log_archive/ (archives).
Rotation automatique : entrées > 30 jours archivées dans un fichier mensuel.
Chaque entrée : timestamp, username, role, action, label, details, category
"""

import json
import os
from datetime import datetime, timedelta
from collections import Counter

_BASE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ACTIVITY_FILE = os.path.join(_BASE_DIR, "activity_log.json")
ARCHIVE_DIR   = os.path.join(_BASE_DIR, "activity_log_archive")

# ── Catégories et couleurs ────────────────────────────────────────────────────
# category → (label_couleur_hex, icône)
CATEGORY_STYLE = {
    "auth":    ("#1565C0", "🔐"),   # bleu
    "client":  ("#1B7A3E", "👥"),   # vert
    "user":    ("#C62828", "👤"),   # rouge
    "quota":   ("#6A1B9A", "📄"),   # violet
    "invoice": ("#E65100", "💶"),   # orange
    "nav":     ("#607D8B", "🧭"),   # gris-bleu
}

ACTION_META = {
    # action: (label, category)
    "login":              ("Connexion",                        "auth"),
    "login_failed":       ("Tentative de connexion échouée",  "auth"),
    "logout":             ("Déconnexion",                      "auth"),
    "create_client":      ("Création d'un client",             "client"),
    "edit_client":        ("Modification d'un client",         "client"),
    "delete_client":      ("Suppression d'un client",          "client"),
    "create_user":        ("Création d'un compte",             "user"),
    "delete_user":        ("Suppression d'un compte",          "user"),
    "suspend_user":       ("Suspension d'un compte",           "user"),
    "reactivate_user":    ("Réactivation d'un compte",         "user"),
    "change_password":    ("Changement de mot de passe",       "user"),
    "change_user_role":   ("Modification du rôle",             "user"),
    "duplicate_user":     ("Duplication d'un compte",          "user"),
    "set_expiry":         ("Modification date d'expiration",   "user"),
    "create_quotation":   ("Création d'une cotation",          "quota"),
    "edit_quotation":     ("Modification d'une cotation",      "quota"),
    "delete_quotation":   ("Suppression d'une cotation",       "quota"),
    "create_invoice":     ("Création d'une facture",           "invoice"),
    "edit_invoice":       ("Modification d'une facture",       "invoice"),
    "delete_invoice":     ("Suppression d'une facture",        "invoice"),
    "navigate":           ("Navigation",                       "nav"),
}

# Seuil de détection brute force
BRUTE_FORCE_THRESHOLD = 5   # échecs dans la fenêtre
BRUTE_FORCE_WINDOW_MIN = 10  # minutes


# ── I/O ───────────────────────────────────────────────────────────────────────

def _load() -> list:
    try:
        with open(ACTIVITY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def _iter_entries(entries: list):
    for entry in entries:
        if isinstance(entry, dict):
            yield entry


def _save(entries: list) -> None:
    with open(ACTIVITY_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


# ── Rotation des logs ─────────────────────────────────────────────────────────

def _rotate_old_entries(entries: list) -> list:
    """
    Archive les entrées de plus de 30 jours dans activity_log_archive/YYYY-MM.json.
    Retourne uniquement les entrées récentes.
    """
    cutoff = datetime.now() - timedelta(days=30)
    recent, old = [], []
    for e in _iter_entries(entries):
        try:
            ts = datetime.strptime(e["timestamp"], "%Y-%m-%d %H:%M:%S")
            (old if ts < cutoff else recent).append(e)
        except (KeyError, ValueError):
            recent.append(e)

    if not old:
        return recent

    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    # Regroupe par mois
    by_month: dict = {}
    for e in old:
        month = e["timestamp"][:7]  # "YYYY-MM"
        by_month.setdefault(month, []).append(e)

    for month, month_entries in by_month.items():
        archive_file = os.path.join(ARCHIVE_DIR, f"{month}.json")
        existing = []
        if os.path.exists(archive_file):
            try:
                with open(archive_file, "r", encoding="utf-8") as f:
                    existing = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        existing.extend(month_entries)
        with open(archive_file, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)

    return recent


# ── Détection brute force ─────────────────────────────────────────────────────

def _check_brute_force(username: str, entries: list) -> bool:
    """
    Retourne True si trop de login_failed récents pour cet utilisateur.
    """
    window_start = datetime.now() - timedelta(minutes=BRUTE_FORCE_WINDOW_MIN)
    count = 0
    for e in _iter_entries(entries):
        if e.get("action") != "login_failed":
            continue
        if e.get("username", "").lower() != username.lower():
            continue
        try:
            ts = datetime.strptime(e["timestamp"], "%Y-%m-%d %H:%M:%S")
            if ts >= window_start:
                count += 1
        except (KeyError, ValueError):
            pass
    return count >= BRUTE_FORCE_THRESHOLD


# ── API publique ──────────────────────────────────────────────────────────────

def log_activity(action: str, details: str = "", username: str = "",
                 role: str = "") -> None:
    """
    Enregistre une action dans l'historique.
    Récupère username/role depuis la session courante si non fournis.
    Effectue la rotation des vieilles entrées (> 30 jours).
    Détecte les tentatives de brute force.
    """
    if not username:
        try:
            from utils.auth_handler import get_current_user
            u = get_current_user()
            if u:
                username = u.get("username", "")
                role     = u.get("role", "")
        except Exception:
            pass

    label, category = ACTION_META.get(action, (action, "nav"))

    entry = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "username":  username,
        "role":      role,
        "action":    action,
        "label":     label,
        "category":  category,
        "details":   details,
    }

    entries = list(_iter_entries(_load()))

    # Alerte brute force
    if action == "login_failed" and username:
        entries.append(entry)
        if _check_brute_force(username, entries):
            alert = {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "username":  "SYSTÈME",
                "role":      "system",
                "action":    "brute_force_alert",
                "label":     "⚠ Alerte sécurité — Trop de tentatives",
                "category":  "auth",
                "details":   f"{BRUTE_FORCE_THRESHOLD} échecs en {BRUTE_FORCE_WINDOW_MIN} min pour : {username}",
            }
            entries.append(alert)
    else:
        entries.append(entry)

    # Rotation des entrées anciennes
    entries = _rotate_old_entries(entries)
    _save(entries)


def get_activity(username: str = "", limit: int = 500,
                 action_filter: str = "",
                 date_from: str = "", date_to: str = "",
                 search: str = "") -> list:
    """
    Retourne les entrées d'activité (ordre antéchronologique).

    Paramètres :
      username      : filtre sur un utilisateur (vide = tous)
      limit         : nombre max d'entrées retournées
      action_filter : filtre sur la catégorie ou l'action exacte
      date_from     : date début "YYYY-MM-DD"
      date_to       : date fin "YYYY-MM-DD"
      search        : recherche texte libre sur tous les champs
    """
    entries = list(_iter_entries(_load()))

    if username:
        entries = [e for e in entries
                   if e.get("username", "").lower() == username.lower()]

    if action_filter:
        entries = [e for e in entries
                   if e.get("category") == action_filter
                   or e.get("action") == action_filter]

    if date_from:
        entries = [e for e in entries if e.get("timestamp", "") >= date_from]

    if date_to:
        dt_end = date_to + " 23:59:59"
        entries = [e for e in entries if e.get("timestamp", "") <= dt_end]

    if search:
        q = search.lower()
        entries = [e for e in entries if any(
            q in str(v).lower() for v in e.values()
        )]

    return list(reversed(entries))[:limit]


def get_all_activity(limit: int = 500, **kwargs) -> list:
    """Retourne toutes les activités. Accepte les mêmes filtres que get_activity."""
    return get_activity(username="", limit=limit, **kwargs)


def get_user_stats(username: str) -> dict:
    """
    Retourne des statistiques d'activité pour un utilisateur :
      - total_actions
      - login_count (30 derniers jours)
      - failed_logins (30 derniers jours)
      - last_login
      - last_action_ts
      - top_actions : [(label, count), ...]
      - daily_counts : {date_str: count}  (30 derniers jours)
      - brute_force_detected : bool
    """
    entries = list(_iter_entries(_load()))
    user_entries = [e for e in entries
                    if e.get("username", "").lower() == username.lower()]

    thirty_days_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")

    login_count    = 0
    failed_logins  = 0
    last_login     = ""
    last_action_ts = ""
    action_counts: Counter = Counter()
    daily_counts: dict = {}

    for e in user_entries:
        ts  = e.get("timestamp", "")
        act = e.get("action", "")
        lbl = e.get("label", act)
        day = ts[:10]

        if ts > last_action_ts:
            last_action_ts = ts

        if act == "login":
            login_count += 1
            if ts > last_login:
                last_login = ts
        elif act == "login_failed":
            failed_logins += 1

        if day >= thirty_days_ago:
            action_counts[lbl] += 1
            daily_counts[day] = daily_counts.get(day, 0) + 1

    brute = _check_brute_force(username, entries)

    return {
        "total_actions":      len(user_entries),
        "login_count":        login_count,
        "failed_logins":      failed_logins,
        "last_login":         last_login,
        "last_action_ts":     last_action_ts,
        "top_actions":        action_counts.most_common(5),
        "daily_counts":       daily_counts,
        "brute_force_detected": brute,
    }


def get_recent_activity_summary(username: str, limit: int = 5) -> list:
    """Retourne les actions récentes d'un utilisateur pour un affichage résumé."""
    return get_activity(username=username, limit=limit)


def get_brute_force_usernames() -> list[str]:
    """Retourne les utilisateurs avec trop d'échecs de connexion récents."""
    entries = list(_iter_entries(_load()))
    failed_users = {
        e.get("username", "").strip()
        for e in entries
        if e.get("action") == "login_failed" and e.get("username", "").strip()
    }
    return sorted(
        username
        for username in failed_users
        if _check_brute_force(username, entries)
    )


def get_archive_months() -> list[str]:
    """Retourne la liste des mois archivés (ex: ['2026-01', '2026-02'])."""
    if not os.path.isdir(ARCHIVE_DIR):
        return []
    return sorted([
        f.replace(".json", "")
        for f in os.listdir(ARCHIVE_DIR)
        if f.endswith(".json")
    ])


def load_archive(month: str) -> list:
    """Charge les entrées archivées d'un mois donné (format YYYY-MM)."""
    archive_file = os.path.join(ARCHIVE_DIR, f"{month}.json")
    try:
        with open(archive_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []

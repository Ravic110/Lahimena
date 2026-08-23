"""
Pytest configuration and fixtures for Lahimena Tours tests
"""

import os
import sys
from pathlib import Path

import pytest

# Racine du projet sur sys.path, pour que `import config` fonctionne.
# `.parent` et non `.parent.parent` : ce fichier est a la racine, pas dans
# tests/. L'ancienne version ajoutait le dossier PARENT du projet ; seule
# l'insertion automatique de la rootdir par pytest masquait l'erreur.
PROJECT_ROOT = Path(__file__).parent.resolve()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


@pytest.fixture
def client_data_dict():
    """Fixture providing sample client data"""
    return {
        "ref_client": "CLI-2026-001",
        "nom": "Test Client",
        "telephone": "0301234567",
        "email": "test@example.com",
        "code_pays": "MG",
        "periode": "5 jours",
        "restauration": "Petit-déjeuner",
        "hebergement": "Hôtel 3*",
        "chambre": "Double",
        "enfant": "Non",
        "age_enfant": "",
        "forfait": "Standard",
        "circuit": "Côte Est",
    }


@pytest.fixture
def hotel_data_dict():
    """Fixture providing sample hotel data"""
    return {
        "nom": "Test Hotel",
        "adresse": "123 Test Street",
        "email": "hotel@test.mg",
        "telephone": "0302345678",
        "code_pays": "MG",
        "etoiles": "3",
        "region": "Centre",
    }


@pytest.fixture
def temp_excel_file(tmp_path):
    """Fixture providing a temporary Excel file path"""
    return str(tmp_path / "test_data.xlsx")


@pytest.fixture
def mock_logger(monkeypatch):
    """Fixture providing a mock logger"""
    from unittest.mock import MagicMock

    mock = MagicMock()

    # Mock the logger functions
    mock.info = MagicMock()
    mock.warning = MagicMock()
    mock.error = MagicMock()
    mock.debug = MagicMock()

    return mock


@pytest.fixture(scope="session")
def project_root():
    """Fixture providing the project root path"""
    return PROJECT_ROOT


# ── Harnais des tests d'interface ─────────────────────────────────────────────
#
# `gui/` est a 0 % de couverture : les cinq ecrans de cotation, soit 4 500
# lignes, n'ont aucun filet. Les refactorer a l'aveugle serait imprudent, d'ou
# ces fixtures. Elles s'effacent d'elles-memes la ou aucun affichage n'est
# disponible, pour que la suite reste executable partout.


def _affichage_disponible():
    """Vrai si un serveur graphique repond."""
    import tkinter as tk

    try:
        racine = tk.Tk()
    except Exception:
        return False
    racine.destroy()
    return True


@pytest.fixture(scope="session")
def _tk_disponible():
    return _affichage_disponible()


@pytest.fixture(scope="session")
def _tk_racine(_tk_disponible):
    """Une seule fenetre racine pour toute la session.

    Creer et detruire un interpreteur Tk par test coutait ~1,2 s piece. La
    racine est partagee ; l'isolation entre tests vient du cadre jetable
    fourni par `tk_root`.
    """
    if not _tk_disponible:
        yield None
        return

    import tkinter as tk

    racine = tk.Tk()
    racine.withdraw()
    try:
        yield racine
    finally:
        try:
            racine.destroy()
        except Exception:
            pass


@pytest.fixture
def tk_root(_tk_racine):
    """Cadre parent jetable, detruit avec toute sa descendance en fin de test.

    Ignore le test plutot que d'echouer quand il n'y a pas d'affichage
    (conteneur CI sans Xvfb, session SSH sans X11).
    """
    if _tk_racine is None:
        pytest.skip("aucun affichage disponible pour les tests d'interface")

    import tkinter as tk

    cadre = tk.Frame(_tk_racine)
    cadre.pack(fill="both", expand=True)
    try:
        yield cadre
    finally:
        try:
            cadre.destroy()
        except Exception:
            pass


@pytest.fixture
def client_type():
    """Client complet, tel qu'un ecran de cotation le recoit."""
    return {
        "ref_client": "CLI-2026-042",
        "numero_dossier": "D-2026-042",
        "nom": "RAKOTO",
        "prenom": "Jean",
        "nombre_participants": "4",
        "nombre_adultes": "2",
        "enfants_2_12": "2",
        "duree_sejour": "5",
        "ville_depart": "Antananarivo",
        "ville_arrivee": "Antsirabe",
        "circuit": "Antananarivo > Antsirabe > Fianarantsoa",
        "periode": "Haute saison",
        "restauration": "Demi-pension",
        "hebergement": "Hôtel 3*",
        "compagnie": "Tsaradia",
    }


def pytest_configure(config):
    """Configure pytest with custom settings"""
    config.addinivalue_line("markers", "gui: tests instanciant des widgets Tk")

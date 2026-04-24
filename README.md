# Lahimena Tours — Application de gestion touristique

![Python](https://img.shields.io/badge/python-3.10+-blue)
![License](https://img.shields.io/badge/license-proprietary-red)
![Status](https://img.shields.io/badge/status-production-brightgreen)

Application desktop de gestion complète pour agence de voyage, développée avec **Python** et **CustomTkinter** pour une interface moderne et ergonomique.

## 🎯 Objectif

Fournir une solution tout-en-un pour les agences de voyage en Afrique de l'Ouest, avec :
- Gestion intégrée des clients et voyages
- Cotations automatisées (hôtels, transports, visites)
- Facturation et PDF
- Comptabilité via module TsaraKonta
- Contrôle d'accès multi-rôles avec audit d'activité

---

## ✨ Fonctionnalités principales

### Gestion clients & voyages
- Formulaire de demande client (type, séjour, participants, hébergement, restauration)
- Itinéraires avec sous-onglets : Circuit · Voiture · Guide national
- Rooming (SGL / DBL / TWN / TPL / FML)
- Liste clients avec recherche, tri, filtres et aperçu rapide

### Cotations & factures
- Cotation hôtelière avec calcul automatique
- Cotation transport, billets d'avion, visites/excursions, charges collectives
- Historique des devis et factures
- Génération PDF automatique

### Comptabilité (module TsaraKonta)
- État financier intégré (Bilan, Compte de résultat, Flux de trésorerie)
- Tableaux d'amortissement, immobilisations
- Prévisions 5 ans et 12 mois
- Journal comptable

### Sécurité & gestion des accès
- Authentification par login/mot de passe
- Trois rôles : `admin`, `agent`, `comptable`
- Hachage PBKDF2-HMAC-SHA256 (260 000 itérations) avec sel individuel
- Migration automatique des anciens hachages SHA-256 à la connexion
- Verrouillage automatique après 5 échecs en 10 minutes (15 min de délai)
- Expiration des mots de passe (90 jours) avec changement forcé
- Expiration configurable de l'accès par compte

### Gestion des utilisateurs (admin)
- Panneau latéral de détail par utilisateur
- Compteurs : total / actifs / suspendus / expirés
- Avatars avec initiales colorées
- Badges de statut visuels
- Recherche, filtres par rôle et statut, tri par colonne
- Modification du rôle, prolongation d'accès (+30j / +90j / +1 an / illimité)
- Duplication de compte
- Alerte brute force en temps réel

### Historique d'activité
- Enregistrement de toutes les actions (connexions, clients, cotations, navigation)
- Filtres par utilisateur, catégorie, date, recherche libre
- Statistiques par utilisateur (connexions, échecs, top actions, graphique 30j)
- Export CSV (UTF-8 BOM)
- Rotation automatique des logs > 30 jours vers archives mensuelles

---

## Structure du projet

```
Lahimena/
├── main.py                          # Point d'entrée — patch CTk + lancement
├── config.py                        # Constantes, thèmes, chemins
├── users.json                       # Comptes utilisateurs (hachés)
├── activity_log.json                # Journal d'activité courant
├── activity_log_archive/            # Archives mensuelles des logs
│
├── gui/
│   ├── sidebar.py                   # Barre latérale de navigation
│   ├── main_content.py              # Topbar + zone de contenu principale
│   ├── ui_style.py                  # Composants UI réutilisables
│   └── forms/
│       ├── login_form.py            # Écran de connexion
│       ├── home_page.py             # Tableau de bord accueil
│       ├── client_form.py           # Formulaire demande client
│       ├── client_list.py           # Liste des clients
│       ├── client_page.py           # Hub page clients
│       ├── account_management.py    # Gestion des comptes (admin)
│       ├── client_quotation.py      # Cotation client + PDF
│       ├── client_quotation_history.py
│       ├── client_quotation_page.py
│       ├── hotel_quotation.py       # Cotation hôtelière
│       ├── hotel_quotation_page.py
│       ├── hotel_quotation_summary.py
│       ├── hotel_form.py / hotel_list.py
│       ├── transport_*.py           # Cotation transport
│       ├── air_ticket_*.py          # Billets d'avion
│       ├── visite_excursion_*.py    # Visites & excursions
│       ├── collective_expense_*.py  # Charges collectives
│       ├── invoice_management.py    # Gestion des factures
│       ├── billing_quotes_hub_page.py
│       ├── cotation_hub_page.py
│       ├── database_hub_page.py
│       ├── circuit_db_page.py
│       ├── transport_db_page.py
│       ├── parametrage_*.py         # Paramètres application
│       └── expenses_page.py
│
├── models/
│   ├── client_data.py               # Modèle de données client
│   └── hotel_data.py                # Modèle de données hôtel
│
├── utils/
│   ├── auth_handler.py              # Authentification, sessions, rôles
│   ├── activity_log.py              # Journal d'activité utilisateur
│   ├── excel_handler.py             # Lecture/écriture Excel (openpyxl)
│   ├── cache.py                     # Cache TTL pour données fréquentes
│   ├── validators.py                # Validation email, téléphone, dates
│   └── logger.py                    # Logging applicatif centralisé
│
├── finances/
│   └── tsarakonta/                  # Module financier autonome
│       ├── main.py                  # Lancement TsaraKonta
│       └── ...
│
├── assets/
│   └── logo.png
│
├── tests/                           # Suite de tests pytest
│   ├── test_config.py
│   ├── test_excel_handler.py
│   ├── test_models.py
│   ├── test_validators.py
│   ├── test_invoicing.py
│   └── test_financial_tool.py
│
└── audit/                           # Rapports d'audit et correctifs
```

---

## Installation

### Prérequis
- Python 3.10+
- tkinter (inclus dans Python standard)

### Installation des dépendances

```bash
# Dépendances principales
pip install -e .

# Avec le module financier TsaraKonta
pip install -e .[financial]
```

Ou avec les fichiers requirements :

```bash
pip install -r requirements.txt
pip install -r requirements-financial.txt   # optionnel
```

Dépendances principales :

| Paquet | Usage |
|--------|-------|
| `customtkinter` | Interface graphique moderne |
| `Pillow` | Avatars utilisateurs, logo |
| `openpyxl` | Lecture/écriture fichiers Excel |
| `phonenumbers` | Validation numéros de téléphone |

### Lancement

```bash
python main.py
```

---

## Sécurité

| Mécanisme | Détail |
|-----------|--------|
| Hachage MDP | PBKDF2-HMAC-SHA256, 260 000 itérations, sel 32 octets |
| Migration automatique | Les anciens hash SHA-256 migrent vers PBKDF2 à la prochaine connexion |
| Expiration MDP | 90 jours, changement forcé à la connexion suivante |
| Rate limiting | Verrouillage 15 min après 5 échecs en 10 min (en mémoire) |
| Détection brute force | Alerte dans le journal + badge rouge dans le gestionnaire d'utilisateurs |
| Contrôle d'accès | Navigation sidebar désactivée selon le rôle (`comptable` → accès restreint) |

---

## Module financier TsaraKonta

Module autonome dans `finances/tsarakonta/`. Lancé en sous-processus depuis l'application principale.

```bash
python finances/tsarakonta/main.py --excel ./journal.xlsx --etat Bilan
# ou
python -m finances.tsarakonta --excel journal.xlsx --etat Bilan
```

---

## Tests

```bash
pytest -q
# Résultat attendu : ~43 passed, 2 skipped
```

---

## Rôles utilisateurs

| Rôle | Accès |
|------|-------|
| `admin` | Accès complet + gestion des comptes utilisateurs |
| `agent` | Clients, cotations, factures |
| `comptable` | Module financier (TsaraKonta) uniquement |

---

## Configuration

### Variables d'environnement

Créez un fichier `.env` à la racine du projet :

```bash
# Paramètres de l'application
APP_DEBUG=False
APP_THEME=dark  # light ou dark

# Base de données (optionnel - sinon JSON)
# DB_TYPE=sqlite
# DB_PATH=./data/app.db

# Logs
LOG_LEVEL=INFO
LOG_DIR=./logs
```

### Fichiers de configuration clés

| Fichier | Description |
|---------|---|
| `config.py` | Constantes applicatives, thèmes, chemins |
| `users.json` | Comptes utilisateurs (ne pas commiter) |
| `activity_log.json` | Journal d'activité courant (ne pas commiter) |
| `pyproject.toml` | Configuration setuptools et dépendances |

---

## Démarrage rapide

### Installation en développement

```bash
# 1. Cloner le repo et naviguer
cd Lahimena

# 2. Créer un environnement virtuel
python -m venv .env
source .env/bin/activate  # Linux/Mac
# ou
.env\Scripts\activate      # Windows

# 3. Installer les dépendances
pip install -e .

# 4. (Optionnel) Installer les dépendances financières
pip install -e .[financial]

# 5. Lancer l'application
python main.py
```

### Compte par défaut (première connexion)

Les comptes par défaut ne sont créés qu'au premier lancement :

```
Identifiant : admin
Mot de passe : DefaultPassword123!
```

**⚠️ Changez ce mot de passe immédiatement après la première connexion.**

---

## Development

### Structure de code

- **`gui/forms/`** : Écrans de l'interface utilisateur
- **`models/`** : Modèles de données (clients, hôtels)
- **`utils/`** : Services partagés (auth, excel, validation)
- **`finances/`** : Module TsaraKonta (autonome)
- **`tests/`** : Suite de tests pytest

### Exécuter les tests

```bash
# Tous les tests
pytest

# Avec verbose
pytest -v

# Test spécifique
pytest tests/test_validators.py -v

# Couverture de code
pytest --cov=. --cov-report=html
```

### Linting & formatage

```bash
# Vérifier les styles (optionnel)
flake8 . --max-line-length=100

# Formater avec black (optionnel)
black . --line-length=100
```

---

## Troubleshooting

### Problème : `ModuleNotFoundError: No module named 'customtkinter'`

**Solution :**
```bash
pip install --upgrade customtkinter
```

### Problème : Application ne démarre pas sous macOS

**Solution :**
```bash
python -m main
# ou via le Makefile
make run
```

### Problème : Fichiers Excel ne se chargent pas

**Vérifier :**
1. Le fichier est fermé (pas ouvert dans Excel)
2. Permissions d'accès OK sur le fichier
3. Format .xlsx valide (pas .xls ancien format)

### Problème : Tests échouent avec `ImportError`

**Solution :**
```bash
# Réinstaller les dépendances de test
pip install -r requirements-test.txt
pytest --cache-clear
```

---

## Contribution

### Avant de commencer

1. Créer une branche feature :
   ```bash
   git checkout -b feature/ma-fonctionnalite
   ```

2. Respecter le style Python (PEP 8)

3. Ajouter des tests pour les nouvelles fonctionnalités

4. Mettre à jour le README si nécessaire

### Workflow

```bash
# 1. Faire les modifications
# 2. Ajouter les tests
pytest tests/

# 3. Vérifier la couverture
pytest --cov=.

# 4. Commiter avec messages clairs
git add .
git commit -m "feat: ajouter nouvelle cotation"

# 5. Push et créer une Pull Request
git push origin feature/ma-fonctionnalite
```

---

## Licence

Ce projet est propriétaire. Tous droits réservés © 2026 Lahimena Tours.

---

## Support & Contact

- 📧 Email : [contact@lahimena.mg](mailto:contact@lahimena.mg)
- 🌐 Site : [www.lahimena.mg](https://www.lahimena.mg)
- 📞 Support : +261 XX XXX XXXX

---

## Historique des versions

### v2.0.0 (Actuel - 2026-04)
- Module financier TsaraKonta intégré
- Nouvel UI avec CustomTkinter
- Gestion d'accès par rôles
- Journal d'activité avec statistiques

### v1.0.0 (2025-Q4)
- Initialisation du projet
- Formulaire client & cotations de base
- Gestion des utilisateurs

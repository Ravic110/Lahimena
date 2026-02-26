# Lahimena Tours Devis Generation

Application de gestion de devis pour Lahimena Tours utilisant CustomTkinter.

## Structure du projet

```
ApiOK/
├── main.py                    # Point d'entrée principal
├── config.py                  # Constantes et configuration
├── gui/                       # Interface utilisateur
│   ├── __init__.py
│   ├── sidebar.py             # Barre latérale de navigation
│   ├── main_content.py        # Zone de contenu principale
│   └── forms/
│       ├── __init__.py
│       └── client_form.py     # Formulaire client
├── models/                    # Modèles de données
│   ├── __init__.py
│   └── client_data.py         # Modèle de données client
├── utils/                     # Utilitaires
│   ├── __init__.py
│   ├── excel_handler.py       # Gestion des fichiers Excel
│   └── validators.py          # Validations (email, téléphone)
├── assets/                    # Ressources statiques
│   └── logo.png
├── finances/                 # Outils financiers externes
│   └── tsarakonta/           # application d'état financier du projet
│       ├── main.py           # Lancement de TsaraKonta
│       └── ...               # code de l'outil
└── data.xlsx                  # Fichier de données Excel
```

## Fonctionnalités

- Formulaire de demande client avec validation
- Sauvegarde automatique des données dans Excel
- Interface moderne avec thème sombre
- Navigation par menu latéral extensible

## Dépendances

Le projet est maintenant empaqueté avec `pyproject.toml` et peut être installé avec `pip`.

### Installation des dépendances

```bash
pip install -e .                # installe les dépendances de base
pip install -e .[financial]     # ajoute les paquets nécessaires pour TsaraKonta
```

Les options équivalentes en utilisant les fichiers de requirements sont encore disponibles :

```bash
pip install -r requirements.txt
pip install -r requirements-financial.txt
```

### Prérequis
- Python 3.7+
- tkinter (inclus dans Python)

## Installation

1. **Cloner ou télécharger le projet**
2. **Installer Python 3.7+** si ce n'est pas déjà fait
3. **Installer les dépendances** :
   ```bash
   pip install -r requirements.txt
   ```
4. **Lancer l'application** :
   ```bash
   python main.py
   ```

### Dépendances détaillées

- **customtkinter** : Interface graphique moderne
- **Pillow** : Traitement d'images (logo)
- **openpyxl** : Manipulation Excel
- **phonenumbers** : Validation numéros téléphone

*Note : L'application fonctionne même sans certaines dépendances optionnelles (avec fonctionnalités réduites)*

### Module financier TsaraKonta

Le sous-répertoire `finances/tsarakonta` contient l'outil d'"état financier" du projet. Il faisait auparavant l'objet d'un dépôt Git séparé ; ce dernier a été supprimé et tout le code est désormais suivi dans le même dépôt principal. Il est toujours autonome ; on peut lancer la version embarquée

Les dépendances de cet outil sont listées dans `requirements-financial.txt` (pandas, openpyxl, etc.) :

```bash
pip install -r requirements-financial.txt
```


```bash
python finances/tsarakonta/main.py --excel ./chemin/vers/journal.xlsx --etat Bilan
```

ou comme un paquet Python :

```bash
python -m finances.tsarakonta --excel journal.xlsx --etat Bilan
```

Cette application utilise ses propres dépendances (pandas, etc.) et est séparée du cœur de Lahimena.

## Utilisation

1. L'application se lance avec l'écran d'accueil
2. Utiliser le menu latéral pour naviguer
3. Le formulaire client permet de saisir et valider les demandes
4. Les données sont automatiquement sauvegardées dans `data.xlsx`

## Développement

L'application est structurée de manière modulaire pour faciliter la maintenance et l'extension :

- **config.py** : Toutes les constantes et configurations
- **gui/** : Composants d'interface utilisateur
- **models/** : Modèles de données
- **utils/** : Fonctions utilitaires réutilisables

Chaque module peut être testé et modifié indépendamment.
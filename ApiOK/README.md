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
└── data.xlsx                  # Fichier de données Excel
```

## Fonctionnalités

- Formulaire de demande client avec validation
- Sauvegarde automatique des données dans Excel
- Interface moderne avec thème sombre
- Navigation par menu latéral extensible

## Dépendances

### Obligatoires
- Python 3.7+
- tkinter (inclus dans Python)
- customtkinter

### Optionnelles (recommandées)
- openpyxl : Pour la sauvegarde Excel
- phonenumbers : Pour la validation avancée des numéros de téléphone
- PIL (Pillow) : Pour le chargement du logo

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
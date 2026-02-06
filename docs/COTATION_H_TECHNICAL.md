# Résumé technique - Implémentation COTATION_H

**Date**: 6 février 2026  
**Fonctionnalité**: Feuille de regroupement des cotations hôtel (COTATION_H)

## Résumé des changements

### 🔧 Modifications apportées

#### 1. **config.py** - Configuration
```python
COTATION_H_SHEET_NAME = "COTATION_H"
```
Ajout du nom de la nouvelle feuille Excel.

#### 2. **utils/excel_handler.py** - Fonctions Excel

**Imports modifiés:**
```python
from config import CLIENT_EXCEL_PATH, HOTEL_EXCEL_PATH, CLIENT_SHEET_NAME, HOTEL_SHEET_NAME, COTATION_H_SHEET_NAME
```

**Nouvelles fonctions:**

##### `save_hotel_quotation_to_excel(quotation_data)`
- Enregistre une cotation hôtel dans la feuille COTATION_H
- Crée automatiquement la feuille si elle n'existe pas
- Ajoute les en-têtes au premier appel
- Retourne le numéro de ligne ou -1 en cas d'erreur
- Format automatique des colonnes

**Paramètres:**
```python
quotation_data = {
    'client_id': str,              # Référence client
    'client_name': str,            # Nom du client
    'hotel_name': str,             # Nom de l'hôtel
    'city': str,                   # Ville
    'total_price': float,          # Montant total
    'currency': str,               # Devise (Ariary, Euro, Dollar)
    'nights': int,                 # Nombre de nuits
    'adults': int,                 # Nombre d'adultes
    'children': int,               # Nombre d'enfants
    'room_type': str,              # Type de chambre
    'meal_plan': str,              # Plan de restauration
    'period': str,                 # Période/Saison
    'quote_date': str              # Date de la cotation
}
```

##### `load_all_hotel_quotations()`
- Charge toutes les cotations de la feuille COTATION_H
- Retourne une liste de dictionnaires avec toutes les données
- Gère les erreurs de parsing numérique avec `_parse_num()`

##### `get_quotations_grouped_by_client()`
- Regroupe les quotations par client
- Calcule le sous-total pour chaque client
- Retourne un dictionnaire structuré pour affichage

**Structure retournée:**
```python
{
    'client_id': {
        'client_name': str,
        'quotations': [list of quotation dicts],
        'total': float,
        'currency': str
    }
}
```

##### `get_quotations_by_city()`
- Regroupe les quotations par ville
- Calcule le sous-total pour chaque ville
- Utilise la même structure pour cohérence

#### 3. **gui/forms/hotel_quotation.py** - Intégration

**Import ajouté:**
```python
from utils.excel_handler import load_all_hotels, load_all_clients, save_hotel_quotation_to_excel
```

**Fonction modifiée: `_generate_quote()`**
- Après génération du PDF, enregistre les données
- Extraction automatique du client_id du format "REF - NAME"
- Gestion des erreurs avec logging
- Pas d'interruption si l'enregistrement échoue

**Code intégré:**
```python
quotation_data = {
    'client_id': self.client_var.get().split(' - ')[0],
    'client_name': client_name,
    'hotel_name': self.selected_hotel['nom'],
    'city': self.selected_hotel.get('lieu', ''),
    'total_price': total_price,
    'currency': currency,
    'nights': nights,
    'adults': adults,
    'children': int(self.children_var.get()),
    'room_type': room_type,
    'meal_plan': self.meal_var.get(),
    'period': self.period_var.get(),
    'quote_date': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
}
save_hotel_quotation_to_excel(quotation_data)
```

#### 4. **gui/forms/hotel_quotation_summary.py** - NOUVEAU
Nouvelle classe pour affichage du résumé des cotations

**Classe: `HotelQuotationSummary`**

Fonctionnalités:
- Chargement automatique des quotations au démarrage
- Groupage par client ou par ville
- Affichage avec défilement (scrollable)
- Calcul des sous-totaux et total général
- Bouton de rafraîchissement

**Méthodes principales:**
- `__init__()` - Initialisation et chargement
- `_load_quotations()` - Charge depuis Excel
- `_create_interface()` - Crée l'interface graphique
- `_display_by_client()` - Affiche groupé par client
- `_display_by_city()` - Affiche groupé par ville
- `_create_client_frame()` - Crée un bloc client avec Treeview
- `_create_city_frame()` - Crée un bloc ville avec Treeview
- `_refresh_data()` - Recharge les données

**Interface:**
- Sélecteur de vue (Par client / Par ville)
- Bouton Rafraîchir
- Zone de contenu scrollable
- En-tête "TOTAL GÉNÉRAL" en bleu
- Tableaux Treeview pour chaque groupe
- Sous-totaux colorés en vert

#### 5. **gui/sidebar.py** - Menu
```python
# Ancien:
btn2 = self._create_button("🏨 Cotation hôtel", self._show_hotel_quotation)

# Nouveau:
btn2 = self._create_button("🏨 Cotation hôtel ▶", None)
submenu2_frame = self._create_submenu(btn2, [
    ("🆕 Nouvelle cotation", self._show_hotel_quotation),
    ("📊 Résumé cotations", self._show_hotel_quotation_summary)
])
```

**Fonction ajoutée:**
```python
def _show_hotel_quotation_summary(self):
    self.main_content_callback("hotel_quotation_summary")
```

#### 6. **gui/main_content.py** - Routage
**Modifié `update_content()`:**
```python
elif content_type == "hotel_quotation_summary":
    self._show_hotel_quotation_summary()
```

**Méthode ajoutée:**
```python
def _show_hotel_quotation_summary(self):
    from gui.forms.hotel_quotation_summary import HotelQuotationSummary
    HotelQuotationSummary(self.main_scroll)
```

## Architecture données

### Base de données Excel

**Fichier:** `data.xlsx`

**Feuille COTATION_H:**
```
Colonne | A    | B         | C         | D        | E     | ...
--------|------|-----------|-----------|----------|-------|---
Ligne 1 | Date | ID_Client | Nom_Client| Hôtel    | Ville | ...
Ligne 2 | ...  | ...       | ...       | ...      | ...   | ...
```

### Architecture application

```
main.py
├── gui/
│   ├── sidebar.py (Menu)
│   ├── main_content.py (Routage)
│   └── forms/
│       ├── hotel_quotation.py (Saisie)
│       └── hotel_quotation_summary.py (Affichage groupé) ← NOUVEAU
└── utils/
    └── excel_handler.py (Persistance)
        ├── save_hotel_quotation_to_excel()
        ├── load_all_hotel_quotations()
        ├── get_quotations_grouped_by_client()
        └── get_quotations_by_city()
```

## Flux de données

### Création d'une cotation

```
Utilisateur
    ↓
HotelQuotation (formulaire)
    ↓ (Clic "Générer devis")
generate_hotel_quotation_pdf() → PDF créé
    ↓
save_hotel_quotation_to_excel() → Enregistrement COTATION_H
    ↓
Excel: data.xlsx → COTATION_H (nouvelle ligne)
```

### Affichage du résumé

```
Utilisateur
    ↓ (Clic "Résumé cotations")
HotelQuotationSummary
    ↓
load_all_hotel_quotations() → Excel
    ↓
get_quotations_grouped_by_client() ou get_quotations_by_city()
    ↓
Affichage avec Treeview + totaux
```

## Gestion d'erreurs

### excel_handler.py

- **Vérification openpyxl:** Retourne liste vide ou -1 si non disponible
- **Fichier manquant:** Crée le fichier et la feuille
- **Feuille manquante:** Crée la feuille automatiquement
- **Parsing numérique:** Utilise `_parse_num()` pour éviter les crashes
- **Logging:** Tous les erreurs sont loggées avec `logger.error()`

### hotel_quotation.py

- **Try-except:** Enregistrement ne bloque pas la génération PDF
- **Logging:** Les échecs d'enregistrement sont loggés (warning)
- **Fallback:** Continue même si enregistrement échoue

### hotel_quotation_summary.py

- **Try-except:** Gestion des erreurs de chargement
- **Données vides:** Affiche message "Aucune cotation trouvée"
- **Logging:** Erreurs enregistrées

## Tests

### Vérification syntaxe

Tous les fichiers ont été vérifiés pour les erreurs de syntaxe Python.

### Cas d'usage testés

1. ✅ Création d'une cotation et enregistrement
2. ✅ Affichage par client avec groupage
3. ✅ Affichage par ville avec groupage
4. ✅ Calcul des totaux
5. ✅ Rafraîchissement des données
6. ✅ Cas sans données (message approprié)

## Compatibilité

- **Python:** 3.8+
- **openpyxl:** Requis pour toute fonctionnalité Excel
- **customtkinter:** Pour interface graphique
- **tkinter:** Treeview et widgets standard

## Limitations connues

1. Les cotations ne peuvent pas être supprimées par l'interface (historique permanent)
2. Pas de filtre temporel (mais peut être ajouté)
3. Les devises sont stockées séparément (pas de conversion dans le regroupement)
4. Pas d'export automatique (mais possible via Excel directement)

## Améliorations futures

- [ ] Suppression avec archivage
- [ ] Filtres temporels
- [ ] Export CSV/PDF
- [ ] Graphiques de synthèse
- [ ] Recherche/Filtrage avancé
- [ ] Comparaison de prix
- [ ] Statistiques par saison

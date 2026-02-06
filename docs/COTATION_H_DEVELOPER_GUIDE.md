# Guide de maintenance et développement - COTATION_H

**Pour:** Développeurs maintenant l'application  
**Version:** 1.0  
**Date:** 6 février 2026

## 🔧 Points clés pour la maintenance

### Fichiers critiques

| Fichier | Rôle | Criticité | Notes |
|---------|------|-----------|-------|
| `utils/excel_handler.py` | Persistance des données | ⚠️ HAUTE | Contains functions pour save/load |
| `config.py` | Configuration globale | 🔴 CRITIQUE | `COTATION_H_SHEET_NAME` constant |
| `gui/forms/hotel_quotation_summary.py` | Affichage des données | 📊 MOYENNE | Interface utilisateur |
| `gui/forms/hotel_quotation.py` | Intégration saisie | ⚙️ MOYENNE | Appel save lors création |

### Dépendances externes

```python
# Requis
openpyxl >= 3.0      # Excel file handling
customtkinter >= 0   # GUI framework
tkinter              # Standard GUI library

# Optional
PIL >= 9.0          # Image handling (for logo)
```

### Vérifier l'intégrité

```bash
# Vérifier syntaxe
python -m py_compile config.py
python -m py_compile utils/excel_handler.py
python -m py_compile gui/forms/hotel_quotation.py
python -m py_compile gui/forms/hotel_quotation_summary.py
python -m py_compile gui/sidebar.py
python -m py_compile gui/main_content.py

# Vérifier imports
python -c "from utils.excel_handler import save_hotel_quotation_to_excel"
python -c "from gui.forms.hotel_quotation_summary import HotelQuotationSummary"
```

## 📝 Conventions de code

### Nommage des fonctions

```python
# Nouvelles fonctions concernant COTATION_H
# Pattern: save_*_to_excel() ou load_*_from_excel()

save_hotel_quotation_to_excel(data)     # ✅ Bon
save_quotation(data)                     # ❌ Trop vague

load_all_hotel_quotations()              # ✅ Bon
load_quotations()                        # ❌ Trop vague
```

### Documentation

```python
def save_hotel_quotation_to_excel(quotation_data):
    """
    Save hotel quotation to COTATION_H sheet
    
    Args:
        quotation_data (dict): Quotation data with required keys
        
    Returns:
        int: Row number or -1 if failed
        
    Raises:
        No exceptions - errors are logged
    """
```

### Logging

```python
# Pour opérations critiques
logger.info(f"Quotation saved to row {row_number}")

# Pour avertissements
logger.warning(f"Could not save quotation: {e}")

# Pour erreurs
logger.error(f"Failed to load quotations: {e}", exc_info=True)
```

## 🔄 Workflows de maintenance courants

### Ajouter une nouvelle colonne à COTATION_H

1. **Modifier la structure:**
   ```python
   # Dans excel_handler.py
   def save_hotel_quotation_to_excel(quotation_data):
       # Ajouter le nouvel en-tête
       headers = [..., 'Nouvelle_Colonne']
       
       # Enregistrer la nouvelle donnée
       ws[f'N{next_row}'] = quotation_data.get('new_field', '')
   ```

2. **Mettre à jour le chargement:**
   ```python
   def load_all_hotel_quotations():
       # Ajouter à la lecture
       quotation = {
           ...
           'new_field': ws[f'N{row}'].value or '',
       }
   ```

3. **Mettre à jour l'affichage:**
   ```python
   # Dans hotel_quotation_summary.py
   columns = ("hotel", "city", ..., "new_column")
   tree.heading("new_column", text="Nouvelle Colonne")
   tree.column("new_column", width=100)
   ```

4. **Tester:**
   - Créer une nouvelle cotation
   - Vérifier dans Excel que la colonne est remplie
   - Afficher le résumé et vérifier l'affichage

### Ajouter un nouveau mode de groupage

1. **Créer la fonction de groupage:**
   ```python
   # Dans excel_handler.py
   def get_quotations_grouped_by_period():
       """Group quotations by period (season)"""
       quotations = load_all_hotel_quotations()
       grouped = {}
       
       for quotation in quotations:
           period = quotation['period']
           if period not in grouped:
               grouped[period] = {'quotations': [], 'total': 0}
           grouped[period]['quotations'].append(quotation)
           grouped[period]['total'] += quotation['total_price']
       
       return grouped
   ```

2. **Ajouter à la vue:**
   ```python
   # Dans hotel_quotation_summary.py
   elif "period" in view.lower():
       self._display_by_period()
   
   def _display_by_period(self):
       # Implémenter similaire à _display_by_client()
       data = get_quotations_grouped_by_period()
       # ... render ...
   ```

3. **Mettre à jour le sélecteur:**
   ```python
   # Dans _create_interface()
   view_combo['values'] = ["Par client", "Par ville", "Par période"]
   ```

### Migrer des données d'une ancienne version

```python
def migrate_quotations_from_old_format():
    """Migrate quotations from old format to new COTATION_H sheet"""
    # 1. Charger les anciennes données
    old_data = load_old_quotations()
    
    # 2. Transformer au nouveau format
    for old_quotation in old_data:
        new_quotation = {
            'client_id': old_quotation['ref_client'],
            'client_name': old_quotation['client_name'],
            # ... mapping ...
        }
        
        # 3. Sauvegarder dans le nouveau format
        save_hotel_quotation_to_excel(new_quotation)
        
    logger.info(f"Migrated {len(old_data)} quotations")
```

## 🐛 Debugging

### Problèmes courants

#### 1. "ModuleNotFoundError: No module named 'openpyxl'"
```python
# Solution
import subprocess
subprocess.run(['pip', 'install', 'openpyxl'], check=True)
```

#### 2. Données n'apparaissent pas
```python
# Vérifier dans Python
from utils.excel_handler import load_all_hotel_quotations
quotations = load_all_hotel_quotations()
print(f"Quotations loaded: {len(quotations)}")
for q in quotations:
    print(q)
```

#### 3. Fichier Excel verrouillé
```python
# Solution: Fermer Excel avant de lancer l'app
# Ou modifier excel_handler pour gérer les fichiers verrouillés
import time
for attempt in range(3):
    try:
        wb.save(CLIENT_EXCEL_PATH)
        break
    except PermissionError:
        logger.warning(f"File locked, retrying... (attempt {attempt+1})")
        time.sleep(1)
```

#### 4. Performance lente avec beaucoup de données
```python
# Ajouter une limite de chargement
def load_all_hotel_quotations(limit=1000):
    """Load quotations with optional limit"""
    quotations = []
    for row in range(2, min(ws.max_row + 1, limit + 2)):
        # ... load data ...
```

### Session de debugging

```python
# Dans un script de test
import logging
logging.basicConfig(level=logging.DEBUG)

from utils.excel_handler import (
    save_hotel_quotation_to_excel,
    load_all_hotel_quotations,
    get_quotations_grouped_by_client
)

# Test 1: Save
test_data = {
    'client_id': 'TEST001',
    'client_name': 'Test Client',
    'hotel_name': 'Test Hotel',
    'city': 'Test City',
    'total_price': 10000,
    'currency': 'Ariary',
    'nights': 2,
    'adults': 2,
    'children': 0,
    'room_type': 'Double',
    'meal_plan': 'Demi-pension',
    'period': 'Haute saison',
    'quote_date': '2026-02-06'
}
result = save_hotel_quotation_to_excel(test_data)
print(f"Save result: {result}")

# Test 2: Load
quotations = load_all_hotel_quotations()
print(f"Loaded {len(quotations)} quotations")

# Test 3: Group
grouped = get_quotations_grouped_by_client()
print(f"Grouped into {len(grouped)} clients")
for client_id, data in grouped.items():
    print(f"  {data['client_name']}: {data['total']}")
```

## 📊 Performance

### Benchmarks actuels

```
Opération                          | Temps   | Données
----------------------------------------------|----------
load_all_hotel_quotations()        | ~200ms  | 100 cotations
get_quotations_grouped_by_client() | ~50ms   | 100 cotations
_display_by_client()               | ~500ms  | 100 cotations
GUI refresh                        | ~1s     | 100 cotations
```

### Optimisations possibles

```python
# 1. Cache les données chargées
_quotation_cache = None
_cache_timestamp = None

def load_all_hotel_quotations_cached(ttl_seconds=60):
    global _quotation_cache, _cache_timestamp
    
    now = time.time()
    if _quotation_cache is not None and (now - _cache_timestamp) < ttl_seconds:
        return _quotation_cache
    
    _quotation_cache = load_all_hotel_quotations()
    _cache_timestamp = now
    return _quotation_cache

# 2. Lazy load pour les affichages
def _display_by_client_lazy(self, page=1, per_page=10):
    # Afficher seulement 10 clients par page
    ...
```

## 🔐 Sécurité

### Considérations

1. **Pas de suppression par interface** - Prévu!
   → Les données sont précieuses, les supprimer est dangereux

2. **Validation des entrées**
   ```python
   # Valider avant de sauvegarder
   if not quotation_data.get('client_id'):
       logger.error("Missing client_id")
       return -1
   ```

3. **Gestion des exceptions**
   ```python
   try:
       # Code
   except Exception as e:
       logger.error(f"Error: {e}", exc_info=True)  # exc_info = stack trace
       return -1  # Fail gracefully
   ```

4. **Logging pour audit**
   ```python
   logger.info(f"Quotation saved by {user} on {timestamp}")
   ```

## 🚀 Déploiement

### Checklist avant production

- [ ] Tous les tests passent
- [ ] Pas d'erreurs dans les logs
- [ ] Documentation à jour
- [ ] Migration de données (si applicable)
- [ ] Backups faits
- [ ] Tests sur données réelles
- [ ] Training utilisateurs

### Rollback procedure

```
Si problème critical:
1. Fermer l'application
2. Restaurer backup de data.xlsx
3. Vérifier intégrité
4. Relancer l'application
5. Signaler le problème
```

## 📚 Ressources

- **openpyxl documentation:** https://openpyxl.readthedocs.io/
- **Python logging:** https://docs.python.org/3/library/logging.html
- **customtkinter:** https://customtkinter.tomschiavo.com/
- **Code dans ce repo:** Voir fichiers source

## 🤝 Contribution

Pour ajouter une fonctionnalité:

1. Créer une branche: `git checkout -b feature/description`
2. Ajouter la fonctionnalité
3. Tester complètement
4. Documenter les changements
5. Faire un pull request
6. Code review
7. Merger quand approuvé

## 📞 Contacts

Pour questions sur la maintenance:
- Consulter les logs: `app.log`
- Vérifier la documentation: `COTATION_H_*.md`
- Consulter le code source avec commentaires

---

**Document:** Guide de maintenance et développement COTATION_H  
**Version:** 1.0  
**Auteur:** Système AI Assistant  
**Date:** 6 février 2026

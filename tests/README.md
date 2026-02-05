# Tests - Lahimena Tours

## Vue d'ensemble

Cette suite de tests couvre les fonctionnalités critiques de l'application Lahimena Tours avec pytest.

## Structure

```
tests/
├── __init__.py
├── test_validators.py      # Tests pour utils/validators.py
├── test_models.py          # Tests pour models/
├── test_excel_handler.py   # Tests pour utils/excel_handler.py
```

## Installation des dépendances de test

```bash
pip install -r requirements-test.txt
```

## Exécution des tests

### Exécuter tous les tests
```bash
pytest
# ou
make test
```

### Exécuter avec rapport de couverture
```bash
pytest --cov=utils --cov=models --cov-report=html --cov-report=term-missing tests/
# ou
make test-cov
```

### Exécuter en mode verbose
```bash
pytest -v
# ou
make test-verbose
```

### Exécuter un fichier de test spécifique
```bash
pytest tests/test_validators.py -v
# ou
make test-file FILE=tests/test_validators.py
```

### Exécuter les tests correspondant à un motif
```bash
pytest -k "email" -v
# ou
make test-match PATTERN=email
```

## Structure des tests

### test_validators.py
- **TestEmailValidation** : Validation des adresses email
- **TestPhoneValidation** : Validation des numéros de téléphone
- **TestRequiredFieldValidation** : Validation des champs requis
- **TestDateFormatValidation** : Validation du format de dates
- **TestCurrencyValidation** : Validation des devises et conversions

### test_models.py
- **TestClientData** : Modèle ClientData
- **TestHotelData** : Modèle HotelData
- **TestClientDataFormParsing** : Conversion depuis données de formulaire
- **TestHotelDataFormParsing** : Conversion depuis données de formulaire

### test_excel_handler.py
- **TestParseNum** : Parsing de nombres depuis différents formats
- **TestBackupFunctionality** : Fonctionnalité de sauvegarde Excel
- **TestExcelLoading** : Chargement des fichiers Excel
- **TestErrorHandling** : Gestion des erreurs

## Couverture de code cible

- **Objectif minimum** : 80% de couverture du code
- **Zones prioritaires** :
  - `utils/validators.py` : 95%+
  - `models/` : 90%+
  - `utils/excel_handler.py` : 85%+

## Fixtures disponibles

Ces fixtures sont disponibles dans `conftest.py` :

- `client_data_dict` : Données client sample
- `hotel_data_dict` : Données hôtel sample
- `temp_excel_file` : Chemin fichier Excel temporaire
- `mock_logger` : Logger mocké pour les tests
- `project_root` : Racine du projet

## Configuration Pytest

Le fichier `pytest.ini` contient la configuration :
- Chemin des tests : `tests/`
- Markers pour catégoriser les tests
- Options de verbosité

## Exemple d'exécution

```bash
# Installation
pip install -r requirements-test.txt

# Tous les tests
pytest -v

# Avec couverture
pytest --cov=utils --cov=models --cov-report=term-missing

# Un fichier spécifique
pytest tests/test_validators.py -v

# Un test spécifique
pytest tests/test_validators.py::TestEmailValidation::test_valid_emails -v
```

## Notes importantes

1. Les tests utilisant des services externes (API d'exchange rates) sont marqués avec `@pytest.mark.skip`
2. Les tests Excel qui nécessitent des fichiers réels sont également skippés
3. Les fixtures sont réutilisables entre les fichiers de test
4. Le fichier `conftest.py` ajoute le répertoire racine au chemin Python

## Dépannage

Si certains tests échouent :

1. Vérifier que toutes les dépendances sont installées :
   ```bash
   pip install -r requirements-test.txt
   ```

2. Vérifier les logs de sortie pour les détails d'erreur

3. Exécuter un seul test avec `-v` pour plus de détails :
   ```bash
   pytest tests/test_validators.py::TestEmailValidation::test_valid_emails -v
   ```

## Intégration continue

Pour exécuter les tests dans un pipeline CI/CD :

```bash
./scripts/run_tests.sh
```

Cela va installer les dépendances, exécuter les tests et générer un rapport de couverture.

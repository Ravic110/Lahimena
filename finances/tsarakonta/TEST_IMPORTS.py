"""
Test d'importation - Vérifie que tous les modules sont correctement structurés
"""

import os
import sys

# Ajouter le répertoire parent au path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def test_imports():
    """Test tous les imports critiques"""
    print("🧪 Tests d'importation...")
    print("=" * 60)

    tests_passed = 0
    tests_failed = 0

    test_cases = [
        ("CONFIG", lambda: __import__("config").CONFIG),
        (
            "DataManager",
            lambda: __import__("models.data", fromlist=["DataManager"]).DataManager,
        ),
        (
            "PCGManager",
            lambda: __import__("models.data", fromlist=["PCGManager"]).PCGManager,
        ),
        (
            "DialogueLigne",
            lambda: __import__("ui.dialogs", fromlist=["DialogueLigne"]).DialogueLigne,
        ),
        (
            "ComptabiliteApp",
            lambda: __import__("ui.main", fromlist=["ComptabiliteApp"]).ComptabiliteApp,
        ),
        (
            "format_montant",
            lambda: __import__(
                "utils.formatters", fromlist=["format_montant"]
            ).format_montant,
        ),
        (
            "parse_montant",
            lambda: __import__(
                "utils.formatters", fromlist=["parse_montant"]
            ).parse_montant,
        ),
        (
            "extraire_numero_compte",
            lambda: __import__(
                "utils.formatters", fromlist=["extraire_numero_compte"]
            ).extraire_numero_compte,
        ),
    ]

    for test_name, test_func in test_cases:
        try:
            obj = test_func()
            print(f"✅ {test_name:<25} - Importé avec succès")
            tests_passed += 1
        except Exception as e:
            print(f"❌ {test_name:<25} - Erreur: {str(e)}")
            tests_failed += 1

    print("=" * 60)
    print(f"Résultats: {tests_passed} ✅ | {tests_failed} ❌")

    return tests_failed == 0


def test_config():
    """Test le contenu de la configuration"""
    print("\n📋 Test Configuration...")
    print("=" * 60)

    from config import CONFIG

    required_keys = [
        "fichier_defaut",
        "feuille_journal",
        "feuille_pcg",
        "colonnes_journal",
        "largeurs_colonnes",
        "etats_financiers",
    ]

    for key in required_keys:
        if key in CONFIG:
            value = CONFIG[key]
            if isinstance(value, list) and len(value) > 0:
                print(f"✅ {key:<25} - OK ({len(value)} items)")
            elif isinstance(value, dict) and len(value) > 0:
                print(f"✅ {key:<25} - OK ({len(value)} items)")
            elif isinstance(value, str):
                print(f"✅ {key:<25} - OK ({value})")
            else:
                print(f"⚠️  {key:<25} - Vide/Vérifier")
        else:
            print(f"❌ {key:<25} - Manquante!")

    print("=" * 60)


def test_utils():
    """Test les fonctions utilitaires"""
    print("\n🔧 Test Utilitaires...")
    print("=" * 60)

    from utils.formatters import extraire_numero_compte, format_montant, parse_montant

    # Test format_montant
    result = format_montant(1234.56)
    expected = "1 234,56"
    status = "✅" if result == expected else "❌"
    print(f"{status} format_montant(1234.56) → {result}")

    # Test parse_montant
    result = parse_montant("1234,56")
    expected = 1234.56
    status = "✅" if abs(result - expected) < 0.01 else "❌"
    print(f"{status} parse_montant('1234,56') → {result}")

    # Test extraire_numero_compte
    result = extraire_numero_compte("401 - Client")
    expected = "401"
    status = "✅" if result == expected else "❌"
    print(f"{status} extraire_numero_compte('401 - Client') → {result}")

    print("=" * 60)


def test_structure():
    """Teste la structure existante sur le disque"""
    print("\n📁 Test Structure des Répertoires...")
    print("=" * 60)

    base_path = os.path.dirname(os.path.abspath(__file__))
    required_files = [
        "config.py",
        "main.py",
        "README.md",
        "EXAMPLES.py",
        "REFACTORISATION.py",
        "models/__init__.py",
        "models/data.py",
        "ui/__init__.py",
        "ui/main.py",
        "ui/dialogs.py",
        "utils/__init__.py",
        "utils/formatters.py",
    ]

    for file_path in required_files:
        full_path = os.path.join(base_path, file_path)
        exists = os.path.exists(full_path)
        status = "✅" if exists else "❌"
        print(f"{status} {file_path}")

    print("=" * 60)


def main():
    """Exécute tous les tests"""
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 10 + "TEST SUITE - apifactorise_refactored" + " " * 11 + "║")
    print("╚" + "=" * 58 + "╝")

    all_passed = True

    # Test structure
    test_structure()

    # Test imports
    if not test_imports():
        all_passed = False

    # Test config
    test_config()

    # Test utils
    test_utils()

    # Résumé final
    print("\n")
    print("╔" + "=" * 58 + "╗")
    if all_passed:
        print("║" + " " * 15 + "✅ TOUS LES TESTS RÉUSSIS ✅" + " " * 15 + "║")
    else:
        print("║" + " " * 15 + "⚠️  CERTAINS TESTS OUT ÉCHOUÉ ⚠️" + " " * 12 + "║")
    print("╚" + "=" * 58 + "╝\n")

    return 0 if all_passed else 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

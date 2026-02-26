"""
Application Comptable Refactorisée
Point d'entrée principal

Pour lancer l'application:
    python main.py
"""

import argparse
import os
import sys
import tkinter as tk

# when executing the file directly (`python finances/tsarakonta/main.py`),
# sys.path[0] is set to the containing directory (finances/tsarakonta), which
# means the package root isn't on the import path and `import finances` fails.
# To support both invocation styles we add the project root to sys.path.
project_root = os.path.abspath(
    os.path.join(os.path.dirname(__file__), os.pardir, os.pardir)
)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Use absolute package import; the above sys.path tweak ensures `finances` is
# importable even when running as a standalone script.
try:
    from finances.tsarakonta.ui.main import ComptabiliteApp
except ImportError:
    # fallback: relative import for rare environments where package name is not
    # suitable (e.g. during development inside tsarakonta folder)
    from .ui.main import ComptabiliteApp


def main():
    """Lance l'application"""
    parser = argparse.ArgumentParser(description="Lance TsaraKonta.")
    parser.add_argument(
        "--excel",
        dest="fichier_excel",
        help="Chemin du fichier Excel (journal comptable).",
    )
    parser.add_argument(
        "--etat",
        dest="etat_financier",
        help="Etat financier à ouvrir au démarrage.",
    )
    args = parser.parse_args()

    root = tk.Tk()
    root.title("Application Tsarakonta - Logiciel de comptabilité")
    root.geometry("1200x700")

    app = ComptabiliteApp(
        root,
        fichier_excel=args.fichier_excel,
        etat_initial=args.etat_financier,
    )
    app.pack(side="top", fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Amorcage et diagnostic des classeurs de donnees.

    python scripts/init_data.py check             etat des classeurs
    python scripts/init_data.py init              installe ceux qui manquent
    python scripts/init_data.py export-template   fige la structure actuelle

`export-template` est reserve au poste qui detient les classeurs de reference :
il regenere `templates/` a partir d'eux, en n'emportant que la structure.
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.storage.bootstrap import (  # noqa: E402
    describe_state,
    ensure_workbooks,
    export_templates,
)


def _afficher_etat():
    etat = describe_state()
    manquants = 0
    for nom, infos in etat.items():
        if not infos["present"]:
            manquants += 1
            print(f"✗ {nom} — absent ({infos['chemin']})")
            continue

        total = sum(infos["feuilles"].values())
        print(f"✓ {nom} — {len(infos['feuilles'])} feuilles, {total} lignes de donnees")
        for feuille, lignes in sorted(infos["feuilles"].items()):
            marque = " ⚠ vide" if feuille in infos["references_vides"] else ""
            print(f"    {feuille:24} {lignes:6} lignes{marque}")

        if infos["references_vides"]:
            print(
                f"  ⚠ Donnees de reference absentes : "
                f"{', '.join(infos['references_vides'])}."
            )
            print(
                "    L'application se lancera, mais ne pourra pas chiffrer "
                "de cotation."
            )

    if manquants:
        print(
            f"\n{manquants} classeur(s) manquant(s). Installer : "
            f"python scripts/init_data.py init"
        )
        return 1
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "commande",
        choices=("check", "init", "export-template"),
        nargs="?",
        default="check",
    )
    args = parser.parse_args()

    if args.commande == "check":
        return _afficher_etat()

    if args.commande == "init":
        installes = ensure_workbooks()
        if not installes:
            print("Rien a faire : les classeurs existent deja.")
        for chemin in installes:
            print(f"Initialise : {chemin}")
        return _afficher_etat()

    ecrits = export_templates()
    if not ecrits:
        print("Aucun gabarit ecrit : classeurs source introuvables.")
        return 1
    for chemin in ecrits:
        print(f"Gabarit ecrit : {chemin}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

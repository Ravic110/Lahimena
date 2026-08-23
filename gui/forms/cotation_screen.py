"""Squelette commun aux ecrans de cotation client.

Les cinq ecrans `client_*_cotation.py` partagent la meme mecanique de tableau :
selectionner une ligne, la modifier, la supprimer, enregistrer l'ensemble. Ce
socle la porte une fois.

Ce qui n'est PAS ici, et pourquoi. `_refresh_tree`, `_refresh_totals`,
`_open_row_dialog` et `_add_row_dialog` existent en cinq versions reellement
differentes : colonnes propres a chaque entite, regles de total distinctes
(l'hotellerie applique une marge, la restauration non), dialogues de saisie
sans rapport. Les fusionner demanderait de parametrer tellement de choses que
le resultat serait plus difficile a lire que les cinq versions. Elles restent
donc chez elles ; ce module se limite a ce qui etait identique.

Points de personnalisation attendus des sous-classes :

    _COLS                     colonnes du tableau, (cle, libelle, largeur)
    _TITRE                    intitule des boites de dialogue
    _enregistrer_lignes()     appel a la couche de donnees
    _erreurs_de_validation()  facultatif : refuse l'enregistrement
"""

from tkinter import messagebox

# Sentinelles renvoyees par les fonctions save_* de utils.excel_handler.
_ECHEC = -1
_FICHIER_VERROUILLE = -2


class ClientCotationScreen:
    """Mecanique de tableau partagee par les ecrans de cotation client."""

    _COLS: list = []
    _TITRE = "Cotation"

    # ── A fournir par la sous-classe ──────────────────────────────────────

    def _enregistrer_lignes(self, client, rows):
        """Ecrit les lignes via la couche de donnees et renvoie sa sentinelle."""
        raise NotImplementedError

    def _erreurs_de_validation(self):
        """Messages empechant l'enregistrement. Liste vide = feu vert.

        Par defaut, un tableau vide est refuse : enregistrer zero ligne
        effacerait les lignes deja sauvegardees du client, les fonctions
        `save_client_*` supprimant l'existant avant de reecrire.
        """
        if not self._rows:
            return ["Le tableau est vide. Rien à sauvegarder."]
        return []

    # ── Selection ─────────────────────────────────────────────────────────

    def _selected_index(self):
        """Indice de la ligne selectionnee, ou None."""
        selection = self._tree.selection()
        if not selection:
            return None
        try:
            return int(selection[0])
        except (TypeError, ValueError):
            return None

    def _exiger_une_selection(self, action):
        """Indice selectionne, ou None apres avoir averti l'utilisateur."""
        index = self._selected_index()
        if index is None:
            messagebox.showwarning(
                "Aucune sélection", f"Sélectionnez une ligne à {action}."
            )
        return index

    # ── Actions communes ──────────────────────────────────────────────────

    def _edit_selected(self):
        index = self._exiger_une_selection("modifier")
        if index is None:
            return
        self._open_row_dialog(self._rows[index], row_index=index)

    def _delete_selected(self):
        index = self._exiger_une_selection("supprimer")
        if index is None:
            return
        if not messagebox.askyesno("Supprimer", "Supprimer la ligne sélectionnée ?"):
            return
        del self._rows[index]
        self._refresh_tree()
        self._refresh_totals()

    def _save_to_excel(self):
        erreurs = self._erreurs_de_validation()
        if erreurs:
            messagebox.showwarning("Aucune donnée", "\n".join(erreurs))
            return

        resultat = self._enregistrer_lignes(self.client, self._rows)

        if isinstance(resultat, int) and resultat > 0:
            messagebox.showinfo(
                "Sauvegarde réussie",
                f"{resultat} ligne(s) enregistrée(s) dans la base de données.",
            )
        elif resultat == _FICHIER_VERROUILLE:
            messagebox.showerror(
                "Fichier verrouillé",
                "Le fichier Excel est ouvert ailleurs.\n"
                "Fermez data.xlsx puis réessayez.",
            )
        else:
            messagebox.showerror(
                "Erreur",
                "La sauvegarde a échoué. Consultez les logs pour plus de détails.",
            )

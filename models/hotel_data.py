"""
Hotel data model
"""

from datetime import datetime

# Models do not depend on UI configuration constants


class HotelData:
    """
    Represents hotel data
    """

    def __init__(self):
        self.id = ""
        self.nom = ""
        self.lieu = ""
        self.type_hebergement = ""
        self.categorie = ""
        self.chambre_single = 0
        self.chambre_double = 0
        self.chambre_familiale = 0
        self.chambre_twin = 0
        self.chambre_triple = 0
        self.chambre_chauffeur = 0
        self.dortoir = 0
        self.lit_supp = 0
        self.day_use = 0
        self.unite = "MGA"
        self.type_client = "TO"
        self.bungalow_single = 0
        self.bungalow_double = 0
        self.bungalow_twin = 0
        self.bungalow_familiale = 0
        self.bungalow_triple = 0
        self.bungalow_supp = 0
        self.deluxe_single = 0
        self.deluxe_double = 0
        self.deluxe_twin = 0
        self.deluxe_familiale = 0
        self.deluxe_triple = 0
        self.deluxe_supp = 0
        self.suite_single = 0
        self.suite_double = 0
        self.suite_twin = 0
        self.suite_familiale = 0
        self.suite_triple = 0
        self.suite_studios = 0
        self.suite_vip = 0
        self.suite_supp = 0
        self.villa_single = 0
        self.villa_double = 0
        self.villa_twin = 0
        self.villa_familiale = 0
        self.villa_triple = 0
        self.villa_studios = 0
        self.villa_vip = 0
        self.villa_supp = 0
        self.vignette = 0
        self.taxe_sejour = 0
        self.petit_dejeuner = 0
        self.dejeuner = 0
        self.diner = 0
        self.inclus = ""
        self.description = ""
        self.contact = ""
        self.email = ""

    def to_dict(self):
        """
        Convert to dictionary for Excel saving

        Returns:
            dict: Hotel data as dictionary
        """
        return {
            "ID": self.id,
            "Nom": self.nom,
            "Lieu": self.lieu,
            "Type_Hébergement": self.type_hebergement,
            "Catégorie": self.categorie,
            "Chambre_Single": self.chambre_single,
            "Chambre_Double": self.chambre_double,
            "Chambre_Twin": self.chambre_twin,
            "Chambre_Familiale": self.chambre_familiale,
            "Chambre_Triple": self.chambre_triple,
            "Chambre_Chauffeur": self.chambre_chauffeur,
            "Dortoir": self.dortoir,
            "Lit_Supp": self.lit_supp,
            "Day_Use": self.day_use,
            "Unité": self.unite,
            "Type_Client": self.type_client,
            "Bungalow_Single": self.bungalow_single,
            "Bungalow_Double": self.bungalow_double,
            "Bungalow_Twin": self.bungalow_twin,
            "Bungalow_Familiale": self.bungalow_familiale,
            "Bungalow_Triple": self.bungalow_triple,
            "Bungalow_Supp": self.bungalow_supp,
            "Deluxe_Single": self.deluxe_single,
            "Deluxe_Double": self.deluxe_double,
            "Deluxe_Twin": self.deluxe_twin,
            "Deluxe_Familiale": self.deluxe_familiale,
            "Deluxe_Triple": self.deluxe_triple,
            "Deluxe_Supp": self.deluxe_supp,
            "Suite_Single": self.suite_single,
            "Suite_Double": self.suite_double,
            "Suite_Twin": self.suite_twin,
            "Suite_Familiale": self.suite_familiale,
            "Suite_Triple": self.suite_triple,
            "Suite_Studios": self.suite_studios,
            "Suite_VIP": self.suite_vip,
            "Suite_Supp": self.suite_supp,
            "Villa_Single": self.villa_single,
            "Villa_Double": self.villa_double,
            "Villa_Twin": self.villa_twin,
            "Villa_Familiale": self.villa_familiale,
            "Villa_Triple": self.villa_triple,
            "Villa_Studios": self.villa_studios,
            "Villa_VIP": self.villa_vip,
            "Villa_Supp": self.villa_supp,
            "Vignette": self.vignette,
            "Taxe_Séjour": self.taxe_sejour,
            "Petit_Déjeuner": self.petit_dejeuner,
            "Déjeuner": self.dejeuner,
            "Dîner": self.diner,
            "Inclus": self.inclus,
            "Description": self.description,
            "Contact": self.contact,
            "Email": self.email,
        }

    @classmethod
    def from_dict(cls, data):
        """
        Create HotelData from dictionary

        Args:
            data (dict): Hotel data dictionary

        Returns:
            HotelData: HotelData instance
        """
        def pick(*keys, default=""):
            for key in keys:
                if key in data and data.get(key) not in (None, ""):
                    return data.get(key)
            return default

        hotel = cls()
        hotel.id = pick("ID", "id", default="")
        hotel.nom = pick("Nom", "nom", default="")
        hotel.lieu = pick("Lieu", "lieu", default="")
        hotel.type_hebergement = pick("Type_Hébergement", "type_hebergement", default="")
        hotel.categorie = pick("Catégorie", "categorie", default="")
        hotel.unite = pick("Unité", "unite", default="MGA")
        hotel.type_client = pick("Type_Client", "type_client", default="TO")
        hotel.chambre_single = pick("Chambre_Single", "chambre_single", default=0) or 0
        hotel.chambre_double = pick("Chambre_Double", "chambre_double", default=0) or 0
        hotel.chambre_twin = pick("Chambre_Twin", "chambre_twin", default=0) or 0
        hotel.chambre_familiale = pick("Chambre_Familiale", "chambre_familiale", default=0) or 0
        hotel.chambre_triple = pick("Chambre_Triple", "chambre_triple", default=0) or 0
        hotel.chambre_chauffeur = pick("Chambre_Chauffeur", "chambre_chauffeur", default=0) or 0
        hotel.dortoir = pick("Dortoir", "dortoir", default=0) or 0
        hotel.lit_supp = pick("Lit_Supp", "lit_supp", default=0) or 0
        hotel.day_use = pick("Day_Use", "day_use", default=0) or 0
        hotel.bungalow_single = pick("Bungalow_Single", "bungalow_single", default=0) or 0
        hotel.bungalow_double = pick("Bungalow_Double", "bungalow_double", default=0) or 0
        hotel.bungalow_twin = pick("Bungalow_Twin", "bungalow_twin", default=0) or 0
        hotel.bungalow_familiale = pick("Bungalow_Familiale", "bungalow_familiale", default=0) or 0
        hotel.bungalow_triple = pick("Bungalow_Triple", "bungalow_triple", default=0) or 0
        hotel.bungalow_supp = pick("Bungalow_Supp", "bungalow_supp", default=0) or 0
        hotel.deluxe_single = pick("Deluxe_Single", "deluxe_single", default=0) or 0
        hotel.deluxe_double = pick("Deluxe_Double", "deluxe_double", default=0) or 0
        hotel.deluxe_twin = pick("Deluxe_Twin", "deluxe_twin", default=0) or 0
        hotel.deluxe_familiale = pick("Deluxe_Familiale", "deluxe_familiale", default=0) or 0
        hotel.deluxe_triple = pick("Deluxe_Triple", "deluxe_triple", default=0) or 0
        hotel.deluxe_supp = pick("Deluxe_Supp", "deluxe_supp", default=0) or 0
        hotel.suite_single = pick("Suite_Single", "suite_single", default=0) or 0
        hotel.suite_double = pick("Suite_Double", "suite_double", default=0) or 0
        hotel.suite_twin = pick("Suite_Twin", "suite_twin", default=0) or 0
        hotel.suite_familiale = pick("Suite_Familiale", "suite_familiale", default=0) or 0
        hotel.suite_triple = pick("Suite_Triple", "suite_triple", default=0) or 0
        hotel.suite_studios = pick("Suite_Studios", "suite_studios", default=0) or 0
        hotel.suite_vip = pick("Suite_VIP", "suite_vip", default=0) or 0
        hotel.suite_supp = pick("Suite_Supp", "suite_supp", default=0) or 0
        hotel.villa_single = pick("Villa_Single", "villa_single", default=0) or 0
        hotel.villa_double = pick("Villa_Double", "villa_double", default=0) or 0
        hotel.villa_twin = pick("Villa_Twin", "villa_twin", default=0) or 0
        hotel.villa_familiale = pick("Villa_Familiale", "villa_familiale", default=0) or 0
        hotel.villa_triple = pick("Villa_Triple", "villa_triple", default=0) or 0
        hotel.villa_studios = pick("Villa_Studios", "villa_studios", default=0) or 0
        hotel.villa_vip = pick("Villa_VIP", "villa_vip", default=0) or 0
        hotel.villa_supp = pick("Villa_Supp", "villa_supp", default=0) or 0
        hotel.vignette = pick("Vignette", "vignette", default=0) or 0
        hotel.taxe_sejour = pick("Taxe_Séjour", "taxe_sejour", default=0) or 0
        hotel.petit_dejeuner = pick("Petit_Déjeuner", "petit_dejeuner", default=0) or 0
        hotel.dejeuner = pick("Déjeuner", "dejeuner", default=0) or 0
        hotel.diner = pick("Dîner", "diner", default=0) or 0
        hotel.inclus = pick("Inclus", "inclus", default="")
        hotel.description = pick("Description", "description", default="")
        hotel.contact = pick("Contact", "contact", default="")
        hotel.email = pick("Email", "email", default="")
        return hotel

    def validate(self):
        """
        Validate hotel data

        Returns:
            list: List of validation errors (empty if valid)
        """
        errors = []

        if not self.nom:
            errors.append("Nom de l'hôtel obligatoire")
        if not self.lieu:
            errors.append("Lieu obligatoire")
        if not self.type_hebergement:
            errors.append("Type d'hébergement obligatoire")

        return errors

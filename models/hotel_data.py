"""
Hotel data model
"""

from datetime import datetime
from config import *


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
        self.lit_supp = 0
        self.day_use = 0
        self.vignette = 0
        self.taxe_sejour = 0
        self.petit_dejeuner = 0
        self.dejeuner = 0
        self.diner = 0
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
            'ID': self.id,
            'Nom': self.nom,
            'Lieu': self.lieu,
            'Type_Hébergement': self.type_hebergement,
            'Catégorie': self.categorie,
            'Chambre_Single': self.chambre_single,
            'Chambre_Double': self.chambre_double,
            'Chambre_Familiale': self.chambre_familiale,
            'Lit_Supp': self.lit_supp,
            'Day_Use': self.day_use,
            'Vignette': self.vignette,
            'Taxe_Séjour': self.taxe_sejour,
            'Petit_Déjeuner': self.petit_dejeuner,
            'Déjeuner': self.dejeuner,
            'Dîner': self.diner,
            'Description': self.description,
            'Contact': self.contact,
            'Email': self.email
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
        hotel = cls()
        hotel.id = data.get('ID', '')
        hotel.nom = data.get('Nom', '')
        hotel.lieu = data.get('Lieu', '')
        hotel.type_hebergement = data.get('Type_Hébergement', '')
        hotel.categorie = data.get('Catégorie', '')
        hotel.chambre_single = data.get('Chambre_Single', 0) or 0
        hotel.chambre_double = data.get('Chambre_Double', 0) or 0
        hotel.chambre_familiale = data.get('Chambre_Familiale', 0) or 0
        hotel.lit_supp = data.get('Lit_Supp', 0) or 0
        hotel.day_use = data.get('Day_Use', 0) or 0
        hotel.vignette = data.get('Vignette', 0) or 0
        hotel.taxe_sejour = data.get('Taxe_Séjour', 0) or 0
        hotel.petit_dejeuner = data.get('Petit_Déjeuner', 0) or 0
        hotel.dejeuner = data.get('Déjeuner', 0) or 0
        hotel.diner = data.get('Dîner', 0) or 0
        hotel.description = data.get('Description', '')
        hotel.contact = data.get('Contact', '')
        hotel.email = data.get('Email', '')
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
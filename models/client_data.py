"""
Client data model
"""

from datetime import datetime
from config import *


class ClientData:
    """
    Represents client request data
    """

    def __init__(self):
        self.timestamp = ""
        self.ref_client = ""
        self.numero_dossier = ""
        self.type_client = ""  # Individuel ou Groupe
        self.prenom = ""
        self.nom = ""
        self.date_arrivee = ""
        self.date_depart = ""
        self.duree_sejour = ""
        self.nombre_participants = ""
        self.nombre_adultes = ""
        self.nombre_enfants_2_12 = ""
        self.nombre_bebes_0_2 = ""
        self.telephone = ""
        self.telephone_whatsapp = ""
        self.email = ""
        self.periode = ""
        self.restauration = ""
        self.hebergement = ""
        self.chambre = ""
        self.enfant = ""
        self.age_enfant = ""
        self.forfait = ""
        self.circuit = ""
        self.type_circuit = ""
        self.ville_depart = ""
        self.ville_arrivee = ""
        self.type_hotel_arrivee = ""
        # Rooming list
        self.sgl_count = ""  # Single
        self.dbl_count = ""  # Double
        self.twn_count = ""  # Twin
        self.tpl_count = ""  # Triple
        self.fml_count = ""  # Familiale

    def to_dict(self):
        """
        Convert to dictionary for Excel saving

        Returns:
            dict: Client data as dictionary
        """
        return {
            'Timestamp': self.timestamp,
            'Ref_Client': self.ref_client,
            'Numero_Dossier': self.numero_dossier,
            'Type_Client': self.type_client,
            'Prénom': self.prenom,
            'Nom': self.nom,
            'Date_Arrivée': self.date_arrivee,
            'Date_Départ': self.date_depart,
            'Durée_Séjour': self.duree_sejour,
            'Nombre_Participants': self.nombre_participants,
            'Nombre_Adultes': self.nombre_adultes,
            'Enfants_2_12_ans': self.nombre_enfants_2_12,
            'Bébés_0_2_ans': self.nombre_bebes_0_2,
            'Téléphone': self.telephone,
            'Téléphone_WhatsApp': self.telephone_whatsapp,
            'Email': self.email,
            'Période': self.periode,
            'Restauration': self.restauration,
            'Hébergement': self.hebergement,
            'Chambre': self.chambre,
            'Enfant': self.enfant,
            'Âge_Enfant': self.age_enfant,
            'Forfait': self.forfait,
            'Circuit': self.circuit,
            'Type_Circuit': self.type_circuit,
            'Ville_Depart': self.ville_depart,
            'Ville_Arrivee': self.ville_arrivee,
            'Type_Hotel_Arrivee': self.type_hotel_arrivee,
            'SGL_Count': self.sgl_count,
            'DBL_Count': self.dbl_count,
            'TWN_Count': self.twn_count,
            'TPL_Count': self.tpl_count,
            'FML_Count': self.fml_count
        }

    @classmethod
    def from_form_data(cls, form_data):
        """
        Create ClientData from form data dictionary

        Args:
            form_data (dict): Form data dictionary

        Returns:
            ClientData: ClientData instance
        """
        client = cls()
        client.timestamp = form_data.get('timestamp', datetime.now().strftime('%d/%m/%Y %H:%M'))
        client.ref_client = form_data.get('ref_client', '').strip()
        client.numero_dossier = form_data.get('numero_dossier', '').strip()
        client.type_client = form_data.get('type_client', '').strip()
        client.prenom = form_data.get('prenom', '').strip()
        client.nom = form_data.get('nom', '').strip()
        client.date_arrivee = form_data.get('date_arrivee', '').strip()
        client.date_depart = form_data.get('date_depart', '').strip()
        client.duree_sejour = form_data.get('duree_sejour', '').strip()
        client.nombre_participants = form_data.get('nombre_participants', '').strip()
        client.nombre_adultes = form_data.get('nombre_adultes', '').strip()
        client.nombre_enfants_2_12 = form_data.get('nombre_enfants_2_12', '').strip()
        client.nombre_bebes_0_2 = form_data.get('nombre_bebes_0_2', '').strip()
        client.telephone = form_data.get('telephone', '')
        client.telephone_whatsapp = form_data.get('telephone_whatsapp', '')
        client.email = form_data.get('email', '').strip()
        client.periode = form_data.get('periode', '')
        client.restauration = form_data.get('restauration', '')
        client.hebergement = form_data.get('hebergement', '')
        client.chambre = form_data.get('chambre', '')
        client.enfant = form_data.get('enfant', '')
        client.age_enfant = form_data.get('age_enfant', '')
        client.forfait = form_data.get('forfait', '')
        client.circuit = form_data.get('circuit', '')
        client.type_circuit = form_data.get('type_circuit', '')
        client.ville_depart = form_data.get('ville_depart', '')
        client.ville_arrivee = form_data.get('ville_arrivee', '')
        client.type_hotel_arrivee = form_data.get('type_hotel_arrivee', '')
        client.sgl_count = form_data.get('sgl_count', '').strip()
        client.dbl_count = form_data.get('dbl_count', '').strip()
        client.twn_count = form_data.get('twn_count', '').strip()
        client.tpl_count = form_data.get('tpl_count', '').strip()
        client.fml_count = form_data.get('fml_count', '').strip()
        return client

    def validate(self):
        """
        Validate client data

        Returns:
            list: List of validation errors (empty if valid)
        """
        errors = []

        if not self.ref_client:
            errors.append("Référence client obligatoire")
        if not self.type_client:
            errors.append("Type de client obligatoire")
        if not self.prenom:
            errors.append("Prénom obligatoire")
        if not self.nom:
            errors.append("Nom obligatoire")
        if not self.date_arrivee:
            errors.append("Date d'arrivée obligatoire")
        if not self.date_depart:
            errors.append("Date de départ obligatoire")
        if not self.nombre_participants:
            errors.append("Nombre de participants obligatoire")
        if not self.nombre_adultes:
            errors.append("Nombre d'adultes obligatoire")
        if not self.telephone:
            errors.append("Téléphone obligatoire")
        if not self.email:
            errors.append("Email obligatoire")
        if not self.periode:
            errors.append("Période obligatoire")
        if not self.forfait:
            errors.append("Forfait obligatoire")
        if not self.circuit:
            errors.append("Circuit obligatoire")

        return errors

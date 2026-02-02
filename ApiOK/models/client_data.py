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
        self.nom = ""
        self.telephone = ""
        self.email = ""
        self.periode = ""
        self.restauration = ""
        self.hebergement = ""
        self.chambre = ""
        self.enfant = ""
        self.age_enfant = ""
        self.forfait = ""
        self.circuit = ""

    def to_dict(self):
        """
        Convert to dictionary for Excel saving

        Returns:
            dict: Client data as dictionary
        """
        return {
            'Timestamp': self.timestamp,
            'Ref_Client': self.ref_client,
            'Nom': self.nom,
            'Téléphone': self.telephone,
            'Email': self.email,
            'Période': self.periode,
            'Restauration': self.restauration,
            'Hébergement': self.hebergement,
            'Chambre': self.chambre,
            'Enfant': self.enfant,
            'Âge_Enfant': self.age_enfant,
            'Forfait': self.forfait,
            'Circuit': self.circuit
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
        client.nom = form_data.get('nom', '').strip()
        client.telephone = form_data.get('telephone', '')
        client.email = form_data.get('email', '').strip()
        client.periode = form_data.get('periode', '')
        client.restauration = form_data.get('restauration', '')
        client.hebergement = form_data.get('hebergement', '')
        client.chambre = form_data.get('chambre', '')
        client.enfant = form_data.get('enfant', '')
        client.age_enfant = form_data.get('age_enfant', '')
        client.forfait = form_data.get('forfait', '')
        client.circuit = form_data.get('circuit', '')
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
        if not self.nom:
            errors.append("Nom obligatoire")
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
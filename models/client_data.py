"""
Client data model
"""

from datetime import datetime

# No config constants required in the model; keep models independent of UI config


class ClientData:
    """
    Represents client request data
    """

    def __init__(self):
        self.timestamp = ""
        self.ref_client = ""
        self.numero_dossier = ""
        self.statut = "En cours"
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
        self.accompagnement_guide = ""
        self.accompagnement_chauffeur = ""
        self.location_voiture = ""
        self.enfant = ""
        self.age_enfant = ""
        self.heure_arrivee = ""
        self.heure_depart = ""
        self.compagnie = ""
        self.aeroport = ""
        self.ext_ref = ""
        self.forfait = ""
        self.circuit = ""
        self.type_circuit = ""
        self.id_circuit = ""
        self.itineraire_circuit = ""
        self.activite_circuit = ""
        self.duree_circuit = ""
        self.condition_physique_circuit = ""
        self.type_voiture_circuit = ""
        self.hotels_defaut_villes_circuit = ""
        self.prestations_incluses_circuit = ""
        self.transports_associes_circuit = ""
        self.ville_depart = ""
        self.ville_arrivee = ""
        self.itineraire_detail = ""
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
            "Timestamp": self.timestamp,
            "Ref_Client": self.ref_client,
            "Numero_Dossier": self.numero_dossier,
            "Type_Client": self.type_client,
            "Prénom": self.prenom,
            "Nom": self.nom,
            "Date_Arrivée": self.date_arrivee,
            "Date_Départ": self.date_depart,
            "Durée_Séjour": self.duree_sejour,
            "Nombre_Participants": self.nombre_participants,
            "Nombre_Adultes": self.nombre_adultes,
            "Enfants_2_12_ans": self.nombre_enfants_2_12,
            "Bébés_0_2_ans": self.nombre_bebes_0_2,
            "Téléphone": self.telephone,
            "Téléphone_WhatsApp": self.telephone_whatsapp,
            "Email": self.email,
            "Période": self.periode,
            "Restauration": self.restauration,
            "Hébergement": self.hebergement,
            "Chambre": self.chambre,
            "Statut": self.statut,
            "Accompagnement_Guide": self.accompagnement_guide,
            "Accompagnement_Chauffeur": self.accompagnement_chauffeur,
            "Location_Voiture": self.location_voiture,
            "Enfant": self.enfant,
            "Âge_Enfant": self.age_enfant,
            "Heure_Arrivee": self.heure_arrivee,
            "Heure_Depart": self.heure_depart,
            "Compagnie": self.compagnie,
            "Aeroport": self.aeroport,
            "Ext_Ref": self.ext_ref,
            "Forfait": self.forfait,
            "Circuit": self.circuit,
            "Type_Circuit": self.type_circuit,
            "ID_Circuit": self.id_circuit,
            "Itineraire_Circuit": self.itineraire_circuit,
            "Activite_Circuit": self.activite_circuit,
            "Duree_Circuit": self.duree_circuit,
            "Condition_Physique_Circuit": self.condition_physique_circuit,
            "Type_Voiture_Circuit": self.type_voiture_circuit,
            "Hotels_Defaut_Villes_Circuit": self.hotels_defaut_villes_circuit,
            "Prestations_Incluses_Circuit": self.prestations_incluses_circuit,
            "Transports_Associes_Circuit": self.transports_associes_circuit,
            "Ville_Depart": self.ville_depart,
            "Ville_Arrivee": self.ville_arrivee,
            "Itineraire_Detail": self.itineraire_detail,
            "Type_Hotel_Arrivee": self.type_hotel_arrivee,
            "Statut": self.statut,
            "Heure_Arrivee": self.heure_arrivee,
            "Heure_Depart": self.heure_depart,
            "Compagnie": self.compagnie,
            "Aeroport": self.aeroport,
            "Ext_Ref": self.ext_ref,
            "SGL_Count": self.sgl_count,
            "DBL_Count": self.dbl_count,
            "TWN_Count": self.twn_count,
            "TPL_Count": self.tpl_count,
            "FML_Count": self.fml_count,
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
        client.timestamp = form_data.get(
            "timestamp", datetime.now().strftime("%d/%m/%Y %H:%M")
        )
        client.ref_client = form_data.get("ref_client", "").strip()
        client.numero_dossier = form_data.get("numero_dossier", "").strip()
        client.type_client = form_data.get("type_client", "").strip()
        client.prenom = form_data.get("prenom", "").strip()
        client.nom = form_data.get("nom", "").strip()
        client.date_arrivee = form_data.get("date_arrivee", "").strip()
        client.date_depart = form_data.get("date_depart", "").strip()
        client.duree_sejour = form_data.get("duree_sejour", "").strip()
        client.nombre_participants = form_data.get("nombre_participants", "").strip()
        client.nombre_adultes = form_data.get("nombre_adultes", "").strip()
        client.nombre_enfants_2_12 = form_data.get("nombre_enfants_2_12", "").strip()
        client.nombre_bebes_0_2 = form_data.get("nombre_bebes_0_2", "").strip()
        client.telephone = form_data.get("telephone", "")
        client.telephone_whatsapp = form_data.get("telephone_whatsapp", "")
        client.email = form_data.get("email", "").strip()
        client.periode = form_data.get("periode", "")
        client.restauration = form_data.get("restauration", "")
        client.hebergement = form_data.get("hebergement", "")
        client.chambre = form_data.get("chambre", "")
        client.statut = form_data.get("statut", client.statut)
        client.accompagnement_guide = form_data.get("accompagnement_guide", "")
        client.accompagnement_chauffeur = form_data.get("accompagnement_chauffeur", "")
        client.location_voiture = form_data.get("location_voiture", "")
        client.enfant = form_data.get("enfant", "")
        client.age_enfant = form_data.get("age_enfant", "")
        client.heure_arrivee = form_data.get("heure_arrivee", "").strip()
        client.heure_depart = form_data.get("heure_depart", "").strip()
        client.compagnie = form_data.get("compagnie", "").strip()
        client.aeroport = form_data.get("aeroport", "").strip()
        client.ext_ref = form_data.get("ext_ref", "").strip()
        client.forfait = form_data.get("forfait", "")
        client.circuit = form_data.get("circuit", "")
        client.type_circuit = form_data.get("type_circuit", "")
        client.id_circuit = form_data.get("id_circuit", "")
        client.itineraire_circuit = form_data.get("itineraire_circuit", "")
        client.activite_circuit = form_data.get("activite_circuit", "")
        client.duree_circuit = form_data.get("duree_circuit", "")
        client.condition_physique_circuit = form_data.get(
            "condition_physique_circuit", ""
        )
        client.type_voiture_circuit = form_data.get("type_voiture_circuit", "")
        client.hotels_defaut_villes_circuit = form_data.get(
            "hotels_defaut_villes_circuit", ""
        )
        client.prestations_incluses_circuit = form_data.get(
            "prestations_incluses_circuit", ""
        )
        client.transports_associes_circuit = form_data.get(
            "transports_associes_circuit", ""
        )
        client.ville_depart = form_data.get("ville_depart", "")
        client.ville_arrivee = form_data.get("ville_arrivee", "")
        client.itineraire_detail = form_data.get("itineraire_detail", "")
        client.type_hotel_arrivee = form_data.get("type_hotel_arrivee", "")
        client.sgl_count = form_data.get("sgl_count", "").strip()
        client.dbl_count = form_data.get("dbl_count", "").strip()
        client.twn_count = form_data.get("twn_count", "").strip()
        client.tpl_count = form_data.get("tpl_count", "").strip()
        client.fml_count = form_data.get("fml_count", "").strip()
        return client

    def validate(self):
        """
        Validate client data

        Returns:
            list: List of validation errors (empty if valid)
        """
        errors = []

        # ref_client est auto-généré si vide — pas d'erreur ici
        if not self.type_client:
            errors.append("Titre obligatoire")
        if self.type_client != "CIE" and not self.prenom:
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
        # forfait est optionnel (champ hérité)
        if not self.circuit:
            errors.append("Circuit obligatoire")

        return errors

# Feuille COTATION_H - Guide d'utilisation

## Vue d'ensemble

Une nouvelle feuille **COTATION_H** a été créée dans le fichier `data.xlsx` pour centraliser et regrouper toutes les cotations hôtel. Cette feuille permet un suivi consolidé des réservations d'hôtels par client et par ville.

## Données stockées

Pour chaque cotation hôtel, les informations suivantes sont enregistrées:

| Colonne | Description |
|---------|-------------|
| **Date** | Date et heure de création de la cotation |
| **ID_Client** | Identifiant/Référence du client |
| **Nom_Client** | Nom complet du client |
| **Hôtel** | Nom de l'hôtel |
| **Ville** | Ville de localisation de l'hôtel |
| **Nuits** | Nombre de nuits réservées |
| **Type_Chambre** | Type de chambre (Single, Double, Triple, Familiale) |
| **Adultes** | Nombre d'adultes |
| **Enfants** | Nombre d'enfants |
| **Plan_Repas** | Plan de restauration sélectionné |
| **Période** | Période/Saison (Haute, Moyenne, Basse) |
| **Total_Devise** | Montant total de la cotation |
| **Devise** | Devise du montant (Ariary, Euro, Dollar) |

## Fonctionnalités

### 1. Enregistrement automatique des cotations

Chaque fois qu'une cotation hôtel est générée (devis PDF créé), les données sont **automatiquement enregistrées** dans la feuille COTATION_H.

**Processus:**
1. Utilisateur crée une cotation dans "🏨 Cotation hôtel > 🆕 Nouvelle cotation"
2. Utilisateur clique sur "📄 Générer devis"
3. Le PDF est généré ET les données sont enregistrées dans COTATION_H

### 2. Affichage groupé par client

Via le menu "🏨 Cotation hôtel > 📊 Résumé cotations":

**Par Client:**
- Chaque client est affiché avec ses réservations d'hôtels
- Les hôtels réservés peuvent être dans différentes villes
- Un **sous-total** est calculé pour chaque client
- Un **total général** est affiché en haut

**Structure:**
```
┌─────────────────────────────────────┐
│ TOTAL GÉNÉRAL: XX,XXX.XX Ar         │
├─────────────────────────────────────┤
│ Client: John Doe (ID: CLI001)       │
│  - Hôtel Zanzibar - Antananarivo    │
│  - Hôtel Andromeda - Nosy Be        │
│ Sous-total: XX,XXX.XX Ar            │
│                                     │
│ Client: Jane Smith (ID: CLI002)     │
│  - Hôtel Silberrand - Antalaha      │
│ Sous-total: XX,XXX.XX Ar            │
└─────────────────────────────────────┘
```

### 3. Affichage groupé par ville

Permet de voir le total des réservations par destination:

**Par Ville:**
- Chaque ville affiche tous les hôtels réservés
- Les clients sont listés pour chaque hôtel
- Un **sous-total** est calculé pour chaque ville
- Un **total général** pour toutes les villes

**Cas d'usage:**
- Identifier les villes les plus demandées
- Analyser les dépenses par destination
- Négocier avec les hôtels en bloc

### 4. Vue détaillée

Tableau affichant pour chaque cotation:
- Détails du client et de l'hôtel
- Paramètres du séjour (nuits, adultes, enfants)
- Plan de restauration
- Montant total avec devise

## Utilisation

### Accéder au résumé des cotations

1. Cliquez sur **"🏨 Cotation hôtel"** dans le menu latéral
2. Sélectionnez **"📊 Résumé cotations"**
3. Choisissez le mode d'affichage:
   - **Par client** - pour analyser les dépenses par client
   - **Par ville** - pour analyser les dépenses par destination

### Rafraîchir les données

Cliquez sur **"🔄 Rafraîchir"** pour recharger les données depuis Excel et voir les dernières cotations ajoutées.

## Exemples de données

### Client avec plusieurs réservations dans différentes villes

```
┌─────────────────────────────────────────────────────┐
│ Client: Marie Dupont (ID: CLI003)                   │
├─────────────────────────────────────────────────────┤
│ Hôtel Sakamanga - Antananarivo                      │
│ - 3 nuits | Double | 2 adultes | Demi-pension      │
│ - Total: 150,000 Ar                                 │
│                                                     │
│ Hôtel Sakalava - Nosy Be                            │
│ - 5 nuits | Triple | 3 adultes, 1 enfant           │
│ - Total: 250,000 Ar                                 │
│                                                     │
│ SOUS-TOTAL: 400,000 Ar                              │
└─────────────────────────────────────────────────────┘
```

### Groupement par ville

```
┌─────────────────────────────────────┐
│ Ville: Nosy Be                      │
├─────────────────────────────────────┤
│ Hôtel Sakalava                      │
│ - Client: Marie Dupont              │
│ - Total: 250,000 Ar                 │
│                                     │
│ Hôtel Andromeda                     │
│ - Client: John Doe                  │
│ - Total: 175,000 Ar                 │
│                                     │
│ SOUS-TOTAL: 425,000 Ar              │
└─────────────────────────────────────┘
```

## Fichiers modifiés

### 1. `config.py`
- Ajout: `COTATION_H_SHEET_NAME = "COTATION_H"`

### 2. `utils/excel_handler.py`
Nouvelles fonctions:
- `save_hotel_quotation_to_excel(quotation_data)` - Enregistre une cotation
- `load_all_hotel_quotations()` - Charge toutes les cotations
- `get_quotations_grouped_by_client()` - Groupe par client avec totaux
- `get_quotations_by_city()` - Groupe par ville avec totaux

### 3. `gui/forms/hotel_quotation.py`
- Import: `save_hotel_quotation_to_excel`
- Ajout: Enregistrement automatique lors de la génération du PDF

### 4. `gui/forms/hotel_quotation_summary.py` (NOUVEAU)
- Nouvelle classe: `HotelQuotationSummary`
- Affichage groupé par client ou par ville
- Calcul des sous-totaux et total général

### 5. `gui/sidebar.py`
- Menu "Cotation hôtel" converti en sous-menu
- Ajout: "📊 Résumé cotations"
- Ajout: Callback `_show_hotel_quotation_summary()`

### 6. `gui/main_content.py`
- Ajout: Gestion du type "hotel_quotation_summary"
- Ajout: Méthode `_show_hotel_quotation_summary()`

## Notes techniques

### Structure de la feuille COTATION_H

- **Ligne 1**: En-têtes (formatés en gras avec fond bleu)
- **À partir de la ligne 2**: Données des cotations
- **Colonnes**: A à M (13 colonnes)
- **Format automatique**: Les largeurs de colonnes sont ajustées automatiquement

### Conservation des données

Les cotations ne sont **jamais supprimées** - elles forment un historique permanent. Pour une gestion complète, vous pouvez:
- Archiver les anciennes cotations dans une feuille séparée
- Créer des filtres temporels dans Excel
- Exporter les données pour analyses externes

### Devise

Les totaux sont enregistrés avec leur devise d'origine (Ariary, Euro, Dollar). Les regroupements respectent la devise de chaque cotation.

## Améliorations futures possibles

- [ ] Filtres temporels (par date)
- [ ] Export vers formats externes (CSV, PDF)
- [ ] Graphiques de synthèse
- [ ] Alertes de prix pour clients récurrents
- [ ] Historique de modifications
- [ ] Comparaison de prix entre hôtels
- [ ] Statistiques par période/saison

## Support

Pour toute question ou problème:
1. Vérifiez que `data.xlsx` existe et est accessible
2. Vérifiez que openpyxl est installé: `pip install openpyxl`
3. Consultez les logs de l'application dans le fichier `app.log`

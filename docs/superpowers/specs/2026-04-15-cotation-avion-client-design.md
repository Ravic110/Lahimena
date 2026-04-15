# Design: Cotation Avion Client

## Contexte

Le projet dispose deja :

- d'un module avion generique (`air_ticket_page`, `air_ticket_quotation`, `air_ticket_quotation_summary`)
- de cotations client dediees pour l'hotel, les frais collectifs et le transport

Le besoin est d'ajouter une **cotation avion rattachee a un client precis**, sur le meme modele d'usage que la cotation transport recente :

- acces direct depuis la page d'accueil sur un client
- chargement par client
- sauvegarde par client
- reouverture avec restitution des lignes precedemment sauvegardees

Le besoin inclut aussi une **marge en pourcentage** sur chaque ligne, avec une logique alignee sur la cotation hotel.

## Objectif

Permettre a l'utilisateur de gerer une cotation avion pour un client donne via un ecran dedie, pre-rempli depuis la fiche client, capable de generer automatiquement un trajet `aller simple` ou `aller-retour`, de calculer les montants par ligne avec marge (%) et de sauvegarder ces lignes dans une feuille Excel dediee remplacant les anciennes lignes du meme client.

## Hors scope

- refonte du module avion generique existant
- factorisation profonde entre l'ecran client et l'ecran avion generique
- ajout de logique complexe de grille tarifaire ou de base de donnees tarifaire supplementaire
- changement du point d'entree actuel du module avion generique

## Approche retenue

L'approche retenue est la creation d'un **ecran client dedie minimal**, separe du module avion generique existant.

Cette approche est retenue parce qu'elle :

- suit le pattern deja etabli par `client_hotel_cotation` et `client_transport_cotation`
- limite le risque de regression sur l'ecran avion generique actuel
- permet d'introduire rapidement la logique client, la generation aller/retour et la marge (%)

## Parcours utilisateur

### Acces

Depuis la page d'accueil, chaque client doit disposer d'une action directe `Cotation avion`, au meme niveau que les actions existantes de cotation hotel et cotation transport.

Cette action ouvre directement un nouvel ecran de cotation avion pour le client selectionne.

### Chargement initial

Au chargement de l'ecran :

1. l'application cherche d'abord des lignes de cotation avion deja sauvegardees pour ce client
2. si elles existent, elles sont rechargees telles quelles
3. sinon, une cotation initiale est generee depuis la fiche client

### Generation initiale

Le mode de generation doit supporter trois cas metier :

- `aller simple`
- `aller-retour`
- edition libre apres generation

Generation par defaut :

- en mode `aller simple`, l'ecran cree une ligne `aller`
- en mode `aller-retour`, l'ecran cree deux lignes :
  - `aller`: `ville_depart client -> ville_arrivee client`
  - `retour`: `ville_arrivee client -> ville_depart client`

Les lignes generees automatiquement restent entierement modifiables.

### Edition

L'utilisateur peut :

- modifier une ligne existante
- supprimer une ligne existante
- ajouter une ligne manuelle
- basculer un total calcule vers un total manuel

L'ecran reste un tableau oriente lignes, comme les autres cotations client.

## Donnees affichees

### En-tete client

L'ecran affiche en tete les informations client utiles au contexte :

- nom et prenom
- numero de dossier
- nombre de participants si disponible
- nombre d'adultes
- nombre d'enfants
- ville de depart
- ville d'arrivee

### Mode de trajet

L'ecran propose un choix de mode :

- `aller simple`
- `aller-retour`

Ce mode sert a la generation initiale et/ou a la regeneration d'une base de lignes pour un client qui n'a pas encore de cotation sauvegardee.

### Structure de ligne

Chaque ligne de cotation avion contient les champs suivants :

- `type_trajet` : `aller` ou `retour`
- `compagnie`
- `ville_depart`
- `ville_arrivee`
- `nb_adultes`
- `nb_enfants`
- `tarif_adulte`
- `tarif_enfant`
- `montant_adultes`
- `montant_enfants`
- `sous_total`
- `marge_pct`
- `total`
- `total_manuel` (indicateur persiste, pour savoir si le total a ete force par l'utilisateur)

`total_manuel` doit etre persiste avec la ligne afin de restituer exactement le comportement apres rechargement. Le stockage doit donc conserver a la fois la valeur de `total` et l'indicateur `total_manuel`.

### Pre-remplissage

Pour une cotation initiale non encore sauvegardee :

- `compagnie` est pre-remplie depuis la fiche client si l'information existe, sinon reste vide
- `ville_depart` est pre-remplie depuis la fiche client
- `ville_arrivee` est pre-remplie depuis la fiche client
- `nb_adultes` est pre-rempli depuis la fiche client
- `nb_enfants` est pre-rempli depuis la fiche client
- `marge_pct` est vide ou `0`
- `tarif_adulte` et `tarif_enfant` sont vides

Les montants derives sont recalcules automatiquement des qu'une valeur utile change.

## Regles de calcul

### Calcul automatique par ligne

Pour une ligne en mode calcule :

- `montant_adultes = nb_adultes * tarif_adulte`
- `montant_enfants = nb_enfants * tarif_enfant`
- `sous_total = montant_adultes + montant_enfants`
- `montant_marge = sous_total * marge_pct / 100`
- `total = sous_total + montant_marge`

La logique de marge doit etre alignee sur la cotation hotel : la marge est portee par ligne en pourcentage, puis integree au total de la ligne.

### Total manuel

Le comportement retenu est un mode mixte :

- le calcul automatique est actif par defaut
- l'utilisateur peut saisir manuellement un `total`

Regles :

- si l'utilisateur renseigne manuellement un total valide, ce total prend priorite pour la ligne
- les autres champs calcules (`montant_adultes`, `montant_enfants`, `sous_total`) restent affiches comme reference
- si l'utilisateur vide le total manuel, la ligne revient en mode de calcul automatique

### Totaux d'ecran

L'ecran affiche un total global correspondant a la somme des `total` de toutes les lignes, qu'ils soient calcules ou manuels.

## Validation

Une ligne est sauvegardable seulement si :

- `ville_depart` est renseignee
- `ville_arrivee` est renseignee
- `nb_adultes` est numerique et >= 0
- `nb_enfants` est numerique et >= 0
- `tarif_adulte` est numerique et >= 0
- `tarif_enfant` est numerique et >= 0
- `marge_pct` est numerique et >= 0
- `total`, s'il est saisi manuellement, est numerique et >= 0

Si aucune ligne valide n'existe, la sauvegarde doit etre refusee avec un message clair.

## Persistance

### Feuille Excel

La cotation avion client doit etre sauvegardee dans une **feuille Excel dediee** de `data.xlsx`, sur le modele des autres cotations client.

Le nom exact de la feuille doit etre explicite et coherent avec les conventions existantes. La reference retenue pour l'implementation est `COTATION_AVION`.

### Strategie de sauvegarde

La sauvegarde suit le modele des autres cotations client :

1. identifier le client par sa reference
2. supprimer dans la feuille dediee toutes les lignes existantes de ce client
3. ecrire l'etat courant de la cotation avion du client

Les donnees minimales a stocker doivent inclure :

- metadata client : date, identifiant client, numero dossier, nom, prenom
- contenu metier ligne : type trajet, compagnie, villes, pax, tarifs, montants, sous-total, marge %, total, total_manuel

### Chargement

Au rechargement, l'application lit toutes les lignes de la feuille dediee correspondant au client courant et reconstruit les lignes de l'ecran dans le meme ordre fonctionnel.

Si aucune ligne n'est trouvee, l'application retombe sur la generation initiale depuis la fiche client.

## Navigation et integration

Les integrations minimales attendues sont :

- ajout d'un nouvel ecran `client_air_ticket_cotation`
- ajout d'un handler correspondant dans `gui/main_content.py`
- ajout d'une action `Cotation avion` dans la page d'accueil, au meme niveau que les actions hotel et transport

Le retour utilisateur peut suivre le pattern deja etabli des autres cotations client, avec retour a l'accueil.

## Gestion d'erreurs

Les erreurs a traiter explicitement :

- fichier Excel verrouille/ouvert ailleurs
- absence de feuille cible, avec creation si necessaire
- absence de reference client exploitable
- erreurs de conversion numerique sur les champs modifiables

Le comportement attendu reste coherent avec les autres cotations :

- message utilisateur clair
- logs applicatifs conserves
- pas de crash de l'interface

## Impacts architecture

Les composants probables sont :

- un nouvel ecran GUI dedie de type `client_air_ticket_cotation.py`
- des fonctions de chargement/sauvegarde dans `utils/excel_handler.py`
- un branchement de navigation dans `gui/main_content.py`
- un ajout d'action dans `gui/forms/home_page.py`

Le module avion generique existant doit rester operationnel et independent.

## Strategie de tests

Les tests d'implementation devront couvrir au minimum :

- chargement d'une cotation avion client existante
- generation initiale `aller simple`
- generation initiale `aller-retour`
- calcul automatique :
  - montants adultes
  - montants enfants
  - sous-total
  - marge
  - total
- conservation d'un `total` manuel
- retour au calcul automatique si le total manuel est vide
- sauvegarde remplacant les anciennes lignes du client
- navigation depuis l'accueil vers la cotation avion client

## Risques et points de vigilance

- ambiguite de persistance autour du `total manuel` si aucun marqueur n'est stocke
- divergence possible entre l'ecran client et le module avion generique si une logique commune evolue plus tard
- risque de duplication acceptable a court terme, mais a surveiller si le perimetre avion grossit

## Decision finale

Le projet ajoute une **cotation avion par client** via un **nouvel ecran dedie**, accessible directement depuis l'accueil, avec :

- pre-remplissage depuis la fiche client
- generation `aller simple` ou `aller-retour`
- calcul automatique des montants
- marge (%) par ligne
- total manuel possible
- sauvegarde/rechargement par client dans une feuille Excel dediee

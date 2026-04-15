# Design: correction du calcul de distance transport client et nettoyage des données

Date: 2026-04-14
Projet: Lahimena Tours
Statut: validé en discussion, en attente de relecture utilisateur avant plan d'implémentation

## Contexte

Le parcours `double clic client -> cotation transport` n'affiche pas correctement la distance entre villes.

L'analyse du code et des fichiers Excel du projet a montré quatre causes combinées:

1. `gui/forms/client_transport_cotation.py` utilise le KM de la ville d'arrivée au lieu de calculer `abs(KM arrivée - KM départ)`.
2. Les données clients existantes contiennent des noms de villes non normalisés, par exemple `Antananarivo(1 jours)`, `Ranohira (Isalo)(2 jours))`, `Tuler`.
3. Le lookup `KM_MADA` repose sur un index fragile qui écrase les doublons de repères.
4. Certains repères utiles sont absents de `KM_MADA`; ce point n'est pas traité structurellement par cette intervention.

L'objectif de cette intervention est de corriger le comportement à l'exécution et de nettoyer les données existantes, sans refondre le modèle transport global.

## Objectifs

- Corriger le calcul de distance dans la cotation transport client.
- Centraliser une normalisation métier des villes/répères.
- Fiabiliser la recherche dans `KM_MADA` en présence de doublons.
- Nettoyer les données existantes de `INFOS_CLIENTS`.
- Ajouter une couverture de tests sur les cas réels observés.

## Hors périmètre

- Refonte complète du modèle `KM_MADA` par axe/segment.
- Ajout exhaustif de tous les repères manquants dans `data-hotel.xlsx`.
- Modification du parcours principal de cotation transport hors besoin de cohérence de calcul.

## Approche retenue

Approche hybride:

- correction runtime dans les modules de calcul;
- normalisation métier centralisée;
- migration légère des données existantes;
- conservation de la structure Excel actuelle.

Cette approche corrige le bug utilisateur immédiatement tout en réduisant le risque de régression sur les dossiers déjà saisis.

## Design technique

### 1. Normalisation centralisée des repères

Une fonction unique de normalisation sera ajoutée dans `utils/excel_handler.py` et utilisée par tous les accès `KM_MADA`.

Règles prévues:

- trim des espaces, homogénéisation des espaces internes;
- suppression des suffixes de durée saisis dans les villes, par exemple `(1 jours)`, `(2 jours)`, `(3 jour)`;
- suppression contrôlée d'annotations parasites;
- normalisation casse/accents;
- mapping d'alias métier connu.

Exemples d'alias métier pris en charge:

- `Tuler` -> `Toliary`
- `Tulear` -> `Toliary`
- `Ranohira (Isalo)` -> `Ranohira`

Cette normalisation ne doit pas être dispersée dans plusieurs écrans.

### 2. Lookup `KM_MADA` robuste

Le cache `_KM_MADA_CACHE["lookup"]` ne stockera plus une seule ligne par repère brut.

Comportement cible:

- indexer par nom normalisé;
- conserver toutes les lignes candidates pour un même repère;
- sélectionner une ligne exploitable avec une règle déterministe.

Règle de priorité:

1. préférer les lignes avec `km > 0`;
2. parmi elles, préférer la plus grande valeur de `km`;
3. sinon, retomber sur la première ligne disponible.

Cette règle vise à éviter les faux zéros provoqués par les lignes `ORIGINE` ou les repères dupliqués sur plusieurs axes.

### 3. Correction du calcul dans la cotation transport client

`gui/forms/client_transport_cotation.py` sera aligné sur la logique déjà utilisée dans `gui/forms/transport_quotation.py`.

Comportement cible:

- lors de la génération des segments, nettoyer les villes avant lookup;
- calculer la distance avec `abs(KM(arrivee) - KM(depart))`;
- appliquer cette logique:
  - au chargement initial de la table;
  - dans la fenêtre d'édition d'un trajet.

Résultat attendu:

- la colonne trajet affiche bien une distance cohérente;
- le champ KM utilisé pour le carburant est cohérent avec la distance du segment, et non avec le KM cumulé de l'arrivée seule.

### 4. Nettoyage des données existantes

Une migration légère sera ajoutée pour nettoyer `INFOS_CLIENTS`.

Colonnes ciblées:

- `Ville Départ`
- `Ville Arrivée`
- `Itinéraire Circuit`

Actions:

- réécriture des valeurs normalisées;
- conservation du format multi-ville déjà présent, mais avec repères nettoyés;
- journalisation des cellules modifiées pour audit.

Le nettoyage de `KM_MADA` sera limité aux variantes triviales si nécessaire. La structure de la feuille ne sera pas modifiée.

## Flux de données cible

1. L'écran client transport reçoit un client issu de `INFOS_CLIENTS`.
2. Les villes du client sont normalisées avant toute recherche.
3. Le lookup `KM_MADA` résout un repère via l'index robuste.
4. La distance du segment est calculée à partir des deux extrémités.
5. La table et les calculs carburant affichent des valeurs cohérentes.

## Gestion des erreurs

- Si un repère reste introuvable après normalisation, le calcul retourne `0` comme aujourd'hui.
- Le comportement silencieux ne sera pas durci dans cette itération, sauf si nécessaire pour préserver la cohérence UI existante.
- Les changements de données via migration doivent être tracés dans les logs applicatifs.

## Tests

Des tests ciblés seront ajoutés pour couvrir:

- normalisation des villes avec suffixes de durée;
- alias métier;
- lookup `KM_MADA` avec doublons;
- calcul de distance client transport entre départ et arrivée;
- cas réels observés dans les données du projet.

Exemples de cas à couvrir:

- `Antananarivo(1 jours)` -> `Antananarivo`
- `Antsirabe(2 jours)` -> `Antsirabe`
- `Tuler` -> `Toliary`
- `Ranohira (Isalo)` -> `Ranohira`
- `MORONDAVA` ne doit plus retourner `0` si une ligne destination avec KM utile existe

## Risques et compromis

- La normalisation métier peut transformer des valeurs ambiguës; il faut limiter les alias à des cas observés et sûrs.
- Le choix "plus grand KM positif" sur un repère dupliqué est une heuristique; elle corrige les faux zéros sans garantir une modélisation parfaite par axe.
- La migration de données existantes doit rester idempotente et auditée.

## Plan d'implémentation envisagé

1. Ajouter les tests qui reproduisent les cas cassés.
2. Introduire la normalisation centralisée et le nouveau lookup `KM_MADA`.
3. Corriger `client_transport_cotation.py`.
4. Ajouter la migration de nettoyage des données existantes.
5. Exécuter les tests ciblés puis la suite de vérification pertinente.

## Critères de succès

- Après double-clic sur un client puis ouverture de la cotation transport, la distance s'affiche pour les cas couverts par `KM_MADA`.
- Les dossiers existants contenant des suffixes type `(1 jours)` sont récupérés correctement.
- Les repères dupliqués de `KM_MADA` ne produisent plus de faux zéros évidents.
- Les tests ajoutés passent.

# Exemples d'utilisation - COTATION_H

## Scénario 1: Client avec plusieurs hôtels dans différentes villes

### Situation
Le client "John Doe" (CLI001) effectue 3 réservations:
- Hôtel Sakamanga à Antananarivo (3 nuits)
- Hôtel Sakalava à Nosy Be (5 nuits)
- Hôtel Andromeda à Nosy Be (2 nuits)

### Processus

#### Étape 1: Créer la 1ère cotation
1. Menu → "🏨 Cotation hôtel" → "🆕 Nouvelle cotation"
2. Sélectionner client: "CLI001 - John Doe"
3. Sélectionner: Hôtel Sakamanga (Antananarivo)
4. Paramètres: 3 nuits, Double, 2 adultes, Demi-pension
5. Calculer → Générer devis

**Résultat:** Données enregistrées dans COTATION_H ligne 2
```
Date: 2026-02-06 14:30:00
ID_Client: CLI001
Nom_Client: John Doe
Hôtel: Sakamanga
Ville: Antananarivo
Nuits: 3
Type_Chambre: Double/twin
Adultes: 2
Enfants: 0
Plan_Repas: Demi-pension
Période: Haute saison
Total_Devise: 150000.00
Devise: Ariary
```

#### Étape 2: Créer la 2e cotation
1. Menu → "🏨 Cotation hôtel" → "🆕 Nouvelle cotation"
2. Même client: "CLI001 - John Doe"
3. Sélectionner: Hôtel Sakalava (Nosy Be)
4. Paramètres: 5 nuits, Triple, 2 adultes, 1 enfant, Pension complète
5. Calculer → Générer devis

**Résultat:** Enregistrement en ligne 3

#### Étape 3: Créer la 3e cotation
1. Même processus
2. Hôtel Andromeda (Nosy Be)
3. Paramètres: 2 nuits, Double, 2 adultes

**Résultat:** Enregistrement en ligne 4

### Affichage groupé par client

Menu → "🏨 Cotation hôtel" → "📊 Résumé cotations" → Sélectionner "Par client"

**Affichage:**
```
┌──────────────────────────────────────────────────────┐
│           TOTAL GÉNÉRAL: 600,000.00 Ar              │
├──────────────────────────────────────────────────────┤
│                                                      │
│ ┌────────────────────────────────────────────────┐  │
│ │ Client: John Doe (ID: CLI001)                  │  │
│ ├────────────────────────────────────────────────┤  │
│ │ Hôtel           │ Ville          │ Total      │  │
│ ├─────────────────┼────────────────┼────────────┤  │
│ │ Sakamanga       │ Antananarivo   │ 150,000.00 │  │
│ │ Sakalava        │ Nosy Be        │ 300,000.00 │  │
│ │ Andromeda       │ Nosy Be        │ 150,000.00 │  │
│ ├──────────────────────────────────────────────┤  │
│ │ Sous-total John Doe: 600,000.00 Ar          │  │
│ └────────────────────────────────────────────────┘  │
│                                                      │
└──────────────────────────────────────────────────────┘
```

**Analyse:**
- Le client a 3 réservations dans 2 villes différentes
- Montant total de 600,000 Ar
- 2 réservations au même hôtel (Nosy Be)

---

## Scénario 2: Analyse par ville pour négociation

### Situation
Vous avez plusieurs clients réservant dans les mêmes villes et souhaitez analyser par destination.

### Processus

#### Créer plusieurs cotations (clients différents, même ville)

**Client 1: Marie Dupont (CLI002)**
- Hôtel: Sakamanga (Antananarivo) - 350,000 Ar

**Client 2: Pierre Martin (CLI003)**
- Hôtel: Sakamanga (Antananarivo) - 280,000 Ar

**Client 3: Sophie Leclerc (CLI004)**
- Hôtel: Sakalava (Nosy Be) - 420,000 Ar

### Affichage groupé par ville

Menu → "🏨 Cotation hôtel" → "📊 Résumé cotations" → Sélectionner "Par ville"

**Affichage:**
```
┌──────────────────────────────────────────────────────────┐
│           TOTAL GÉNÉRAL: 1,050,000.00 Ar                │
├──────────────────────────────────────────────────────────┤
│                                                          │
│ ┌────────────────────────────────────────────────────┐  │
│ │ Ville: Antananarivo                                │  │
│ ├─────────────────────────────────────────────────┤  │
│ │ Hôtel      │ Client         │ Nuits   │ Total   │  │
│ ├────────────┼────────────────┼─────────┼─────────┤  │
│ │ Sakamanga  │ Marie Dupont   │ 3       │ 350,000 │  │
│ │ Sakamanga  │ Pierre Martin  │ 2       │ 280,000 │  │
│ ├──────────────────────────────────────────────────┤  │
│ │ Sous-total Antananarivo: 630,000.00 Ar          │  │
│ └────────────────────────────────────────────────────┘  │
│                                                          │
│ ┌────────────────────────────────────────────────────┐  │
│ │ Ville: Nosy Be                                     │  │
│ ├─────────────────────────────────────────────────┤  │
│ │ Hôtel      │ Client         │ Nuits   │ Total   │  │
│ ├────────────┼────────────────┼─────────┼─────────┤  │
│ │ Sakalava   │ Sophie Leclerc │ 5       │ 420,000 │  │
│ ├──────────────────────────────────────────────────┤  │
│ │ Sous-total Nosy Be: 420,000.00 Ar               │  │
│ └────────────────────────────────────────────────────┘  │
│                                                          │
└──────────────────────────────────────────────────────────┘
```

**Analyse pour négociation:**
- **Antananarivo (Sakamanga):** 2 clients, total 630,000 Ar
  → Excellente opportunité de négociation de bloc!
- **Nosy Be (Sakalava):** 1 client, 420,000 Ar

---

## Scénario 3: Rapport mixte (Devise multiple)

### Situation
Des clients de différents pays avec devises différentes:
- Clients locaux: Ariary
- Clients expatriés: Euro
- Clients internationaux: Dollar

### Exemple de données dans COTATION_H

```
Ligne 2:
ID_Client: CLI001 | Hôtel: Zanzibar | Ville: Antananarivo
Total_Devise: 150000 | Devise: Ariary

Ligne 3:
ID_Client: CLI005 | Hôtel: Sakamanga | Ville: Antananarivo
Total_Devise: 2500 | Devise: Euro

Ligne 4:
ID_Client: CLI006 | Hôtel: Sakalava | Ville: Nosy Be
Total_Devise: 3000 | Devise: Dollar
```

### Affichage

**Par client:**
```
Client: Alice Johnson (ID: CLI005)
- Sakamanga - Antananarivo - 2,500.00 €
Sous-total: 2,500.00 €

Client: Bob Williams (ID: CLI006)
- Sakalava - Nosy Be - 3,000.00 $
Sous-total: 3,000.00 $
```

**Note:** Chaque client conserve sa devise d'enregistrement.

---

## Scénario 4: Suivi sur le temps

### Janvier 2026
```
Cotations créées: 5
Montant total: 2,500,000 Ar
Clients uniques: 3
Villes: Antananarivo, Nosy Be
```

### Février 2026 (après ajout de 3 nouvelle cotations)
```
Cotations créées: 8 (5 + 3 nouvelles)
Montant total: 3,800,000 Ar (2,500,000 + 1,300,000)
Clients uniques: 4 (ancien + 1 nouveau)
Villes: 3 (Antananarivo, Nosy Be + 1 nouvelle)
```

### Utilisation pratique

1. **Prévisions mensuelles:** Affichage par date de création
2. **ROI par client:** Affichage par client
3. **Capacité hôtels:** Affichage par ville + hôtel
4. **Tendances saisonnières:** Filtrer par "Période"

---

## Scénario 5: Comparaison de chambres

### Question
Quel type de chambre génère le plus de revenus?

### Données
```
Ligne 2: CLI001 | Sakamanga | Double | 3 nuits | 150,000 Ar
Ligne 3: CLI002 | Sakamanga | Single | 2 nuits | 80,000 Ar
Ligne 4: CLI003 | Sakamanga | Familiale | 4 nuits | 280,000 Ar
```

### Analyse manuelle dans Excel
- **Double:** 150,000 Ar (3 nuits) = 50,000 par nuit
- **Single:** 80,000 Ar (2 nuits) = 40,000 par nuit
- **Familiale:** 280,000 Ar (4 nuits) = 70,000 par nuit

**Conclusion:** Les chambres familiales sont les plus profitables!

---

## Cas d'usage avancé: Tableau de bord personnalisé

### Créer un espace de travail dans Excel

**Onglet: COTATION_H** (généré automatiquement)
- Données brutes

**Onglet: ANALYSE** (créé manuellement)
```
=SUMIFS(COTATION_H!L:L, COTATION_H!B:B, "CLI001")
→ Total pour client CLI001

=SUMIF(COTATION_H!E:E, "Nosy Be", COTATION_H!L:L)
→ Total pour ville "Nosy Be"

=COUNTIFS(COTATION_H!B:B, "CLI002", COTATION_H!E:E, "Antananarivo")
→ Nombre de réservations du client CLI002 à Antananarivo
```

---

## Workflow complet: De la saisie à l'analyse

```
1. SAISIE
   └─→ Menu "Cotation hôtel" → "Nouvelle cotation"
       └─→ Remplir formulaire
           └─→ Calculer
               └─→ Générer devis
                   └─→ Données sauvegardées dans COTATION_H

2. ENREGISTREMENT
   └─→ Excel data.xlsx
       └─→ Feuille: COTATION_H
           └─→ Ligne 2, 3, 4, ... (nouvelles cotations)

3. CONSULTATION
   └─→ Menu "Cotation hôtel" → "Résumé cotations"
       ├─→ Vue par CLIENT
       │   └─→ Voir total par client + détails hôtels
       └─→ Vue par VILLE
           └─→ Voir total par ville + détails clients

4. ANALYSE AVANCÉE
   └─→ Ouvrir Excel directement
       └─→ Utiliser les formules SUMIF, COUNTIFS...
           └─→ Créer tableaux de bord personnalisés
               └─→ Générer rapports
```

---

## Tips et bonnes pratiques

### ✅ À faire

1. **Cotations cohérentes**
   - Enregistrer les devis générés = saisie automatique dans COTATION_H
   - Toutes les réservations ont une trace

2. **Catégories claires**
   - Utiliser des références client cohérentes (CLI001, CLI002...)
   - Noms d'hôtels sans variation d'orthographe

3. **Suivi temporel**
   - La date est enregistrée automatiquement
   - Permet de retracer l'historique complet

4. **Vérifications régulières**
   - Rafraîchir les données (🔄) régulièrement
   - Vérifier la cohérence entre devis et cotations

### ❌ À éviter

1. **Ne pas modifier directement** les données dans COTATION_H
   - Risque de corruption
   - Perte de traçabilité
   - Recréer un devis si modification nécessaire

2. **Ne pas supprimer des lignes** manuellement
   - Utiliser l'interface (à développer)
   - Conserver l'historique

3. **Ne pas changer les en-têtes** des colonnes
   - L'application les attend fixes
   - Risque de dysfonctionnement

### 💡 Astuces

1. **Export régulier**
   ```
   Clic droit sur COTATION_H → Copier → Coller dans nouveau fichier
   ```

2. **Tri dans Excel**
   - Sélectionner données → Données → Trier
   - Par client, par ville, par date...

3. **Graphiques**
   - Excel → Insérer → Graphique
   - Visualiser les tendances

4. **Filtre auto**
   - Sélectionner en-têtes → Données → Filtre automatique
   - Filtrer par période, par devise...

---

## Dépannage

### Problème: "Aucune cotation trouvée"
**Cause:** Pas encore créé de devis  
**Solution:** Créer une cotation (Nouvelle cotation → Générer devis)

### Problème: Affichage blanc après rafraîchir
**Cause:** Erreur de chargement  
**Solution:**
1. Vérifier que data.xlsx existe
2. Vérifier que openpyxl est installé: `pip install openpyxl`
3. Consulter les logs d'erreur

### Problème: Les données n'apparaissent pas
**Cause:** COTATION_H n'existe pas ou est vide  
**Solution:**
1. Créer une nouvelle cotation et générer un devis
2. Rafraîchir l'affichage (bouton 🔄)

### Problème: Devise incohérente
**Cause:** Cotations créées avec devises différentes  
**Solution:** C'est normal! Chaque cotation garde sa devise

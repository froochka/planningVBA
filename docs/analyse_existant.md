# Analyse de l'existant : outil SGO d'origine (DCO New Age)

*6 octobre 2026. Sources analysées :*
- `02_PhaseElaborationOffre_DossierConception_V4_07072026.xlsm` : classeur d'origine, 38 feuilles et environ 9 900 lignes de VBA ;
- *Annexe MOP utilisateurs DCO V1 (05/2024)* : mode opératoire utilisateur ;
- *Mode opératoire SGO – Admin v3* : mode opératoire administrateur.

Le but : comprendre ce que fait l'outil d'origine, ce qu'il faut **garder**, **refondre** ou **abandonner**, et en déduire le périmètre complet de la nouvelle application. La maquette actuelle (`VBA_maquette_SGO.xlsm`) ne couvre aujourd'hui que la saisie du questionnaire.

---

## 1. Ce que fait l'outil d'origine

L'outil est un classeur Excel **sans interface dédiée**. On travaille directement dans des feuilles protégées, avec des boutons et des macros.

1. **Page de garde** : l'acheteur choisit la direction et l'univers.
2. **Chargement des standards** : une macro copie dans `3-Elaboration.O` des formules modèles, issues de `150-Modele SGO`, qui :
   - calculent le niveau de standard de chaque question (E, D, U, Sp ou Titre) ;
   - calculent la valeur du standard ;
   - **pré-remplissent la réponse avec le standard**.
3. **Saisie** dans `3-Elaboration.O`, une ligne par question, avec :
   - les questions conditionnelles masquées ou affichées automatiquement à chaque saisie ;
   - la conformité **calculée automatiquement** ;
   - les commentaires saisis dans une fenêtre dédiée ;
   - des lignes de « notes » insérables sous une question.
4. **Revue des cas dérogatoires** : extraction vers les feuilles de synthèse `9-…`.
5. **Annexes** `10-…` : certaines sont alimentées par les réponses (Détail des lots, Criticité du référencement).
6. **Sorties** :
   - PDF du dossier, avec choix des feuilles dans le sommaire ;
   - **extraction GINO** ;
   - **pré-rédaction du CCAP** (feuilles `200-CCAP`, en cours de développement en juillet 2024).

## 2. Règles métier extraites (à conserver)

### 2.1 Niveau et valeur du standard

Le calcul est fait par la formule de `150-Modele SGO`, à partir de la matrice `140-Standards`.
- **Le niveau** est lu dans la colonne `<Direction>_Standard Etablissement (E) / Direction (D) / Univers (U)` :
  - il vaut `E`, `D`, `U`, `Sp` (spécifique à l'offre) ou `Titre` ;
  - colonne absente → `Sp` ;
  - niveau `D` ou `U` alors que la colonne de l'univers contient `Sp` → `Sp`.
- **La valeur du standard** se lit selon le niveau :
  - `E` → colonne `E` ;
  - `D` → colonne de la direction ;
  - `U` → colonne `U_<Univers>`.
- **La réponse est pré-remplie avec la valeur du standard.** Pour les questions de mode « Annexe 10 – … », la réponse est le nom de l'annexe.

### 2.2 Conformité : calculée automatiquement

C'est la réponse à la question Q6 du cadrage.

| Situation | Conformité |
|---|---|
| Titre, ou standard « Non concerné… », ou mode Annexe (hors `Sp`) | « Non concerné » ou « Non conforme au standard » |
| Niveau `Sp` et réponse saisie | **Spécifique Offre** |
| Standard incomplet (niveau ou valeur manquant) | message « Standards incomplets pour cette question » |
| Réponse vide | « Renseigner la réponse DCO » |
| Opérateur standard `=` | Conforme si réponse = standard **exactement** |
| Sinon | Conforme si la réponse **contient** le standard |
| Réponse « Standard non applicable… » | Non concerné |

### 2.3 Les quatre cas (schéma du mode opératoire utilisateur)

| | Conforme au standard | Non conforme | Spécifique (Sp) |
|---|---|---|---|
| **Identique à l'offre précédente** | ① Standard en place (système) | ② **Justification du cas dérogatoire** (saisie) | ③ Cas spécifique (système) |
| **Différent de l'offre précédente** | sans objet | sans objet | ④ **Justification du cas différenciant** (saisie) |

- **« Différence vs précédent »** ne concerne que les questions `Sp`. Elle vaut **« Identique » par défaut**, et l'acheteur la passe à « Différent » si besoin. Le mode opératoire insiste : la revue des cas « identiques » est **impérative**.
- **La justification** est obligatoire pour les cas ② et ④.

### 2.4 Questions conditionnelles

La table est `100-Parametres` : 302 règles. Chaque règle associe une question conditionnante, un opérateur, un critère et la question à afficher.
- `<>` = la réponse **contient** le critère, sans tenir compte des majuscules ; `=` = la réponse est **exactement égale** au critère.
- **Mon hypothèse précédente était fausse. Elle est corrigée dans le cadrage (P5).**
- Une question fille masquée reçoit la réponse **« Non concerné pour l'offre »**.
- Une question sans standard pour la direction ou l'univers est masquée, avec la réponse **« Non concerné pour la direction/univers »**.
- La colonne « Question conditionnante » (`x`) signale les questions mères.
- La règle d'administration impose **une seule condition par question fille**, avec 18 exceptions (exemple : El78c dépend de El71 et de El77).
- Choix multiples : le terme « **Mixte** » est imposé à la place de « Mono & Multi », pour ne pas fausser la recherche par « contient ».

### 2.5 Commentaires

- **Deux familles** :
  - **officiels** : Remarque, Question, Recommandation, Demande de modification. Ils sont imprimés et non supprimables ;
  - **intra** : internes au département et non imprimés.
- Chaque commentaire porte : date, n°, émetteur, destinataire, type, question, réponse apportée, commentaire, **réponse au commentaire**, **statut**.
- La référence d'une question qui porte un commentaire officiel apparaît en jaune.

### 2.6 Notes (« Insérer une ligne »)

Les cellules sont limitées à 255 caractères. L'outil permet donc d'ajouter des lignes `ElxxNote 01`, `02`… sous une question, en « Précision complémentaire ». Le texte va dans la réponse, et la ligne est classée `Sp`. Seul l'auteur peut la supprimer.

### 2.7 Synthèses (onglets `9-…`)

Le bouton « Revue cas dérogatoires » produit cinq listes, chacune avec la **question maître** associée :
- dérogatoires (non conformes) ;
- différenciants ;
- identiques ;
- spécifiques ;
- non concernés D/U.

### 2.8 Annexes (onglets `10-…`)

| Annexe | Particularité |
|---|---|
| Détail des lots | **Une ligne par lot**, alimentée par des réponses du DCO (`ReponseElxxx` ↔ ligne `LigneSourceInfoLots`) et par des listes de choix. **Répond en grande partie à la question Q1 (lots)** |
| Criticité du référencement | Bouton « import réponses DCO » ; protégée par mot de passe |
| Planning, Couverture géographique, Entreprises sourcées, Étude de rentabilité, Analyse critères, Listes annexes AE, Révision des prix sur indice | Tableaux à remplir librement, avec insertion de lignes |
| `10-X` | Annexes libres ajoutées par l'acheteur, incluses au PDF via le sommaire |

### 2.9 Autres feuilles

- `3-Intercalaire` : directions et personnes consultées.
- `3-Décisions El.O` : décisions du COPIL et prochaines étapes, avec les participants.
- `0-Sommaire` : choix des feuilles à inclure dans le PDF.
- `200-Extraction GINO` : 215 variables (univers VI pour l'instant), avec la correspondance entre question SGO et question GINO.
- `200-CCAP` et `100-Parametres CCAP` : environ 1 700 lignes de clauses du CCAP, avec 1 291 règles d'affichage. Chaque règle combine jusqu'à trois conditions sur les réponses du DCO (opérateurs `=` et `<>`), et des variables comme le détail des lots. **C'est un moteur de pré-rédaction du CCAP.**
- Administration :
  - `105-REF SGO` (couples direction/univers) ;
  - `120-REF Listes` (listes de choix multiples) ;
  - `130-References` (prochain identifiant) ;
  - `170-Format` (administrateurs, messages, mise en forme) ;
  - `110-Filtre…` (feuilles de calcul intermédiaires).

## 3. Garder, refondre, abandonner

| Élément | Décision proposée | Pourquoi |
|---|---|---|
| Règles des standards (2.1), de la conformité (2.2) et des quatre cas (2.3) | ✅ **Garder à l'identique** | C'est le cœur métier, validé et documenté |
| Opérateurs de condition `<>` (contient) et `=` | ✅ **Garder** | Compatibilité avec les 302 règles existantes |
| Valeurs « Non concerné pour l'offre / pour la direction/univers » | ✅ Garder | Les synthèses et GINO les utilisent |
| Matrice des standards `140-Standards` | 🔁 **Refondre** en table « une ligne par question × cible » (P4) | Aujourd'hui, ajouter une direction demande d'ajouter des colonnes, de nommer des plages et de copier des formats : 15 étapes dans le mode opératoire admin |
| `105-REF SGO`, `120-REF Listes`, plages nommées `ListeElxxx`, `ModeReponseElxxx`, `ReponseElxxx` | 🔁 Refondre en tables du référentiel | Le paramétrage par noms de plages et liens hypertexte est fragile ; la moitié de la section « Résolution de problèmes » du guide admin en découle |
| Formules modèles `150-Modele SGO` copiées-collées | ❌ **Abandonner**, remplacées par du code | C'est la source des lenteurs (« environ 30 secondes, veuillez patienter ») et des incohérences |
| Masquage de lignes Excel, boutons Afficher / Masquer les lignes, Reboot | ❌ Abandonner | Remplacés par le filtre de l'écran de saisie. « Reboot » contournait des plantages |
| Notes `ElxxNote` contre la limite de 255 caractères | 🔁 Refondre | Une zone de texte du UserForm n'a pas cette limite. On garde la possibilité d'ajouter une **précision par lot** |
| Commentaires officiels et intra, avec réponse et statut | ✅ Garder la logique, 🔁 nouvelle présentation | Page « Commentaires » du MultiPage, avec le type, le destinataire et le statut |
| Synthèses `9-…` et question maître | ✅ Garder | Elles deviennent un **écran de revue** et des exports, et ne sont plus produites par un filtre élaboré |
| Annexes `10-…` | ✅ Garder | Les annexes de saisie libre restent des **feuilles du dossier** ; le Détail des lots devient structuré (voir Q1) |
| Sommaire, export PDF, décisions COPIL, intercalaire | ✅ Garder | Sorties attendues du dossier |
| Extraction GINO | ✅ Garder | Paramétrage à reprendre dans le référentiel |
| Pré-rédaction du CCAP | ⏳ **À cadrer** : chantier à part entière | 1 700 lignes et 1 291 règles. Le moteur de conditions de la nouvelle application pourra le servir |
| Protection des feuilles par mot de passe en dur dans le code, ruban masqué, mot de passe de l'annexe criticité | ❌ Abandonner sous cette forme | Le dossier `.xlsx` ne contient plus de code, et la saisie passe par l'interface |
| Administrateurs déclarés par nom dans `170-Format` | 🔁 Garder le principe (P11) | Liste d'identifiants Windows dans le référentiel |
| Feuille « Modifications » (journal) | 🔁 Remplacée par l'historique Git et les notes de version du référentiel | |

## 4. Vision globale de l'application finale

```
SGO.xlam  (code + écrans + référentiel versionné)
│
├── Acheteur
│   ├── Page de garde ............ direction, univers, équipe projet, lots
│   ├── Saisie du questionnaire .. chapitres > thèmes > sous-thèmes > question
│   │     réponse, conformité (auto), différence vs offre précédente,
│   │     justification, précisions par lot, commentaires officiels/intra
│   ├── Revue ................... dérogatoires, différenciants, identiques,
│   │     spécifiques, non concernés (avec question maître)
│   ├── Annexes ................. Détail des lots (structuré) + annexes libres
│   ├── Décisions COPIL, intercalaire
│   └── Sorties ................. PDF du dossier, HTML, extraction GINO,
│                                 [plus tard] pré-rédaction du CCAP
│
├── Relecteur ................... commentaires officiels, réponses, statuts
│
└── Administrateur
    ├── Questions, structure, conditions (opérateurs <> et =)
    ├── Standards (E / directions / univers), directions et univers
    ├── Listes de choix, correspondance GINO, [plus tard] règles CCAP
    └── Contrôler et publier une version du référentiel
```

## 5. Impacts sur la maquette des écrans

1. **Ajouter « Différence vs offre précédente »** (Identique / Différent) dans le bloc Réponse, pour les questions `Sp`. Ajouter la **justification du cas différenciant** (cas ④).
2. **La conformité devient un affichage calculé**, et non plus trois boutons à cocher. L'acheteur ne choisit que « Non concerné ».
3. **Un bouton « Non concerné »** reprend l'ancien bouton « Assigner non concerné ».
4. **Commentaires** : un type (officiel / intra), une nature (Remarque, Question, Recommandation, Demande de modification), un destinataire, une réponse et un statut.
5. **Un nouvel écran « Revue »** avec les cinq synthèses.
6. **Une zone « Précisions par lot »**, qui remplace les lignes Note, si la réponse à Q1 le confirme.

## 6. Points de vigilance

- ⚠️ La feuille `10-Entreprises Sourcées` du classeur d'origine contient **un identifiant de connexion Creditsafe**. Il ne faut pas reprendre ce fichier tel quel dans le dépôt Git ni dans le modèle de dossier. Je ne l'ai pas commité.
- Le code d'origine contient le mot de passe de protection des feuilles **en clair**, 81 fois. Il ne sera pas repris.
- Le journal « Modifications » liste plusieurs bugs non résolus en juillet 2024 : retour arrière sur une conditionnante, police des justifications restée rouge, synthèses… La nouvelle architecture les évite en grande partie, puisque plus rien n'est stocké dans des formules ou des lignes masquées.

## 7. Questions ouvertes, mises à jour

| # | Question | État |
|---|---|---|
| Q1 | Lots : saisie par lot ? | **Tranchée** : les lots seront traités après la saisie des questions (feuille Détail des lots générée à partir du nombre de lots), dans une itération ultérieure |
| Q3 | Sens des colonnes | **Résolue** (2.1 à 2.3) |
| Q5 | Opérateurs | **Résolue** (2.4). Reste à confirmer le ET pour les conditions multiples |
| Q6 | Conformité | **Résolue** : calculée automatiquement |
| Q11 | **CCAP** : la pré-rédaction du CCAP est-elle dans le périmètre de la version 1, ou pour plus tard ? | Nouvelle |
| Q12 | **GINO** : seul l'univers VI est paramétré. Les autres univers sont-ils prévus ? | Nouvelle |
| Q13 | **Relecteurs** : utilisent-ils le même outil (`.xlam`), ou relisent-ils le PDF ou le dossier `.xlsx` ? | Nouvelle |
| Q14 | Le **flux de validation Emagin** impose-t-il un format de sortie (PDF uniquement ?) | Nouvelle |

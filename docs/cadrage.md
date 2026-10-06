# Cadrage : outil de saisie du DCO (SGO)

*Version 0.1, 2 octobre 2026. Proposition à valider.*

Ce document part de la maquette `VBA_maquette_SGO.xlsm` et de vos réponses aux questions de cadrage.
Chaque proposition porte un numéro (**P1**, **P2**…) pour que vous puissiez y répondre facilement : « P3 OK », « P5 non parce que… ».

Légende : ✅ acquis, 🔶 proposition à valider, ❓ question ouverte

---

## 1. Objectif

Remplacer la saisie directe dans la feuille `3-Elaboration.O` par une **interface guidée** qui permet à l'acheteur de remplir le questionnaire d'élaboration de son offre :
- environ **415 questions**, réparties en chapitres, thèmes et sous-thèmes ;
- pour chaque question, **le standard applicable** affiché selon la direction d'achat et l'univers ;
- la saisie de la **réponse**, de la **conformité au standard** et, si besoin, de la **justification** ;
- des **questions conditionnelles** (mère/fille) qui n'apparaissent que si la réponse de la question mère le justifie ;
- l'**export** du résultat en PDF ou HTML.

Une **fonction d'administration** permet de faire évoluer le référentiel (questions, structure, conditions, standards) et de livrer une nouvelle version aux acheteurs.

**Hors périmètre de la version 1** : saisie simultanée à plusieurs, utilisation dans Excel Online (navigateur).

## 2. Ce qui est acquis

| # | Sujet | Décision |
|---|---|---|
| ✅ | Utilisateurs | Plusieurs acheteurs, Excel de bureau **64 bits**, postes verrouillés par la DSI, macros internes acceptées |
| ✅ | Périmètre d'un fichier | **Un fichier = une offre = un acheteur**, avec plusieurs lots (marchés) |
| ✅ | Travail à plusieurs | Plusieurs personnes sur un même fichier, mais **jamais en même temps** |
| ✅ | Code et données | Le code doit être **indépendant** des fichiers de saisie |
| ✅ | Page de garde | Doit pouvoir être modifiée, avec le choix de conserver ou non les réponses déjà saisies |
| ✅ | Modes de réponse, statuts | Liste des modes et suivi de l'avancement : validés |
| ✅ | Identifiants des questions | Générés par l'application |
| ✅ | Design des écrans | Validé avec le chef de projet. Améliorations bienvenues pour la **démo aux acheteurs** |
| ✅ | Résolution | Doit fonctionner sur **PC portable** et sur poste avec **écran supplémentaire** |
| ✅ | Feuilles de la maquette | `Feuil3` et `carto` sont des extraits d'analyse. `structure_question` et `VBA_categories` sont des constructions de travail |

---

## 3. Propositions

### P1 🔶 Architecture : un complément `.xlam` et un fichier de saisie `.xlsx` sans macro

```
  SGO.xlam  ── code, écrans et référentiel des questions
     │         livré par l'admin ; une seule version installée par poste
     │
     │  crée / ouvre
     ▼
  Offre_2026-123.xlsx  ── uniquement les données d'une offre :
                           page de garde, lots, réponses, commentaires
                           AUCUNE macro
```

**Pourquoi cette option est la meilleure** (c'est aussi votre intuition) :
- Le fichier de l'acheteur **ne contient aucun code**. Il n'y a pas d'alerte de sécurité à l'ouverture, on peut le déposer sur Teams ou l'envoyer par mail, et il reste lisible sans l'outil.
- **Un seul exemplaire du code** par poste : une correction s'applique à tous les dossiers d'un coup.
- Le complément ajoute un **onglet « SGO » dans le ruban** (Nouveau dossier, Saisie, Export, Admin).

| Option | Avantages | Inconvénients |
|---|---|---|
| **`.xlam` + `.xlsx`** (recommandé) | Code centralisé, fichiers légers et sans macro | Le complément doit être installé une fois sur chaque poste |
| Modèle `.xltm` | Simple à distribuer | Chaque dossier embarque **sa propre copie du code** : une correction n'atteint pas les dossiers existants |
| `.xlsm` sur Teams | Rien à installer | **Le VBA ne fonctionne pas dans Excel Online**. Il faut ouvrir dans l'application de bureau, et le code est dupliqué dans chaque fichier |

⚠️ **Risque principal : la politique de la DSI.** Depuis 2022, Office bloque les macros des fichiers venant d'Internet, de Teams ou d'un mail (« Mark of the Web »). Il faut vérifier auprès de la DSI :
1. qu'un utilisateur peut installer un complément dans `%APPDATA%\Microsoft\AddIns` ;
2. si ce dossier est un **emplacement approuvé**, ou si l'on peut en déclarer un ;
3. si une **signature de code** est possible (certificat interne).

➡️ C'est pourquoi le plan (section 5) commence par une **preuve de concept de déploiement**, avant tout développement.
Solution de repli si le `.xlam` est refusé : un `.xlsm` complet par offre, comme aujourd'hui.

### P2 🔶 Référentiel : embarqué dans le `.xlam`, versionné, publié par l'admin

Vous avez proposé un fichier externe (txt, csv ou xml) éditable par l'admin. Voici mon analyse critique :

| | Fichier externe partagé | Référentiel dans le `.xlam` |
|---|---|---|
| Mise à jour | Immédiate pour tous | À la livraison d'une nouvelle version |
| Travail hors connexion (portable en déplacement) | ❌ Il faut accéder au partage, ou gérer une copie locale | ✅ |
| Cohérence entre code et données | ❌ Un nouveau mode de réponse ou une nouvelle règle peut exiger du code que l'utilisateur n'a pas encore | ✅ Le code et le référentiel sont livrés ensemble |
| Risque de modification non contrôlée | ❌ Le fichier est accessible à tous | ✅ |
| Format | CSV ou XML : accents, retours à la ligne et listes « a;b;c » délicats à gérer en VBA | Tableaux Excel, naturels |

**Recommandation :**
- Le référentiel est stocké dans des **tableaux cachés du `.xlam`**, avec un **numéro de version**.
- L'admin le modifie avec l'interface d'administration, sur **sa** copie, puis clique sur **« Publier »**. Le numéro de version augmente, une note de version est ajoutée, et le référentiel est **exporté en CSV dans Git** : on voit dans un diff exactement quelles questions ont changé.
- **Option (phase ultérieure)** : à l'ouverture, le `.xlam` regarde dans un dossier partagé si une version plus récente existe, et propose de l'installer. On obtient ainsi votre besoin (« l'utilisateur a toujours la liste à jour ») sans ses inconvénients.

### P3 🔶 Appliquer une nouvelle version du référentiel à un dossier en cours

Ce mécanisme réconcilie vos réponses aux questions 13 et 14 : le référentiel évolue, et les dossiers en cours suivent, **sans jamais perdre de réponse**.

Chaque dossier mémorise la version du référentiel avec laquelle il a été rempli. S'il est ouvert avec une version plus récente, l'outil affiche un **rapport des changements**, puis les applique :

| Changement dans le référentiel | Effet sur le dossier |
|---|---|
| Question ajoutée | Elle apparaît, non répondue |
| Libellé reformulé | La réponse est conservée |
| Mode de réponse ou liste de choix modifié | La réponse est conservée, avec le statut **« À revoir »** |
| Standard modifié | La conformité est recalculée. Statut **« À revoir »** si la réponse n'est plus conforme |
| Question retirée | La réponse est **archivée** (feuille `Archives`) et n'est plus affichée |
| Condition mère/fille modifiée | La visibilité est recalculée (voir P5) |

Règles imposées à l'admin :
- un identifiant n'est **jamais réutilisé** ;
- une question n'est jamais supprimée, elle est **désactivée**.

❓ La mise à jour doit-elle être **obligatoire** pour continuer la saisie, avec consultation possible sans mise à jour, ou **au choix** de l'acheteur ?

### P4 🔶 Modèle de données

**Référentiel** (dans le `.xlam`, une source unique) :

| Table | Contenu | Remplace dans la maquette |
|---|---|---|
| `R_Categories` | id (CHAP_1, THEME_1, SST_1…), libellé, type, parent, ordre | `structure_question` (lignes de titre), `VBA_categories`, le TreeView |
| `R_Questions` | id, libellé, sous-thème, ordre, mode de réponse, choix possibles, aide, alimente CCAP, active (oui/non) | `140-Standards` (colonnes A à F), `3-Elaboration.O` (partie référentiel) |
| `R_Conditions` | question fille, question mère, opérateur, valeur | `100-Parametres`, `Feuil3`, `carto` |
| `R_Standards` | question, cible (E, DAV, U_VI…), valeur | Les 40 colonnes de `140-Standards` |
| `R_Directions` | directions d'achat, directions, univers et leurs liens | `Feuil1` et les listes écrites en dur dans le code (`mon_dico_dir`, `colonne_Standard_direction`) |
| `R_Version` | numéro, date, auteur, notes | — |

Pour les standards, je propose **une ligne par question et par cible**, au lieu d'une colonne par cible. Ajouter un univers revient alors à ajouter des lignes, sans toucher au code ni à la structure.

**Dossier d'une offre** (`.xlsx`) :

| Feuille | Contenu |
|---|---|
| `Offre` | acheteur, n° d'offre, direction d'achat, direction, univers, version du référentiel, dates |
| `Lots` | liste des lots (marchés) |
| `Reponses` | question, lot, réponse, conformité, justification, statut, modifié par, modifié le |
| `Commentaires` | historique : question, auteur, date, texte. Remplace le format actuel « NOM_Com1: … » |
| `Archives` | réponses retirées (P3, P5, P6) |

La **migration** des données actuelles vers ce modèle sera automatique, et vérifiée par comptage (questions, conditions, standards).

### P5 🔶 Questions conditionnelles et effets en cascade

Vous m'avez demandé une proposition. La voici :

1. **Une question fille est visible** si toutes ses conditions sont vraies **et** si sa question mère est elle-même visible. La règle s'applique sur toute la chaîne, sans limite de niveaux.
2. **Une question mère non répondue** : ses filles sont masquées.
3. **Les réponses d'une fille qui devient masquée ne sont pas supprimées**, elles sont mises **en sommeil**. Elles ne comptent plus dans l'avancement et ne sont pas exportées. Si la condition redevient vraie, **elles réapparaissent**. Une erreur de clic ne fait donc rien perdre.
4. **Avant d'enregistrer une réponse de question mère qui masque des filles déjà remplies**, l'outil demande confirmation :
   *« Cette réponse masque 3 questions déjà renseignées (El100, El105, El107). Leurs réponses seront conservées mais ignorées. Continuer ? »*
5. Une fonction **« Purger les réponses en sommeil »** permet de nettoyer le dossier avant l'export final.
6. **Opérateurs** : `=`, `≠`, `contient` (pour les choix multiples), `est renseigné`. Si une fille a plusieurs conditions, je propose qu'elles s'appliquent **toutes à la fois** (ET).
7. **Contrôles côté admin** : interdire les boucles (A dépend de B qui dépend de A) et les questions mères inexistantes, et signaler une fille placée avant sa mère.

**Masquer ou griser ?** Je recommande de **masquer** les filles dans la liste de saisie, avec un compteur (« 12 questions, dont 3 masquées ») et un filtre « Afficher les masquées » qui les montre grisées, en lecture seule.

✅ **Sens de `100-Parametres` : tranché par le mode opératoire administrateur et le code d'origine** (voir `docs/analyse_existant.md`). Mon hypothèse de la version 0.1 était fausse.
- `<>` signifie « **la réponse contient** le critère », sans tenir compte des majuscules. Le code d'origine fait `InStr(LCase(réponse), LCase(critère)) > 0`.
  Exemple : El100 s'affiche si la réponse à El99 contient « Non ».
- `=` signifie « la réponse est **exactement égale** au critère ». Les deux règles `=` (El174, El78a) sont donc cohérentes.
- Une question fille masquée reçoit la réponse « Non concerné pour l'offre ». Une question non applicable à la direction ou à l'univers reçoit « Non concerné pour la direction/univers ».
- ❓ La règle d'administration impose **une seule condition par question fille**. Il existe pourtant 18 exceptions, comme El78c, conditionnée par El71 et par El77. Je propose que **toutes les conditions** doivent être remplies (ET). À confirmer.

### P6 🔶 Changement de la page de garde (direction, univers)

Quand l'acheteur modifie la page de garde, l'outil calcule d'abord l'impact : *« 47 questions ont un standard différent avec la nouvelle direction »*. Puis il propose :
- **a) Conserver les réponses** et marquer **« À revoir »** les 47 questions concernées (recommandé) ;
- **b) Tout réinitialiser** : les réponses sont archivées et non supprimées.

### P7 🔶 Statuts et avancement

| Statut | Signification |
|---|---|
| ○ Non répondue | |
| ✓ Conforme | Réponse conforme au standard |
| ! À justifier | Réponse différente du standard, sans justification |
| ✓ Justifiée | Réponse différente du standard, avec justification |
| ⟳ À revoir | Après une mise à jour du référentiel ou un changement de page de garde |
| – Non concernée | |
| (masquée) | Condition non remplie : n'apparaît pas et ne compte pas |

L'avancement est affiché par chapitre, thème et sous-thème dans la navigation, et en global en haut de l'écran.

❓ La conformité doit-elle être **calculée automatiquement** quand c'est possible (réponse choisie dans une liste, comparée au standard), ou toujours **déclarée** par l'acheteur ?

### P8 🔶 Écrans : une seule fenêtre en trois zones

**Le constat.** Aujourd'hui, trois fenêtres indépendantes vivent côte à côte. Sur un portable en 1366×768, elles se chevauchent, et l'utilisateur doit sans cesse les réorganiser. Dans le code, chaque fenêtre pilote les contrôles des autres : c'est la principale source de fragilité de la maquette.

**La proposition** : **une fenêtre principale** qui garde vos trois zones (chapitres, questions, détail), disposées côte à côte :

```
┌────────────────────────────────────────────────────────────────────────┐
│ Offre 2026-123 · DAI / Logiciels      [🔍 Rechercher…]   ▓▓▓▓▓▓░░░ 62 % │
├────────────────┬────────────────────────┬──────────────────────────────┤
│ CHAPITRES      │ QUESTIONS · 02.1 …     │ El104 · Clauses modificatives│
│ ✓ 01 Présent.  │ ✓ El102  Une partie…   │                              │
│ ◐ 02 Périmètre │ ● El104  Quelles sont… │ Standard (D) : « Clause de   │
│   ▸ 02.1 …     │ ○ El105  Si autre…     │ réexamen »  [Reprendre ▶]    │
│   ▸ 02.2 …     │ ! El107  Si clause…    │                              │
│ ○ 03 Prix      │                        │ Réponse     [______________] │
│ ○ 04 …         │                        │ Conformité  (•) Conforme     │
│                │ Filtre : [Toutes   ▾]  │ Justification [____________] │
│                │ 12 questions (3 masq.) │ Commentaires (2) ▸           │
├────────────────┴────────────────────────┴──────────────────────────────┤
│ ◀ Précédente     Suivante non répondue ▶          ✔ Enregistré à 10:42 │
└────────────────────────────────────────────────────────────────────────┘
```

**Des idées pour impressionner à la démo**, chacune représentant un vrai gain de temps pour l'acheteur :
- un bouton **« Reprendre le standard »** : la réponse est pré-remplie avec la valeur standard. Comme la majorité des réponses sont conformes, l'acheteur gagne beaucoup de temps ;
- **« Valider tout le sous-thème au standard »** en un clic, en laissant de côté les questions déjà répondues ;
- des **filtres** : non répondues, à justifier, à revoir, écarts au standard ;
- une **recherche** par identifiant (`El104`) ou par mot-clé ;
- un bouton **« Suivante non répondue »** ;
- des **raccourcis clavier**, par exemple `Ctrl+Entrée` pour valider et passer à la suivante ;
- l'**enregistrement automatique** à chaque changement de question, sans bouton « Enregistrer » ;
- un **aperçu des conséquences** d'une réponse : *« cette réponse fera apparaître 2 questions »*.

### P9 🔶 Résolution et écrans multiples

Les UserForms ont une taille fixe en points. Sur un portable dont l'affichage Windows est réglé à 125 ou 150 %, ils débordent. Sur un grand écran, ils paraissent minuscules. Je propose :
1. **À l'ouverture**, l'outil mesure la zone disponible **sur l'écran où se trouve Excel**, et la fenêtre en occupe environ 95 %. C'est ce qui règle le cas du second écran.
2. **La disposition est calculée par le code.** Les zones s'étirent : la navigation a une largeur fixe, la liste et le détail se partagent le reste. Ce n'est pas un simple zoom, qui rend le texte flou ou incohérent.
3. **En dessous d'une largeur minimale** (portable), la colonne des chapitres se replie et devient un bouton ☰.
4. Des boutons **A− / A+** règlent la taille du texte, et ce réglage est mémorisé pour chaque utilisateur.
5. **Aucun contrôle ActiveX** : le TreeView et la ProgressBar de la maquette (MSCOMCTL) sont remplacés par des contrôles standards. Ils sont fragiles en 64 bits et sur les postes verrouillés.
6. **Une configuration de test** systématique : portable 1366×768 à 125 %, 1920×1080 à 150 %, écran externe 1920×1080 à 100 %, et Excel ouvert sur le second écran.

❓ Quels sont les **modèles de portables** et la résolution des écrans utilisés par les acheteurs ?

### P10 🔶 Exports

- **PDF** : un rapport mis en forme (page de garde, puis questions par chapitre, avec standard, réponse, conformité et justification), généré par Excel.
- **HTML** : un fichier unique et autonome, avec un sommaire cliquable. Il s'ouvre dans n'importe quel navigateur et s'envoie facilement.
- **Trois variantes** : dossier complet, écarts au standard uniquement (utile pour la validation), questions à justifier.

❓ Faut-il un export **par lot** ?

### P11 🔶 Administration

- **Accès** (révisé le 06/10) : l'administration est un **complément séparé, `SGO_Admin.xlam`**, remis aux seuls administrateurs. Il partage avec `SGO.xlam` un socle de code commun : référentiel, moteur de conditions, calcul du standard et de la conformité. Les acheteurs n'ont donc pas le code d'administration, ce qui est plus sûr qu'un accès réservé dans le même outil. Les deux outils ne sont pas développés en parallèle (voir `docs/iterations.md`).
- **Fonctions** :
  - questions : création, modification, désactivation, déplacement ;
  - chapitres, thèmes et sous-thèmes ;
  - éditeur de conditions mère/fille, avec les contrôles de P5 ;
  - grille des standards ;
  - listes de choix ;
  - directions et univers ;
  - **publication** d'une nouvelle version.
- **Identifiants** : générés automatiquement. Les identifiants existants (`El1`, `El1b`, `El0a`…) sont conservés.

❓ Quel format pour les nouveaux identifiants ? Je propose `El` suivi du prochain numéro libre (par exemple `El1000`). L'ordre d'affichage est géré séparément, par une colonne « ordre ».

### P12 🔶 Organisation du code et du dépôt

```
planningVBA/
├── maquette/       VBA_maquette_SGO.xlsm + src/   (figée, sert de référence)
├── app/            SGO.xlam + src/                (la nouvelle application)
├── referentiel/    export CSV du référentiel      (historique lisible des évolutions)
├── docs/           cadrage, spécifications, guide utilisateur
└── outils/         VbaSync, scripts
```

- **VbaSync** devra être adapté : un `.xlam` n'est jamais le « classeur actif ». Il faudra pouvoir choisir le classeur cible.
- **Code** :
  - `Option Explicit` partout ;
  - des modules séparés : données, référentiel, conditions, export ;
  - des classes métier ;
  - des écrans qui **ne font qu'afficher** : aucun écran ne pilote les contrôles d'un autre ;
  - un **module de tests**, à lancer dans Excel.
- **Ruban** : un onglet personnalisé « SGO ». Je peux l'intégrer au `.xlam` sans passer par Excel.

---

## 4. Questions ouvertes

| # | Question |
|---|---|
| Q1 | **Lots** : les réponses se font-elles par lot ou pour toute l'offre ? Seulement certaines questions, ou certains chapitres, sont-ils par lot ? Combien de lots en général ? Je relie cette question aux 30 questions de mode « Annexe 10 – Détail lots ». |
| Q2 | Existe-t-il des **dossiers déjà remplis** dans l'ancien format (`3-Elaboration.O`) qu'il faudra **importer** ? |
| Q3 | Que signifient les colonnes **« Cas le plus fréquent dans l'établissement »**, **« Différence vs précédent »**, **« Opérateur standard »** (valeurs `=`, `D`, `U`, `Sp`…) et **« Alimente CCAP »** ? Que doit en faire l'outil ? |
| Q4 | Comment se saisissent les modes **« Annexe 10 – … »** et **« Cartographie ORECA »** ? Est-ce un tableau à remplir ? |
| Q5 | Sens des opérateurs dans `100-Parametres` (voir P5). |
| Q6 | Conformité calculée ou déclarée (voir P7). |
| Q7 | Matériel des acheteurs : modèles de portables, résolutions, version d'Office (Microsoft 365 ?). |
| Q8 | Réponses de la DSI sur le déploiement du `.xlam` (voir P1). |
| Q9 | Où seront stockés les dossiers : Teams/SharePoint synchronisé, lecteur réseau, poste local ? |
| Q10 | **Date de la démo** aux acheteurs. |

Remarque : les données contiennent des **variantes d'écriture** pour un même mode de réponse, par exemple « Liste choix unique » avec et sans espace final (182 et 27 occurrences). La migration les uniformisera.

## 5. Plan révisé

| Phase | Contenu | Livrable que vous validez |
|---|---|---|
| **0. Cadrage** | Ce document, puis vos retours | Cadrage validé (version 1.0) |
| **1. Preuve de concept** | Un `.xlam` minimal (onglet de ruban et fenêtre redimensionnable) et un dossier `.xlsx`. Test d'installation sur un poste verrouillé et sur les différents écrans | Le déploiement est confirmé par la DSI, et l'adaptation à la résolution est validée sur vos écrans |
| **2. Référentiel** | Tables du référentiel, migration automatique depuis la maquette, contrôles de cohérence | Référentiel migré, rapport de contrôle |
| **3. Socle** | Lecture et écriture des dossiers, moteur de conditions, statuts, tests automatiques | Les tests passent dans Excel |
| **4. Écran de saisie** | Fenêtre principale (P8, P9) | **Version de démo pour les acheteurs** |
| **5. Administration** | Fonctions admin, publication, mise à jour des dossiers (P2, P3, P11) | |
| **6. Finitions** | Exports, guide utilisateur, procédure d'installation | Version 1 |

**Si la démo arrive avant la phase 4**, on peut avancer une **maquette de l'écran principal** (P8), avec les vraies questions mais sans sauvegarde. Elle servira à recueillir l'avis des acheteurs sur l'ergonomie. Ces retours valent d'autant plus qu'ils arrivent tôt.

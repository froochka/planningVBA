# Journal des itérations

Une itération = un objectif, un livrable testable et des décisions prises avec le chef de projet.
Ce fichier sert aussi de point de reprise d'une session à l'autre.

---

## Itération 1 : mise en place du travail sous Git (01 → 02/10/2026)
- **Livré** : dépôt, `outils/VbaSync.bas` (ImporterVBA / ExporterVBA), extraction du code de la maquette dans `src/`.
- **Validé** : l'aller-retour Excel ↔ Git fonctionne. Premier correctif : l'ordre de tabulation de UserForm1.
- **Décisions** : le code de référence est dans `src/` ; la maquette s'appelle `VBA_maquette_SGO.xlsm`.

## Itération 2 : cadrage (02 → 05/10/2026)
- **Livré** : `docs/cadrage.md` (propositions P1 à P12, questions ouvertes).
- **Décisions** :
  - architecture : un `.xlam` pour le code et un `.xlsx` sans macro par offre (sous réserve de la DSI) ;
  - priorité à l'outil acheteur ; l'administration vient ensuite ;
  - une seule fenêtre de saisie ; la page de garde reste une boîte de dialogue séparée.

## Itération 3 : maquettes des écrans v1 → v3 (05 → 06/10/2026)
- **Livré** : `src/Maquette_Ecrans.bas` (écrans construits par code, contrôles MSForms uniquement), captures dans `docs/maquettes/`.
- **Validé** :
  - l'écran s'adapte à la résolution du poste ;
  - navigation chapitres > thèmes (TabStrip) > sous-thèmes (TabStrip boutons), détail en MultiPage ;
  - taille de texte pilotée par le code (base 10 pt), style modernisé, bloc Réponse aéré, bouton « Page de garde ».
- **Décisions** :
  - recherche dans tout le dossier, avec un panneau de résultats sous la barre (pas une nouvelle fenêtre) ;
  - par défaut, les éléments masqués disparaissent et une ligne d'information le signale.

## Analyse de l'existant (06/10/2026)
- **Livré** : `docs/analyse_existant.md`, à partir du classeur d'origine et des modes opératoires utilisateur et admin.
- **Corrections** :
  - opérateur de condition `<>` = « contient », `=` = égalité stricte ;
  - la conformité est **calculée** ;
  - les quatre cas (standard / Sp × identique / différent).
- **Décision** : l'administration sera un **complément séparé `SGO_Admin.xlam`**, avec un socle de code commun ; pas d'accès réservé dans l'outil acheteur. Ordre de développement : socle → acheteur → admin.

## Itération 4 : maquette conforme aux règles métier (en cours, 06/10/2026)
- **Objectif** : présenter au chef de projet un écran fidèle au mode opératoire.
- **Livré** (`src/Maquette_Ecrans.bas` v4, `docs/maquettes/v4_*.png`) :
  - conformité **affichée et calculée**, plus de boutons à cocher ;
  - **différence avec l'offre précédente** (Identique / Différente) pour les questions Sp ;
  - justification selon le cas : ② dérogation ou ④ cas différenciant ;
  - bouton **Non concerné** ;
  - page **Commentaires** : officiel / intra, nature, destinataire, réponse, statut ;
  - deux exemples : macros `MaquetteCasStandard` (cas ②) et `MaquetteCasSpecifique` (cas ④).
  - **page de garde refaite** à partir de l'écran validé avec le chef de projet (`UserForm1`) :
    - N° et libellé de la procédure ;
    - 6 noms de l'équipe projet, dans l'ordre de tabulation validé ;
    - direction achat, standard de direction, standard d'univers ;
    - « Conserver les réponses déjà saisies », Valider / Annuler.
    - Le bouton ADMIN est retiré : l'administration sera un outil séparé.
- **Décision (Q1, lots)** : les lots ne sont **pas** dans la page de garde ni dans la saisie des questions. Dans l'outil d'origine, une macro génère une feuille avec une ligne par lot, à partir du nombre de lots. Ce sera **intégré après la saisie des questions**, dans une itération ultérieure. Les lots sont retirés de la maquette.
- **Hypothèse en attendant la réponse** : Q13 relecteurs, ils commentent dans l'outil.
- **À valider avec le chef de projet** : 27 points (page de garde PG-1 à PG-6, écran de saisie ES-1 à ES-18, transverses TR-1 à TR-3), présentés avec capture, contexte et options dans la page de validation https://claude.ai/artifact/XPFB27aRZgxGY1i91PCVBV. Les décisions saisies dans la page sont enregistrées et relues par Claude.

### Décisions du chef de projet (06/10/2026), relevées dans la page de validation
| Point | Décision | Écart avec la proposition |
|---|---|---|
| PG-1 | Liste des 8 champs complète | |
| PG-2 | **Listes ouvertes** pour direction achat, standard de direction et d'univers, comme UserForm1 | ⚠️ proposition : listes déroulantes |
| PG-3 | Champs obligatoires : direction achat + standards uniquement | |
| PG-4 | **Pas de pré-remplissage** du nom de l'acheteur | ⚠️ proposition : pré-remplir |
| PG-5 | « Conserver les réponses » cochée par défaut ; les réponses dont le standard change passent « À revoir » | |
| PG-6 | Bouton ADMIN retiré (outil d'administration séparé) | |
| ES-1 | Une seule fenêtre | |
| ES-2 | Chapitres en onglets verticaux avec statut et pourcentage | |
| ES-3 | Navigation thèmes / sous-thèmes OK. Note : *indiquer de façon discrète la notion de thème et de sous-thème* | |
| ES-4 | **Éléments masqués visibles, grisés et non modifiables** | ⚠️ proposition : les faire disparaître |
| ES-5 | Recherche sur les identifiants et les libellés | |
| ES-6 | Réponses **pré-remplies avec le standard** au chargement, comme l'outil d'origine | |
| ES-7 | Conformité calculée, seulement affichée | |
| ES-8 | « Différence vs offre précédente » affichée en option discrète pour les questions avec standard | |
| ES-9 | « Identique » par défaut, revue obligatoire | |
| ES-10 | Justification non bloquante (statut « À justifier ») | |
| ES-11 | Bouton « Non concerné », même règle que l'origine | |
| ES-12 | Pas de précisions séparées : zone de réponse sans limite | |
| ES-13 | Enregistrement automatique | |
| ES-14 | Commentaires : à ajuster. Note : *garder une zone de commentaire simple, avec un nombre de caractères suffisant (5 000 par exemple)* | ⚠️ à préciser |
| ES-15 | Relecteurs (Q13) : **à revoir**, pas encore tranché | |
| ES-16 | **Onglet « Standards » seulement**, onglet « Historique » retiré | ⚠️ proposition : les deux |
| ES-17 | Texte 10 pt + A− / A+ | |
| ES-18 | Couleurs neutres actuelles, pas de charte | |
| TR-1 | Gestion des données en deux temps (conversion automatique, puis outil d'administration) | |
| TR-2 | CCAP dans une version ultérieure | |
| TR-3 | Contact DSI : le développeur | |

## Prochaines itérations (indicatif)
| N° | Contenu |
|---|---|
| 5 | Maquette de l'écran **Revue** (les 5 synthèses) |
| 6 | **Preuve de concept** : `.xlam` minimal qui lit les questions et les standards et calcule standard et conformité ; test d'installation sur un poste verrouillé de la DSI |
| 7+ | Écran de saisie branché sur les vraies données, enregistrement, commentaires, revue, **puis les lots** (feuille Détail des lots générée à partir du nombre de lots)… puis `SGO_Admin.xlam` |

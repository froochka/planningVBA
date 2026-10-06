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
- **Hypothèses en attendant les réponses** :
  - Q1 lots : pas de saisie par lot dans l'écran pour l'instant ;
  - Q13 relecteurs : ils commentent dans l'outil.
- **À valider avec le chef de projet** : le bloc Réponse, la page Commentaires, les réponses à Q1, Q11 et Q13.

## Prochaines itérations (indicatif)
| N° | Contenu |
|---|---|
| 5 | Maquette de l'écran **Revue** (les 5 synthèses) |
| 6 | **Preuve de concept** : `.xlam` minimal qui lit les questions et les standards et calcule standard et conformité ; test d'installation sur un poste verrouillé de la DSI |
| 7+ | Écran de saisie branché sur les vraies données, enregistrement, commentaires, revue… puis `SGO_Admin.xlam` |

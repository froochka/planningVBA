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
  - architecture : un `.xlam` pour le code et un `.xlsx` sans macro par offre (la DSI validera le `.xlam`) ;
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
- **Q13 tranchée** : pas de fonction relecteur dans l'outil (ES-15).
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
| ES-14 | **Zone de commentaire simple de 5 000 caractères** (option a : pas de type, nature, destinataire ni statut) | ⚠️ proposition : commentaires structurés |
| ES-15 | Relecteurs (Q13) : **pas de fonction relecteur** dans l'outil | ⚠️ hypothèse : ils commentent dans l'outil |
| ES-16 | **Onglet « Standards » seulement**, onglet « Historique » retiré | ⚠️ proposition : les deux |
| ES-17 | Texte 10 pt + A− / A+ | |
| ES-18 | Couleurs neutres actuelles, pas de charte | |
| TR-1 | Gestion des données en deux temps (conversion automatique, puis outil d'administration) | |
| TR-2 | CCAP dans une version ultérieure | |
| TR-3 | Contact DSI : le développeur. **Le complément `.xlam` sera validé par la DSI** | |

## Itération 5 : maquette v5, décisions du chef de projet appliquées
Modifications dans `src/Maquette_Ecrans.bas` (même mode de test : `MaquetteEcranSaisie`, `MaquettePageGarde`, puis `MaquetteNettoyer`) :
- **PG-2** : direction achat, standard de direction et standard d'univers en **listes ouvertes** ;
- **PG-4** : nom de l'acheteur vide ;
- **ES-3** : intitulés discrets « Thèmes » et « Sous-thèmes » au-dessus des barres de navigation ;
- **ES-4** : les éléments masqués restent visibles, **grisés et non modifiables** (onglet de sous-thème « ⊘ 03.2.3 », question El78a), le lien « Afficher les éléments masqués » est supprimé et le symbole ⊘ est ajouté à la légende ;
- **ES-8** : « Différence avec l'offre précédente (facultatif) » en option discrète pour les questions avec standard ;
- **ES-14** : page **Commentaire** réduite à une zone de texte de 5 000 caractères, avec compteur ;
- **ES-16** : onglet « Historique » retiré, il reste Réponse, Standards et Commentaire ;
- bouton « Rétablir les standards du sous-thème ».

## Itération 6 : la saisie sur les vraies données (15/10/2026)
- **Recentrage décidé avec l'utilisateur** : priorité à la saisie. L'écran **Revue** (les 5 synthèses de l'outil d'origine) et les **exports** (PDF / HTML pour présenter l'offre à la direction, **avec le logo de l'UGAP**) viendront une fois la saisie terminée.
- **Décisions** :
  - « Différence avec l'offre précédente » : **déclarative**, comme l'outil d'origine (pas de comparaison automatique avec l'offre précédente) ;
  - une question fille qui redevient masquée : ses réponses sont **mises en sommeil** (conservées, ignorées, reprises si la condition redevient vraie), avec une confirmation avant de masquer des questions déjà renseignées (cadrage P5) ;
  - filtre de l'écran de saisie : Toutes / Sans réponse / Justification manquante / **Standard modifié** (remplace « À revoir ») / Non conformes au standard ;
  - info-bulles et texte « Dépend de la réponse à El77 » ou « Sans standard pour votre direction/univers » sur les éléments grisés ; mention « Conditionne : … » sur les questions mères.
- **Livré** :
  - **référentiel migré** (`referentiel/R_*.csv`, modèle du cadrage P4) par `outils/migrer_referentiel.py`, avec `docs/rapport_migration.md` (comptages, anomalies) ;
  - **preuve d'équivalence** : 60 711 comparaisons avec les formules de l'outil d'origine, **0 écart** (`outils/moteur.py`) ;
  - **moteur VBA** (`SGO_Moteur.bas`) : standard, réponse pré-remplie, conformité, conditions mère/fille (y compris sur des thèmes et sous-thèmes entiers), masquage, statuts, avancement ;
  - **dossier d'offre** `.xlsx` sans macro (`SGO_Dossier.bas`) : feuilles Offre, Reponses, Archives ; enregistrement à chaque changement de question ;
  - **page de garde** réelle (`frmGarde`) et **écran de saisie** réel (`frmSaisie`) ;
  - **tests dans Excel** (`TesterMoteur`) : rejoue les valeurs attendues et des scénarios de conditions ;
  - corrections de la maquette (libellés du filtre, info-bulles, « Conditionne : »).
- **Validé dans Excel (15/10/2026)** : `TesterMoteur` → plus de 14 000 vérifications, **0 écart**, en 1 seconde. Le moteur VBA donne les mêmes résultats que l'outil d'origine.
- **Premier test de l'écran de saisie** : pas de bug. Retour d'ergonomie : libellés des questions tronqués (la liste manque de largeur, pas de hauteur ; 55 % des libellés dépassent 60 caractères). Corrigé, en gardant les 3 blocs côte à côte (décision ES-2 conservée) :
  - listes des chapitres et des questions refaites en lignes de texte qui passent à la ligne (2 lignes pour les chapitres, 3 pour les questions), avec un **vrai grisé** des éléments sans objet et la ligne choisie surlignée ;
  - la bande réservée aux sous-thèmes est récupérée quand le thème n'en a qu'un ;
  - largeurs rééquilibrées : le détail garde une largeur de lecture, la liste prend le reste.
- **Présentation modernisée** (maquette `docs/maquettes/v7_presentation_proposee.png`, validée par l'utilisateur) : code couleur unique des statuts (vert conforme, rouge à justifier, bleu spécifique, gris à saisir / sans objet), chapitre ouvert sur fond blanc avec mini-barres d'avancement, fil d'Ariane, étiquettes identifiant / mode, standard en « carte de référence », pastille « modifiée par rapport au standard », bandeau de conformité avec icône, titre et explication, zone de justification obligatoire en jaune et orange, confirmation « El78a enregistrée ».
- **À tester** : l'écran de saisie (macro `SGO`), voir « Tester l'itération 6 » dans `README.md`.
- **Découvertes de la migration** (détail dans le rapport) :
  - des conditions masquent des **thèmes et sous-thèmes entiers** (exemple : 03.2.3 « Offre exécutée » dépend de El77) ;
  - les chapitres 08 et 09 n'ont pas de thème : un thème « implicite » est créé ;
  - 8 conditions ont un critère absent des choix de leur question mère (la fille ne peut s'afficher que par une saisie libre) ;
  - 3 conditions visent des questions qui n'existent plus (El323, El323b, El323c) et 1 une mère inconnue (El335b).

## Prochaines itérations (indicatif)
| N° | Contenu |
|---|---|
| 7 | Retours sur la saisie réelle ; anomalies du rapport de migration à trancher avec le chef de projet ; **commentaires officiels / intra** si besoin ; **mise à jour du référentiel** dans un dossier existant (cadrage P3) |
| 8 | **Preuve de concept `.xlam`** : déplacer le code et le référentiel dans `SGO.xlam`, onglet de ruban ; test d'installation sur un poste de la DSI |
| 9 | Écran **Revue** (les 5 synthèses), puis **exports** PDF / HTML avec le logo de l'UGAP |
| 10+ | **Lots** (feuille Détail des lots générée à partir du nombre de lots), annexes, puis `SGO_Admin.xlam` |

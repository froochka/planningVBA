# Anomalies du référentiel : validation et intégration

`docs/anomalies_referentiel.xlsx` est généré par `python outils/lister_anomalies.py <classeur d'origine>.xlsm`.

## Le classeur
- **Synthèse** : une ligne par type d'anomalie, avec le nombre et l'explication.
- **À trancher** : une ligne par anomalie. Colonnes : type, gravité, question, libellé, élément lié (question mère…), constat, valeurs constatées, effet dans le nouvel outil, **correction proposée**, puis **Décision** (liste : Accepter la proposition / Autre correction / Ne rien changer) et **Commentaire**, à remplir par le chef de projet.
- **Corrections automatiques** : ce que la migration a déjà uniformisé (espaces, majuscules, niveaux « sp »…), pour information.

## Intégration des décisions
1. Le chef de projet remplit les colonnes Décision et Commentaire et renvoie le classeur (dans le dépôt ou en pièce jointe).
2. Claude traduit chaque décision en une ligne de **`referentiel/corrections.csv`** (n° d'anomalie, élément, champ, ancienne valeur, nouvelle valeur, décision). Ce fichier est relu et versionné dans Git : on sait qui a décidé quoi et quand.
3. `outils/migrer_referentiel.py` applique ces corrections **après** la lecture du classeur d'origine, puis régénère `referentiel/*.csv`, le rapport de migration et les valeurs attendues des tests. Le rapport liste les écarts **voulus** avec l'outil d'origine (ceux qui viennent des corrections) ; tout autre écart reste une erreur.
4. L'utilisateur fait `git pull`, puis **ImporterReferentiel** et **TesterMoteur** (0 écart attendu), et la version du référentiel augmente.

Pourquoi un fichier de corrections plutôt que modifier le classeur d'origine : l'outil d'origine reste intact (il sert encore), les corrections sont traçables et rejouables à chaque nouvelle extraction. Quand l'outil d'administration (`SGO_Admin.xlam`) existera, le référentiel sera modifié directement dans l'outil, avec un numéro de version et une note de version (cadrage P2), et ce fichier de corrections ne servira plus.

## Intégrer une nouvelle version de l'outil d'origine
Tant que l'outil d'origine sert de référence (avant `SGO_Admin.xlam`), ses évolutions (questions, standards, conditions mère/fille, listes de choix) sont reprises ainsi :
1. L'utilisateur fournit le nouveau classeur d'origine (`02_PhaseElaborationOffre_*.xlsm`). **Il n'est jamais commité** (identifiant de connexion dans « 10-Entreprises Sourcées »).
2. Claude lance `python outils/migrer_referentiel.py <classeur>.xlsm` :
   - `referentiel/*.csv` est régénéré ; **le numéro de version augmente** si le contenu change ;
   - `docs/rapport_migration.md` liste **les changements par rapport à la version précédente** (questions ajoutées, retirées, modifiées ; conditions ; standards ; structure), puis les anomalies et la preuve d'équivalence ;
   - les corrections de `referentiel/corrections.csv` sont réappliquées.
3. Claude relance `python outils/lister_anomalies.py <classeur>.xlsm` : les anomalies corrigées dans l'outil d'origine disparaissent du fichier du CDP.
4. L'utilisateur fait `git pull`, **ImporterReferentiel**, **TesterMoteur** (0 écart). À leur ouverture, les dossiers existants signalent qu'ils ont été remplis avec une version antérieure (la mise à jour complète des dossiers, avec rapport et « À revoir », est prévue à l'itération 11).

## Sources des listes de choix
Comme dans l'outil d'origine : 1) la liste déroulante de la cellule Réponse (choix unique) ; 2) la liste nommée `ListeElxxx` de « 120-REF Listes » (fenêtre de choix multiples) ; 3) à défaut, le texte de la colonne « Choix de réponse ».

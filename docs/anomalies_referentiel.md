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

# Consignes pour Claude

Projet : classeur Excel `VBA_maquette_SGO.xlsm` (à la racine) avec macros VBA, versionné sur Git. Utilisateur francophone : répondre en français.

- Le code VBA de référence est dans `src/` (UTF-8, sans en-têtes `Attribute VB_...`, un fichier par module). On modifie le code **uniquement là**.
- Ne pas modifier le `.xlsm` pour changer du code : l'utilisateur réimporte `src/` avec la macro `ImporterVBA` (`outils/VbaSync.bas`).
- Quand l'utilisateur ajoute ou met à jour un `.xlsm` et que `src/` n'est pas à jour, réextraire avec :
  `pip install oletools && python outils/extraire_vba.py <classeur>.xlsm`
  puis vérifier le diff avant de commiter.
- Modifier des feuilles, formules ou mises en forme (openpyxl avec `keep_vba=True`) seulement sur demande explicite, en prévenant des risques (graphiques, contrôles ActiveX…).
- Pour ajouter un module, créer le fichier `.bas` ou `.cls` dans `src/`. Pour un nouveau UserForm, l'utilisateur doit d'abord le créer dans Excel.
- `outils/VbaSync.bas` est importé tel quel par Excel (lu en ANSI) : n'y mettre que des caractères ASCII.
- Règles métier de l'outil d'origine (standards, conformité, 4 cas, conditions `<>` = contient / `=` = égal) : voir `docs/analyse_existant.md`. Cadrage et décisions : `docs/cadrage.md`.
- Ne jamais commiter le classeur d'origine `02_PhaseElaborationOffre_*.xlsm` tel quel : il contient un identifiant de connexion (feuille 10-Entreprises Sourcées).

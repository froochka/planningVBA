# Points en attente (à traiter quand le sujet sera abordé)

Mémo des modifications décidées ou repérées, mais **volontairement non faites**.
À relire au début de chaque itération : quand un sujet ci-dessous est abordé,
faire la modification, puis la cocher ici (date + commit).

## Référentiel : résidus repérés en vérifiant les corrections du CDP (02/11/2026)

Contexte : la liste des corrections du CDP (DCO des 15/04/26, 17/04/26 et 07/07/26)
a été vérifiée contre la base convertie (référentiel version 2, DCO du 07/07/26).
Bilan : toutes les corrections vérifiables y sont déjà. Les points suivants restent.

**Où les traiter** : par `referentiel/corrections.csv` (pas dans les CSV générés),
au moment où les décisions du CDP sur `docs/anomalies_referentiel.xlsx` seront intégrées,
ou dans l'outil d'administration (SGO_Admin) s'il existe d'ici là.
Les ajouter d'abord à `docs/anomalies_referentiel.xlsx` pour validation par le CDP.

- [ ] Faute « outer » (pour « outre ») : 1 occurrence restante, hors de El317a.
- [ ] Faute « applicabe » (pour « applicable ») : 1 occurrence restante, hors du libellé de El392.
- [ ] El151 : standard de l'univers Services_aux_batiments (« Oui, Facturation dépendante d'un échéancier ») absent des choix de la question.
- [ ] El313b : standards d'univers avec « relivraisons » au pluriel et « ; » final, alors que les choix sont au singulier.
- [ ] El507 : standard « Renseigner valeur! » pour Services_aux_batiments (valeur de remplissage à remplacer).
- [ ] El392 : question en mode Texte mais avec des choix « Oui | acompte… » (mode ou choix à revoir).
- [ ] El325 : ni liste déroulante ni liste nommée dans le classeur d'origine (source de ses choix à confirmer).

Non vérifiables sans l'ancienne version du DCO (à revoir si le CDP la fournit) :
standards PII / Logiciels de El111, El112b, El116 ; standards Santé de El150, El151 ;
lien de la cartographie 2026 (El1b, aucun lien stocké dans la base).

Hors périmètre (feuilles non reprises dans SGO) : 10-Entreprises sourcées, 10-Criticité, 10-Planning.

> Sécurité : ne jamais recopier ici, ni ailleurs dans le dépôt, le mot de passe de la feuille
> « 10-Entreprises sourcées » présent dans la liste du CDP.

## Retours du CDP (fichier « remarques CDP.xlsx », 10/11/2026)

- [ ] Lignes 31 et suivantes du fichier : à étudier plus tard (dont « Reprendre le standard » → « Rétablir le standard »). La ligne 32 (encadré « Justification de la dérogation » remonté) est faite le 12/11/2026.
- [ ] Onglet « Standards » de l'écran de saisie : à garder pendant le développement, **à retirer en production**.
- [x] Zone « Comparaison avec l'offre précédente » : tranchée par le CDP le 12/11/2026 (questions spécifiques : Identique par défaut / Différente ; questions avec standard : Non concerné par défaut / Identique / Différente ; mention « (justifier la différence dans l'onglet Commentaire) » à côté de « Différente » ; rappel si le commentaire est vide).
- [ ] Pas de copie de sauvegarde quand le DCO est sur le Bureau : le CDP doit confirmer s'il a vu le message « Pas de copie de sauvegarde » (Bureau synchronisé avec OneDrive ?).

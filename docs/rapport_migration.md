# Rapport de migration du référentiel

*Généré par `outils/migrer_referentiel.py` le 2026-10-22, à partir de `02_PhaseElaborationOffre_DossierConception_V4_07072026.xlsm`.*

## Comptages

| Élément | Nombre |
|---|---|
| Chapitres | 13 |
| Thèmes | 54 |
| Sous-thèmes (dont implicites, un par thème sans sous-thème) | 69 (52) |
| Questions | 413 |
| Questions système (direction, univers : utilisées par des conditions) | 2 |
| Conditions (lignes de `100-Parametres` : 302) | 298 |
| Lignes de standards (une par question et par cible) | 4083 |
| Standards de direction / univers | 6 / 21 |

Modes de réponse après uniformisation :

- Liste choix unique : 209
- Texte : 101
- Liste choix multiple : 62
- Annexe 10 - Détail lots : 30
- Liste de choix ou texte : 4
- Cartographie ORECA : 1
- Nombre : 1
- Annexe 10 - Couverture géographique : 1
- Annexe 10 - Entreprises sourcées : 1
- Annexe 10 - Analyse Critères : 1
- Annexe 10 - Criticité référencement : 1
- Annexe 10 - Listes Annexes AE : 1

## Preuve d'équivalence des standards

Pour chaque standard de direction (7) et chaque univers (21), et pour chaque question, le nouveau moteur (`outils/moteur.py`, qui lit les CSV migrés) calcule le niveau, la valeur du standard, la réponse pré-remplie et la conformité. Les résultats sont comparés à ceux des formules de l'outil d'origine (feuille `150-Modele SGO`), appliquées à la matrice `140-Standards`.

- Comparaisons : **60711**
- Écarts : **0**

Les combinaisons où la formule d'origine renvoie une erreur (`#N/A` : univers sans colonne de standards) sont exclues. Le nouveau moteur ignore alors l'univers au lieu de produire une erreur.

Le fichier `referentiel/tests/attendus_standards.csv` (2891 lignes, 7 combinaisons) permet de rejouer ce contrôle dans Excel avec la macro `TesterMoteur`.

## Changements par rapport à la version précédente (version 2)

### Questions modifiées (35)

- El73 : choix
- El60c : choix
- El134 : choix
- El370 : choix
- El59 : choix
- El343 : choix
- El179 : choix
- El500c : choix
- El387 : choix
- El107 : choix
- El103 : choix
- El493 : choix
- El351 : choix
- El271 : choix
- El316 : choix
- El123 : choix
- El0a : choix
- El111 : choix
- El132b : choix
- El60b : choix
- El206b : choix
- El377 : choix
- El271c : choix
- El114 : choix
- El173 : choix
- El35 : choix
- El104 : choix
- El332 : choix
- El361 : choix
- El334 : choix
- El220 : choix
- El308 : choix
- El30 : choix
- El181b : choix
- El327 : choix


## Anomalies et choix faits pendant la migration

### Liste retenue différente du texte « Choix de réponse » (84)

- El0a (liste nommée ListeEl0a retenue)
- El30 (liste nommée ListeEl30 retenue)
- El30b (liste déroulante retenue)
- El302 (liste déroulante retenue)
- El511 (liste déroulante retenue)
- El59 (liste nommée ListeEl59 retenue)
- El500c (liste nommée ListeEl500c retenue)
- El60b (liste nommée ListeEl60b retenue)
- El60c (liste nommée ListeEl60c retenue)
- El69 (liste déroulante retenue)
- El70 (liste déroulante retenue)
- El71 (liste déroulante retenue)
- El73 (liste nommée ListeEl73 retenue)
- El77 (liste déroulante retenue)
- El78l (liste déroulante retenue)
- El81b (liste déroulante retenue)
- El82 (liste déroulante retenue)
- El107 (liste nommée ListeEl107 retenue)
- El111 (liste nommée ListeEl111 retenue)
- El114 (liste nommée ListeEl114 retenue)
- El123 (liste nommée ListeEl123 retenue)
- El142 (liste déroulante retenue)
- El132 (liste déroulante retenue)
- El132b (liste nommée ListeEl132b retenue)
- El138 (liste déroulante retenue)
- El488 (liste déroulante retenue)
- El179 (liste nommée ListeEl179 retenue)
- El181b (liste nommée ListeEl181b retenue)
- El191 (liste déroulante retenue)
- El183l (liste déroulante retenue)
- El183b (liste déroulante retenue)
- El183e (liste déroulante retenue)
- El183j (liste déroulante retenue)
- El184 (liste déroulante retenue)
- El185 (liste déroulante retenue)
- El186 (liste déroulante retenue)
- El134 (liste nommée ListeEl134 retenue)
- El370 (liste nommée ListeEl370 retenue)
- El377 (liste nommée ListeEl377 retenue)
- El5c (liste déroulante retenue)
- El5e (liste déroulante retenue)
- El103 (liste nommée ListeEl103 retenue)
- El493 (liste nommée ListeEl493 retenue)
- El151b (liste déroulante retenue)
- El152 (liste déroulante retenue)
- El494 (liste déroulante retenue)
- El152b (liste déroulante retenue)
- El494b (liste déroulante retenue)
- El161c (liste déroulante retenue)
- El163 (liste déroulante retenue)
- El206b (liste nommée ListeEl206b retenue)
- El199b (liste déroulante retenue)
- El268 (liste déroulante retenue)
- El308 (liste nommée ListeEl308 retenue)
- El220 (liste nommée ListeEl220 retenue)
- El222 (liste déroulante retenue)
- El224a (liste déroulante retenue)
- El224c (liste déroulante retenue)
- El224d (liste déroulante retenue)
- El314 (liste déroulante retenue)
- El316 (liste nommée ListeEl316 retenue)
- El316a (liste déroulante retenue)
- El506a (liste déroulante retenue)
- El510a (liste déroulante retenue)
- El321 (liste déroulante retenue)
- El327 (liste nommée ListeEl327 retenue)
- El328 (liste déroulante retenue)
- El270 (liste déroulante retenue)
- El271 (liste nommée ListeEl271 retenue)
- El173 (liste nommée ListeEl173 retenue)
- El332 (liste nommée ListeEl332 retenue)
- El343 (liste nommée ListeEl343 retenue)
- El258c (liste déroulante retenue)
- El260a (liste déroulante retenue)
- El260b (liste déroulante retenue)
- El260g (liste déroulante retenue)
- El339c (liste déroulante retenue)
- El352 (liste déroulante retenue)
- El351 (liste nommée ListeEl351 retenue)
- El359 (liste déroulante retenue)
- El35 (liste nommée ListeEl35 retenue)
- El361 (liste nommée ListeEl361 retenue)
- El334 (liste nommée ListeEl334 retenue)
- El335f (liste déroulante retenue)

### Mode de réponse uniformisé (espaces, tiret) (188)

- El2 : « Liste choix unique  » → « Liste choix unique »
- El2b : « Liste choix unique  » → « Liste choix unique »
- El30b : « Liste choix unique  » → « Liste choix unique »
- El302 : « Liste choix unique  » → « Liste choix unique »
- El303 : « Liste choix unique  » → « Liste choix unique »
- El511 : « Liste choix unique  » → « Liste choix unique »
- El88 : « Liste choix unique  » → « Liste choix unique »
- El91 : « Liste choix unique  » → « Liste choix unique »
- El94 : « Liste choix unique  » → « Liste choix unique »
- El95 : « Liste choix unique  » → « Liste choix unique »
- El61 : « Annexe 10- Couverture géographique » → « Annexe 10 - Couverture géographique »
- El24 : « Liste choix unique  » → « Liste choix unique »
- El25 : « Liste choix unique  » → « Liste choix unique »
- El26 : « Liste choix unique  » → « Liste choix unique »
- El69 : « Liste choix unique  » → « Liste choix unique »
- El70 : « Liste choix unique  » → « Liste choix unique »
- El71 : « Liste choix unique  » → « Liste choix unique »
- El72 : « Liste choix unique  » → « Liste choix unique »
- El73 : « Liste choix unique  » → « Liste choix unique »
- El73e : « Liste choix unique  » → « Liste choix unique »
- El77 : « Liste choix unique  » → « Liste choix unique »
- El78f : « Liste choix unique  » → « Liste choix unique »
- El78b : « Liste choix unique  » → « Liste choix unique »
- El78c : « Liste choix unique  » → « Liste choix unique »
- El78e : « Liste choix unique  » → « Liste choix unique »
- El78h : « Liste choix unique  » → « Liste choix unique »
- El78l : « Liste choix unique  » → « Liste choix unique »
- El78m : « Liste choix unique  » → « Liste choix unique »
- El80 : « Liste choix unique  » → « Liste choix unique »
- El81b : « Liste de choix ou texte  » → « Liste de choix ou texte »
- El80a : « Liste choix unique  » → « Liste choix unique »
- El84 : « Liste choix unique  » → « Liste choix unique »
- El85b : « Liste choix unique  » → « Liste choix unique »
- El86 : « Liste choix unique  » → « Liste choix unique »
- El82 : « Liste choix unique  » → « Liste choix unique »
- El269 : « Liste choix unique  » → « Liste choix unique »
- El391 : « Liste choix unique  » → « Liste choix unique »
- El110 : « Liste choix unique  » → « Liste choix unique »
- El113 : « Liste choix unique  » → « Liste choix unique »
- El112b : « Liste choix unique  » → « Liste choix unique »
- El120 : « Liste choix unique  » → « Liste choix unique »
- El121 : « Liste choix unique  » → « Liste choix unique »
- El507 : « Liste choix unique  » → « Liste choix unique »
- El122 : « Liste choix unique  » → « Liste choix unique »
- El140 : « Liste choix unique  » → « Liste choix unique »
- El141 : « Liste choix unique  » → « Liste choix unique »
- El508 : « Liste choix unique  » → « Liste choix unique »
- El142 : « Liste choix unique  » → « Liste choix unique »
- El125 : « Liste choix unique  » → « Liste choix unique »
- El127 : « Liste choix unique  » → « Liste choix unique »
- El132 : « Liste choix unique  » → « Liste choix unique »
- El136 : « Liste choix unique  » → « Liste choix unique »
- El138 : « Liste choix unique  » → « Liste choix unique »
- El445 : « Liste choix unique  » → « Liste choix unique »
- El177 : « Liste choix unique  » → « Liste choix unique »
- El488 : « Liste choix unique  » → « Liste choix unique »
- El180a : « Liste choix unique  » → « Liste choix unique »
- El189 : « Liste choix unique  » → « Liste choix unique »
- El192 : « Liste choix unique  » → « Liste choix unique »
- El191 : « Liste choix unique  » → « Liste choix unique »
- El183 : « Liste choix unique  » → « Liste choix unique »
- El183l : « Liste choix unique  » → « Liste choix unique »
- El183b : « Liste choix unique  » → « Liste choix unique »
- El183e : « Liste choix unique  » → « Liste choix unique »
- El183g : « Liste choix unique  » → « Liste choix unique »
- El183i : « Liste choix unique  » → « Liste choix unique »
- El183j : « Liste choix unique  » → « Liste choix unique »
- El184 : « Liste choix unique  » → « Liste choix unique »
- El185 : « Liste choix unique  » → « Liste choix unique »
- El186 : « Liste choix unique  » → « Liste choix unique »
- El4 : « Liste choix unique  » → « Liste choix unique »
- El376 : « Liste choix unique  » → « Liste choix unique »
- El380 : « Liste choix unique  » → « Liste choix unique »
- El381 : « Liste choix unique  » → « Liste choix unique »
- El97 : « Liste choix unique  » → « Liste choix unique »
- El99 : « Liste choix unique  » → « Liste choix unique »
- El102 : « Liste choix unique  » → « Liste choix unique »
- El143 : « Liste choix unique  » → « Liste choix unique »
- El146 : « Liste choix unique  » → « Liste choix unique »
- El148 : « Liste choix unique  » → « Liste choix unique »
- El150 : « Liste choix unique  » → « Liste choix unique »
- El151 : « Liste choix unique  » → « Liste choix unique »
- El151b : « Liste choix unique  » → « Liste choix unique »
- El152 : « Liste choix unique  » → « Liste choix unique »
- El152b : « Liste choix unique  » → « Liste choix unique »
- El495 : « Liste choix unique  » → « Liste choix unique »
- El153 : « Liste choix unique  » → « Liste choix unique »
- El155 : « Liste choix unique  » → « Liste choix unique »
- El156 : « Liste choix unique  » → « Liste choix unique »
- El156b : « Liste choix unique  » → « Liste choix unique »
- El160 : « Liste choix unique  » → « Liste choix unique »
- El161 : « Liste choix unique  » → « Liste choix unique »
- El161c : « Liste de choix ou texte  » → « Liste de choix ou texte »
- El162 : « Liste choix unique  » → « Liste choix unique »
- El163c : « Liste choix unique  » → « Liste choix unique »
- El163d : « Liste choix unique  » → « Liste choix unique »
- El163e : « Liste choix unique  » → « Liste choix unique »
- El163f : « Liste choix unique  » → « Liste choix unique »
- El205a : « Liste choix unique  » → « Liste choix unique »
- El208 : « Liste choix unique  » → « Liste choix unique »
- El210 : « Liste choix unique  » → « Liste choix unique »
- El211 : « Liste choix unique  » → « Liste choix unique »
- El401 : « Liste choix unique  » → « Liste choix unique »
- El199b : « Liste choix unique  » → « Liste choix unique »
- El200 : « Liste choix unique  » → « Liste choix unique »
- El263 : « Liste choix unique  » → « Liste choix unique »
- El490 : « Liste choix unique  » → « Liste choix unique »
- El265 : « Liste choix unique  » → « Liste choix unique »
- El268 : « Liste de choix ou texte  » → « Liste de choix ou texte »
- El307 : « Liste choix unique  » → « Liste choix unique »
- El309 : « Liste choix unique  » → « Liste choix unique »
- El308g : « Liste choix unique  » → « Liste choix unique »
- El308h : « Liste choix unique  » → « Liste choix unique »
- El222 : « Liste choix unique  » → « Liste choix unique »
- El223 : « Liste choix unique  » → « Liste choix unique »
- El224b : « Liste choix unique  » → « Liste choix unique »
- El224d : « Liste choix unique  » → « Liste choix unique »
- El311 : « Liste choix unique  » → « Liste choix unique »
- El312 : « Liste choix unique  » → « Liste choix unique »
- El313 : « Liste choix unique  » → « Liste choix unique »
- El314 : « Liste choix unique  » → « Liste choix unique »
- El313b : « Liste choix unique  » → « Liste choix unique »
- El315 : « Liste choix unique  » → « Liste choix unique »
- El315b : « Liste choix unique  » → « Liste choix unique »
- El317 : « Liste choix unique  » → « Liste choix unique »
- El317a : « Liste choix unique  » → « Liste choix unique »
- El318a : « Liste choix unique  » → « Liste choix unique »
- El320 : « Liste choix unique  » → « Liste choix unique »
- El321 : « Liste choix unique  » → « Liste choix unique »
- El223d : « Liste choix unique  » → « Liste choix unique »
- El324 : « Liste choix unique  » → « Liste choix unique »
- El325 : « Liste de choix ou texte  » → « Liste de choix ou texte »
- El328 : « Liste choix unique  » → « Liste choix unique »
- El164a : « Liste choix unique  » → « Liste choix unique »
- El164 : « Liste choix unique  » → « Liste choix unique »
- El167 : « Liste choix unique  » → « Liste choix unique »
- El168 : « Liste choix unique  » → « Liste choix unique »
- El174 : « Liste choix unique  » → « Liste choix unique »
- El270 : « Liste choix unique  » → « Liste choix unique »
- El271b : « Liste choix unique  » → « Liste choix unique »
- El173b : « Liste choix unique  » → « Liste choix unique »
- El330 : « Liste choix unique  » → « Liste choix unique »
- El332b : « Liste choix unique  » → « Liste choix unique »
- El342 : « Liste choix unique  » → « Liste choix unique »
- El344 : « Liste choix unique  » → « Liste choix unique »
- El346 : « Liste choix unique  » → « Liste choix unique »
- El258 : « Liste choix unique  » → « Liste choix unique »
- El258c : « Liste choix unique  » → « Liste choix unique »
- El258d : « Liste choix unique  » → « Liste choix unique »
- El259 : « Liste choix unique  » → « Liste choix unique »
- El496 : « Liste choix unique  » → « Liste choix unique »
- El260e : « Liste choix unique  » → « Liste choix unique »
- El336 : « Liste choix unique  » → « Liste choix unique »
- El337b : « Liste choix unique  » → « Liste choix unique »
- El338 : « Liste choix unique  » → « Liste choix unique »
- El339b : « Liste choix unique  » → « Liste choix unique »
- El339c : « Liste choix unique  » → « Liste choix unique »
- El348 : « Liste choix unique  » → « Liste choix unique »
- El350 : « Liste choix unique  » → « Liste choix unique »
- El352 : « Liste choix unique  » → « Liste choix unique »
- El353 : « Liste choix unique  » → « Liste choix unique »
- El356 : « Liste choix unique  » → « Liste choix unique »
- El357 : « Texte  » → « Texte »
- El358 : « Liste choix unique  » → « Liste choix unique »
- El359 : « Liste choix unique  » → « Liste choix unique »
- El285 : « Liste choix unique  » → « Liste choix unique »
- El509 : « Liste choix unique  » → « Liste choix unique »
- El217 : « Liste choix unique  » → « Liste choix unique »
- El36 : « Liste choix unique  » → « Liste choix unique »
- El38 : « Liste choix unique  » → « Liste choix unique »
- El40 : « Liste choix unique  » → « Liste choix unique »
- El41 : « Liste choix unique  » → « Liste choix unique »
- El41b : « Liste choix unique  » → « Liste choix unique »
- El43 : « Liste choix unique  » → « Liste choix unique »
- El44 : « Liste choix unique  » → « Liste choix unique »
- El45 : « Liste choix unique  » → « Liste choix unique »
- El46 : « Liste choix unique  » → « Liste choix unique »
- El33 : « Liste choix unique  » → « Liste choix unique »
- El47 : « Liste choix unique  » → « Liste choix unique »
- El48 : « Liste choix unique  » → « Liste choix unique »
- El49 : « Liste choix unique  » → « Liste choix unique »
- El62 : « Liste choix unique  » → « Liste choix unique »
- El63 : « Liste choix unique  » → « Liste choix unique »
- El116 : « Liste choix unique  » → « Liste choix unique »
- El118 : « Liste choix unique  » → « Liste choix unique »
- El369 : « Liste choix unique  » → « Liste choix unique »
- El335a : « Liste choix unique  » → « Liste choix unique »
- El335f : « Liste choix unique  » → « Liste choix unique »

### Question placée directement sous un chapitre (thème « Généralités » créé) (3)

- El445
- El263
- El266

### Identifiant écrit différemment dans 140-Standards (majuscules) (10)

- EL492 / El492
- El260b / el260b
- El260c / el260c
- El260d / el260d
- El260i / el260i
- El260e / el260e
- El498 / el498
- El260f / el260f
- El260g / el260g
- El260h / el260h

### Choix repris de la colonne « Choix de réponse » (ni liste déroulante, ni liste nommée) (1)

- El325

### Niveau de standard écrit en minuscules (uniformisé) (4)

- El328b / DAG_SOFI : sp
- El328b / DS : sp
- El328b / Non_standardisé : sp
- El33b / DAV : sp

### Condition : critère absent des choix de la question mère (la fille ne peut jamais s'afficher par une liste) (4)

- El174 ← El164a = « Oui »
- El353b ← El353 <> « Oui. Niveau minimal imposé »
- El361b ← El361 <> « Guide de vente  »
- El73b ← El73 <> « autre »

### Condition : identifiant écrit différemment (majuscules) (3)

- el260d ← el260c
- el260f ← el260e
- el260i ← el260c

### Condition : question fille inconnue (règle ignorée) (3)

- El323 ← El320
- El323b ← El320
- El323c ← El320

### Condition : question mère inconnue (règle ignorée) (1)

- El335e ← El335b

### Question fille avec plusieurs conditions (toutes doivent être vraies) (18)

- El181 (2)
- El183b (2)
- El191 (2)
- El78h (2)
- El78i (3)
- El78j (3)
- El78c (2)
- El78e (2)
- El183 (2)
- ElTT502 (2)
- ElTT503 (2)
- El178 (2)
- El179 (2)
- El500 (2)
- El220 (2)
- El60d (2)
- El370 (2)
- El183e (2)

### Marquée « question conditionnante » sans aucune fille (13)

- El394
- El2b
- El511
- El112c
- El500
- El85b
- El82
- El112b
- El180
- El200
- El317a
- El338
- El44

### Question mère non marquée « conditionnante » (1)

- El60c

### Standard de direction sans colonne dans 140-Standards (toutes les questions en Sp) (1)

- DAG_Energie_et_environnement

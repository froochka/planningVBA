# VBA_maquette_SGO

Dépôt `planningVBA`, classeur versionné : `VBA_maquette_SGO.xlsm`, à la racine du dépôt.

Classeur Excel `.xlsm` versionné sur Git. Le code VBA est extrait en fichiers texte dans `src/` pour qu'on puisse le relire, le comparer et le modifier sur GitHub.

```
planningVBA/
├── VBA_maquette_SGO.xlsm ← le classeur (feuilles, données, mise en page des formulaires)
├── src/                 ← le code VBA en texte : c'est LA référence pour le code
│   ├── Module1.bas      (modules standards)
│   ├── ThisWorkbook.cls (classeur, feuilles, modules de classe)
│   └── frmSaisie.frm    (code des UserForms)
├── referentiel/         ← questions, structure, standards, conditions (CSV, migrés de l'outil d'origine)
│   └── tests/           ← valeurs attendues pour la macro TesterMoteur
├── docs/                ← cadrage, analyse de l'existant, itérations, rapport de migration
└── outils/
    ├── VbaSync.bas            ← macros ImporterVBA / ExporterVBA à installer dans Excel
    ├── extraire_vba.py        ← extraction sans Excel (utilisée par Claude)
    ├── migrer_referentiel.py  ← migration du référentiel depuis le classeur d'origine
    └── moteur.py              ← copie Python du moteur VBA (preuve d'équivalence)
```

**Règle d'or : pour le code, c'est `src/` qui fait foi.** Le `.xlsm` n'est jamais fusionné par Git.

## Installation (une seule fois par poste)

1. **Autoriser l'accès au projet VBA** : *Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion… > Paramètres des macros* → cocher **« Accès approuvé au modèle d'objet du projet VBA »**.
2. **Installer VbaSync dans le classeur de macros personnel** :
   - si vous n'avez pas encore de `PERSONAL.XLSB` : *Développeur > Enregistrer une macro* → « Stocker la macro dans : Classeur de macros personnel » → OK, puis *Arrêter l'enregistrement* ;
   - `Alt+F11` → dans l'explorateur de projets, clic droit sur **VBAProject (PERSONAL.XLSB)** → *Importer un fichier…* → `outils/VbaSync.bas` ;
   - enregistrer (`Ctrl+S` dans l'éditeur VBA).

Les macros `ImporterVBA` et `ExporterVBA` sont alors disponibles dans tous vos classeurs via `Alt+F8`. Vous pouvez aussi les ajouter à la barre d'accès rapide.

> Elles agissent sur le **classeur actif**, et lisent ou écrivent dans le dossier `src\` situé à côté de lui.

## Cycle de travail

### Récupérer des modifications (faites par Claude ou un collègue)
1. `git pull`
2. Ouvrir `VBA_maquette_SGO.xlsm`, puis `Alt+F8` → **ImporterVBA**
3. Tester
4. Si tout va bien : enregistrer le classeur, puis `git commit` + `git push` (pour que le `.xlsm` du dépôt soit à jour)

### Envoyer vos propres modifications
1. Modifier le code dans Excel (éditeur VBA) et/ou les feuilles
2. `Alt+F8` → **ExporterVBA** (met à jour `src/`)
3. Enregistrer le classeur
4. `git add -A`, `git commit -m "…"`, `git push`

### Éviter les conflits sur le `.xlsm`
Git ne peut pas fusionner deux versions d'un `.xlsm`. Il faut donc :
- toujours faire `git pull` **avant** d'ouvrir le classeur ;
- éviter que deux personnes modifient le classeur en même temps ;
- en cas de conflit, garder l'une des deux versions (`git checkout --theirs VBA_maquette_SGO.xlsm` ou `--ours`), puis lancer **ImporterVBA** pour réappliquer le code de `src/`.

## Limites à connaître
- Les fichiers de `src/` contiennent **uniquement le code**. Les attributs cachés (`VB_PredeclaredId`, descriptions de procédures…) et la **mise en page des UserForms** restent dans le `.xlsm`.
- Un **nouveau UserForm** doit être créé dans Excel. Son code peut ensuite être géré dans `src/`.
- Le projet VBA ne doit pas être protégé par mot de passe.

## Tester l'itération 6 (saisie sur les vraies données)

**Une seule fois** :
1. `git pull`, ouvrir `VBA_maquette_SGO.xlsm`, `Alt+F8` → **ImporterVBA**.
   Un avertissement signale que `frmGarde` et `frmSaisie` sont absents : c'est normal.
2. `Alt+F8` → **CreerFormulairesSGO** (crée les deux formulaires vides), puis de nouveau **ImporterVBA**.
3. `Alt+F8` → **ImporterReferentiel** (copie `referentiel\*.csv` dans des feuilles cachées `R_...`).
4. `Alt+F8` → **TesterMoteur** : le message doit annoncer **0 écart** (détail dans la feuille `T_Resultats`).
5. Enregistrer le classeur.

**Ensuite** : `Alt+F8` → **SGO**
- *Oui* : nouveau dossier → page de garde → choix du fichier `.xlsx` du dossier → écran de saisie ;
- *Non* : ouvrir un dossier `.xlsx` existant.

Le dossier d'une offre est un fichier `.xlsx` **séparé et sans macro**. Il s'ouvre **avec l'outil** : seule sa feuille d'accueil est visible dans Excel, les données sont masquées et protégées. À chaque ouverture, une copie est faite dans le sous-dossier « Sauvegardes SGO » (5 dernières). Ne le rangez pas dans le dépôt Git.

Après une mise à jour du référentiel (`git pull` qui modifie `referentiel/`), relancer **ImporterReferentiel**.


## Préparer une version de test pour le chef de projet
1. `git pull`, puis dans Excel : **ImporterVBA**, **ImporterReferentiel**, **TesterMoteur** (0 écart).
2. `Alt+F8` → **PreparerVersionTest** : crée la page « Démarrer » (3 boutons : nouveau dossier, ouvrir un dossier, noter une remarque) et la feuille « Remarques », masque les feuilles techniques, puis enregistre une copie, par exemple `SGO_test_CDP_v0.7.xlsm`, **hors du dossier Git**.
3. Envoyer cette copie avec la fiche `docs/fiche_test_CDP.pdf`.
4. Au retour, les remarques du chef de projet sont dans la feuille « Remarques » de son classeur.

Le classeur de développement est lui aussi modifié (page « Démarrer », feuilles masquées) : vous pouvez le fermer sans l'enregistrer, ou le garder ainsi (clic droit sur un onglet › Afficher pour revoir une feuille).

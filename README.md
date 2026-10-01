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
└── outils/
    ├── VbaSync.bas      ← macros ImporterVBA / ExporterVBA à installer dans Excel
    └── extraire_vba.py  ← extraction sans Excel (utilisée par Claude)
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

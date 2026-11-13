# Méthode : développer du VBA Excel avec Git et Claude

*Guide réutilisable, d'un projet à l'autre. Il s'adresse à Claude au début d'une nouvelle session, et à l'utilisateur.*
*Il ne couvre que la chaîne « code VBA ↔ Git ↔ Excel » ; le contenu métier de chaque projet reste dans son propre dépôt.*
*Les sources complètes des outils sont en annexe : ce fichier suffit pour mettre en place la méthode dans un nouveau dépôt.*
*Version 2 (13/11/2026) : menu Git `menu.bat` par projet, menu VbaSync dans Excel, compilation automatique après l'import.*

---

## 1. Le problème et l'idée

Un classeur `.xlsm` est un **fichier binaire** (un zip qui contient lui-même un fichier binaire `vbaProject.bin`). Git ne sait ni afficher ses différences ni fusionner deux versions. Claude ne peut ni l'ouvrir dans Excel, ni compiler ou exécuter le VBA.

**L'idée** : le code VBA vit **en texte**, un fichier par module, dans un dossier `src/` du dépôt. C'est **la référence**. Le classeur n'est qu'un « conteneur » dans lequel on recharge ce code.

```
        Claude (session cloud)                     Utilisateur (poste Windows)
        ─────────────────────                      ───────────────────────────
        modifie src/*.bas, *.cls, *.frm            menu.bat › 1 Récupérer (git pull)
        vérifie, commit, git push      ──GitHub──► Excel : Compléments › VbaSync › Importer
                                                     → le code de src/ remplace celui du classeur
                                                     → compilation automatique
                                                   teste dans Excel
                                                   (s'il modifie du code dans Excel :
                                                     VbaSync › Exporter → src/)
                                       ◄─GitHub──  menu.bat › 3 Envoyer (commit + push)
```

## 2. Les pièces du dispositif

| Élément | Rôle |
|---|---|
| `src/` | Le code VBA en texte : `Module.bas` (modules standards), `Classe.cls` (classes, mais aussi `ThisWorkbook` et les feuilles), `Formulaire.frm` (code des UserForms). **UTF-8, sans les en-têtes `Attribute VB_…`**, fins de ligne Windows. |
| `outils/VbaSync.bas` | Macros installées **une seule fois par poste** dans le classeur de macros personnel (`PERSONAL.XLSB`) : `ImporterVBA` (src → classeur actif, **puis compilation**), `ExporterVBA` (classeur actif → src), `CompilerVBA`. Elles créent au démarrage d'Excel un **menu « VbaSync »** dans l'onglet *Compléments* du ruban (plus besoin de `Alt+F8`). Source en annexe A. |
| `menu.bat` | **Menu Git du projet**, à la racine de chaque dépôt (double-clic, ou VbaSync › Menu Git) : récupérer, état, envoyer, historique, ouvrir le classeur, console. Il travaille toujours dans **son** dossier et affiche le dépôt et la branche : impossible de se tromper de projet. Source en annexe C. |
| `outils/verifier_vba.py` | Vérificateur statique pour Claude, qui ne peut pas compiler : blocs, déclarations, identifiants, noms ambigus, caractères hors cp1252. Source en annexe D. |
| `outils/extraire_vba.py` | Extraction **sans Excel** (pour Claude) : `pip install oletools` puis `python outils/extraire_vba.py Classeur.xlsm`. Produit exactement le même format qu'`ExporterVBA`. Source en annexe B. |
| `.gitattributes` | Déclare les classeurs comme binaires et les fichiers de code en texte avec fins de ligne Windows. |
| `.gitignore` | Ignore les fichiers de verrouillage d'Excel (`~$*`). |
| `CLAUDE.md` | Consignes lues automatiquement par Claude à chaque session (modèle au §6). |
| `README.md` | Le cycle de travail pour l'utilisateur (installation, `git pull`, import…). |

## 3. Ce que fait `ImporterVBA`, précisément

La macro agit sur le **classeur actif** et lit le dossier `src\` **situé à côté de lui** (donc le classeur doit être enregistré dans le dossier du dépôt).

Pour chaque fichier `.bas`, `.cls`, `.frm` de `src\` :
1. le fichier est lu en UTF-8 (ADODB.Stream), les éventuels en-têtes `VERSION…` / `Attribute VB_…` sont retirés (cela permet aussi d'importer un fichier exporté à la main depuis l'éditeur VBA) ;
2. **si le composant existe** dans le classeur (même nom) : tout son code est effacé (`DeleteLines`) puis remplacé (`AddFromString`). Le composant lui-même n'est pas supprimé : c'est ce qui permet de mettre à jour **`ThisWorkbook`, les modules de feuille et les UserForms** sans perdre leur mise en page ni leurs attributs ;
3. **s'il n'existe pas** : un module standard (`.bas`) ou un module de classe (`.cls`) est **créé** ; un UserForm (`.frm`) ne l'est pas : un avertissement demande de le créer d'abord dans l'éditeur VBA (sa mise en page n'est pas dans `src/`).

Enfin, les modules standards et de classe **présents dans le classeur mais absents de `src\`** sont listés, et la macro propose de les supprimer.

Puis la macro **compile le projet** (commande *Débogage > Compiler* de l'éditeur VBA) et l'annonce dans son message final : « Compilation OK » ou « ERREUR DE COMPILATION » (l'éditeur VBA montre alors la ligne en cause). Une compilation réussie grise la commande *Compiler* : c'est ce qui permet de le savoir.

`ExporterVBA` fait l'inverse : chaque composant est écrit dans `src\` (UTF-8 sans BOM, sans en-têtes), et les fichiers de `src\` qui ne correspondent plus à aucun module sont supprimés.

**Pourquoi « remplacer le code » plutôt que « supprimer puis réimporter le fichier »** (la méthode habituelle avec `VBComponents.Import`) :
- un module de document (`ThisWorkbook`, `Feuil1`) ne peut pas être supprimé ;
- un UserForm supprimé perd sa mise en page (le fichier `.frx` binaire) ;
- `Import` lit le fichier en ANSI : les accents d'un fichier UTF-8 seraient abîmés.

## 4. Le cycle de travail

### Mise en place sur un poste (une seule fois)
1. Excel : *Fichier > Options > Centre de gestion de la confidentialité > Paramètres… > Paramètres des macros* → cocher **« Accès approuvé au modèle d'objet du projet VBA »**.
2. Créer le classeur de macros personnel s'il n'existe pas (*Développeur > Enregistrer une macro* → « Classeur de macros personnel », puis arrêter l'enregistrement).
3. `Alt+F11` → clic droit sur `VBAProject (PERSONAL.XLSB)` → *Importer un fichier…* → `outils/VbaSync.bas` → enregistrer (`Ctrl+S`), puis **redémarrer Excel** : le menu **Compléments › VbaSync** apparaît.
   - Mise à jour de VbaSync : supprimer l'ancien module `VbaSync` de PERSONAL.XLSB, réimporter le nouveau, enregistrer.
   - Si PERSONAL.XLSB contient déjà une macro `Auto_Open`, y ajouter l'appel `CreerMenuVbaSync` et supprimer celle de VbaSync.
4. Git : `git clone …`, et de préférence `git config --global core.editor notepad` (évite l'éditeur Vim lors des fusions).
5. Créer sur le Bureau un **raccourci vers le `menu.bat`** de chaque projet (clic droit › Envoyer vers › Bureau), renommé avec le nom du projet.

### Récupérer le travail de Claude
1. Fermer le classeur, puis **menu.bat › 1 Récupérer** : le menu vérifie qu'Excel n'a pas le classeur ouvert, propose d'annuler les changements du `.xlsm` dus au dernier import (cas habituel), fait le `git pull` et affiche la version reçue.
2. **menu.bat › 5** (ouvrir le classeur) → **Compléments › VbaSync › Importer le code** : le message final doit dire « Compilation OK ».
3. Tester.

### Envoyer ses propres modifications
1. Modifier le code dans l'éditeur VBA (et/ou les feuilles, les UserForms).
2. **VbaSync › Exporter le code** (met `src/` à jour).
3. Enregistrer et fermer le classeur, puis **menu.bat › 3 Envoyer** (message, commit, push).

### Quand commiter le `.xlsm`
Uniquement quand **sa structure** change : nouveau UserForm, nouvelle feuille, nouvelle référence, mise en forme. Un simple import du code ne mérite pas de commit : le code est déjà dans `src/`, et le menu propose d'annuler ce changement au prochain « Récupérer ».

### Revenir en arrière
**menu.bat › 4** (historique) pour trouver la version, puis dans la console (**menu.bat › 6**) : `git checkout <commit> -- src/` (ou `git revert <commit>`), et **VbaSync › Importer**.

## 5. Bonnes pratiques

**Organisation**
- **`src/` fait foi pour le code.** On ne corrige jamais le code directement dans le `.xlsm` sans faire ensuite `ExporterVBA`. Claude ne modifie jamais le `.xlsm` pour changer du code.
- **Une seule personne modifie le classeur à la fois** : Git ne fusionne pas les `.xlsm`. Toujours `git pull` **avant** d'ouvrir le classeur. En cas de conflit sur le `.xlsm` : garder une des deux versions (`git checkout --theirs Classeur.xlsm` ou `--ours`), puis **ImporterVBA** pour réappliquer le code de `src/`.
- **Un dépôt = un dossier = un `menu.bat`** : toujours passer par le menu du projet (raccourci sur le Bureau) plutôt que par une console ouverte « quelque part ». Le menu affiche le dépôt et la branche.
- **Commiter le `.xlsm` seulement quand sa structure change** (voir §4).
- **Travailler par petites itérations** : un objectif, un livrable testable, un commit. Tenir un journal (`docs/iterations.md`) qui sert aussi de point de reprise entre deux sessions de Claude.
- **Aucune donnée sensible dans le dépôt** : pas de mot de passe dans le code, pas de classeur réel contenant des identifiants ou des données personnelles. Les données de travail (fichiers produits par l'outil) se rangent hors du dépôt.

**Écriture du code (pour Claude, qui ne peut pas compiler)**
- `Option Explicit` partout.
- **Caractères** : l'éditeur VBA stocke le code en ANSI (cp1252). Les accents français passent ; tout autre caractère (symboles ✓ ◐ ⊘, flèches…) doit être écrit avec `ChrW(&H2713)`. `outils/VbaSync.bas` lui-même doit rester **en ASCII pur** (il est importé tel quel par l'éditeur VBA).
- **VBA ne distingue pas majuscules et minuscules** : une variable `uiFond` et une fonction `UIFond` provoquent « nom ambigu » ; `wc` et `wC` sont la même variable.
- Paramètres objets passés `ByVal` (évite « type d'argument ByRef incompatible ») ; pas de `CStr` d'un nombre décimal dans une chaîne lue par Excel (virgule décimale en français) : utiliser `CLng` ou `Format$`.
- Ne pas dépendre de la langue d'Excel : `CStr(True)` vaut « Vrai » en français.
- **Avant chaque commit**, Claude relit son code et le passe dans le vérificateur statique : `python outils/verifier_vba.py src` (équilibre des blocs `If/For/Do/Select/With`, déclarations en double, identifiants non déclarés, noms ambigus, caractères hors cp1252). Il ne compare que les nouvelles remarques à celles d'avant sa modification.
- Côté utilisateur, **VbaSync › Importer** compile automatiquement : toute erreur restante est signalée dès l'import.
- Les règles métier se testent dans Excel avec une macro de tests (valeurs attendues générées hors Excel quand c'est possible), plutôt qu'à la main.

**UserForms**
- La **mise en page** d'un UserForm (fichier `.frx`) n'est pas dans `src/` : un nouveau UserForm se crée dans Excel, puis son code est géré dans `src/`.
- Alternative pratique : des UserForms **vides** dont tous les contrôles sont créés par le code (`Me.Controls.Add`) ; tout est alors dans `src/`, et une petite macro peut créer le formulaire vide (`VBProject.VBComponents.Add(3)`).
- Dans un événement d'un contrôle, **ne jamais supprimer ce contrôle** (Excel peut planter) : mettre à jour sur place, ou fermer et rouvrir la fenêtre.
- Pas de contrôles ActiveX (TreeView, ProgressBar…) : fragiles en 64 bits et sur postes verrouillés.

## 6. Mise en place dans un nouveau projet (checklist pour Claude)

1. Créer `outils/VbaSync.bas`, `outils/extraire_vba.py`, `menu.bat` et `outils/verifier_vba.py` à partir des **annexes A à D** (copie exacte ; `menu.bat` en ASCII avec fins de ligne Windows).
2. Créer `.gitattributes` :
   ```
   *.xlsm binary
   *.xlsx binary
   *.xlam binary
   *.xlsb binary
   *.frx  binary
   *.bas text eol=crlf
   *.cls text eol=crlf
   *.frm text eol=crlf
   *.bat text eol=crlf
   ```
   et `.gitignore` :
   ```
   ~$*
   Thumbs.db
   desktop.ini
   __pycache__/
   ```
3. Si l'utilisateur fournit un classeur : `pip install oletools && python outils/extraire_vba.py <Classeur>.xlsm` → `src/`. Vérifier le résultat (un fichier par module, accents corrects), puis commiter le classeur et `src/`.
   ⚠️ Vérifier d'abord que le classeur ne contient pas de données sensibles (identifiants, mots de passe, données personnelles) avant de le commiter.
4. Écrire un `README.md` avec les sections « Installation », « Récupérer des modifications », « Envoyer ses modifications », « Éviter les conflits » (§4 et §5).
5. Écrire un `CLAUDE.md`, par exemple :
   ```markdown
   # Consignes pour Claude
   Projet : classeur `<Nom>.xlsm` (racine) avec macros VBA, versionné sur Git. Répondre en français.
   - Le code VBA de référence est dans `src/` (UTF-8, sans en-têtes `Attribute VB_...`, un fichier par module). On modifie le code **uniquement là**.
   - Ne pas modifier le `.xlsm` pour changer du code : l'utilisateur réimporte `src/` avec la macro `ImporterVBA` (`outils/VbaSync.bas`).
   - Si l'utilisateur pousse un `.xlsm` plus récent que `src/` : `pip install oletools && python outils/extraire_vba.py <Nom>.xlsm`, puis vérifier le diff avant de commiter.
   - Pour un nouveau module : créer le `.bas` / `.cls` dans `src/`. Pour un nouveau UserForm : l'utilisateur le crée d'abord dans Excel.
   - `outils/VbaSync.bas` : caractères ASCII uniquement. Ailleurs : cp1252, sinon `ChrW`.
   - Vérifier le VBA avant de commiter (blocs, déclarations, identifiants, noms ambigus, caractères).
   ```
6. Faire valider le cycle par l'utilisateur sur un petit changement (par exemple un commentaire ajouté dans un module) : `git pull` → **ImporterVBA** → le changement apparaît dans l'éditeur VBA.

## 7. Limites connues

- Ne sont pas versionnés en texte : la mise en page des UserForms, les attributs cachés (`VB_PredeclaredId`, descriptions de procédures), les références du projet (*Outils > Références*), les feuilles et leurs formules. Ils restent dans le `.xlsm`.
- Le projet VBA ne doit pas être protégé par mot de passe.
- La compilation automatique repose sur l'état grisé de la commande *Compiler* (identifiant 578) : si une version d'Excel ne la trouve pas, le message le dit, et il suffit de compiler à la main (*Débogage > Compiler*).
- Le menu VbaSync est un menu « classique » (onglet *Compléments*), recréé à chaque démarrage d'Excel par `Auto_Open` de PERSONAL.XLSB.
- Les macros agissent sur le **classeur actif** : pour un complément `.xlam` (jamais actif), il faudra adapter `ClasseurCible` pour choisir le classeur.
- Excel Online n'exécute pas le VBA : tout se fait dans Excel de bureau.

---

## Annexe A : `outils/VbaSync.bas` (copie exacte, ASCII uniquement)

```vb
Attribute VB_Name = "VbaSync"
' =============================================================================
' VbaSync - synchronise le code VBA du classeur ACTIF avec le dossier src\
'           situe a cote de ce classeur (dans le depot Git).
'
'   ImporterVBA : src\  -> classeur, puis compilation (apres un git pull)
'   ExporterVBA : classeur -> src\   (avant un git commit)
'   CompilerVBA : Debogage > Compiler sur le projet du classeur actif
'
' Menu "VbaSync" dans l'onglet Complements du ruban, cree a chaque demarrage
' d'Excel (Auto_Open) : Importer, Exporter, Compiler, Menu Git du projet
' (menu.bat a la racine du depot), Ouvrir le dossier du projet.
'
' Installation : importer ce fichier dans PERSONAL.XLSB (voir README.md).
' Si PERSONAL.XLSB a deja une macro Auto_Open, y ajouter l'appel
' CreerMenuVbaSync et supprimer celle de ce module.
' Prerequis    : Options Excel > Centre de gestion de la confidentialite >
'                Parametres des macros > "Acces approuve au modele d'objet
'                du projet VBA".
'
' Les fichiers de src\ sont en UTF-8 et ne contiennent que le code
' (pas d'en-tetes "Attribute ..."). La mise en page des UserForms reste
' stockee dans le .xlsm.
' =============================================================================
Option Explicit

Private Const DOSSIER_SRC As String = "src"

' Valeurs de vbext_ComponentType (sans reference a VBIDE)
Private Const CT_STD As Long = 1
Private Const CT_CLASS As Long = 2
Private Const CT_FORM As Long = 3
Private Const CT_DOC As Long = 100

Private Const MENU_NOM As String = "VbaSync"
Private Const ID_COMPILER As Long = 578      ' commande Debogage > Compiler de l'editeur VBA

' -----------------------------------------------------------------------------
' Menu VbaSync (onglet Complements du ruban)
' -----------------------------------------------------------------------------
Public Sub Auto_Open()
    CreerMenuVbaSync
End Sub

Public Sub CreerMenuVbaSync()
    Dim barre As Object, menu As Object
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(MENU_NOM).Delete
    On Error GoTo 0
    Set barre = Application.CommandBars("Worksheet Menu Bar")
    Set menu = barre.Controls.Add(Type:=10, Temporary:=True)       ' 10 : menu deroulant
    menu.Caption = MENU_NOM
    AjouterEntree menu, "Importer le code (src -> classeur) et compiler", "ImporterVBA", 1952, False
    AjouterEntree menu, "Exporter le code (classeur -> src)", "ExporterVBA", 1953, False
    AjouterEntree menu, "Compiler le projet", "CompilerVBA", 2950, False
    AjouterEntree menu, "Menu Git du projet (mise a jour, envoi...)", "OuvrirMenuGit", 2645, True
    AjouterEntree menu, "Ouvrir le dossier du projet", "OuvrirDossierProjet", 23, False
End Sub

Private Sub AjouterEntree(menu As Object, texte As String, macro As String, ByVal icone As Long, ByVal separe As Boolean)
    Dim b As Object
    Set b = menu.Controls.Add(Type:=1, Temporary:=True)             ' 1 : bouton
    b.Caption = texte
    b.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
    b.BeginGroup = separe
    On Error Resume Next
    b.FaceId = icone
End Sub

' Lance menu.bat (racine du depot, a cote du classeur actif)
Public Sub OuvrirMenuGit()
    Dim wb As Workbook, chemin As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    chemin = wb.Path & "\menu.bat"
    If wb.Path = "" Or Dir(chemin) = "" Then
        MsgBox "menu.bat introuvable a cote du classeur actif :" & vbCrLf & chemin, vbExclamation, "VbaSync"
        Exit Sub
    End If
    CreateObject("WScript.Shell").Run "cmd /c start """" """ & chemin & """", 0, False
End Sub

Public Sub OuvrirDossierProjet()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If wb.Path = "" Then Exit Sub
    Shell "explorer.exe """ & wb.Path & """", vbNormalFocus
End Sub

' -----------------------------------------------------------------------------
' src\ -> classeur actif
' -----------------------------------------------------------------------------
Public Sub ImporterVBA()
    Dim wb As Workbook, dossier As String
    Dim fso As Object, f As Object, comp As Object
    Dim nom As String, ext As String, code As String
    Dim presents As Object, aSupprimer As Collection, item As Variant
    Dim nbImportes As Long, avertissements As String, liste As String

    If Not ClasseurCible(wb) Then Exit Sub
    dossier = wb.Path & "\" & DOSSIER_SRC
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossier) Then
        MsgBox "Dossier introuvable : " & dossier, vbExclamation, "VbaSync"
        Exit Sub
    End If

    Set presents = CreateObject("Scripting.Dictionary")
    presents.CompareMode = vbTextCompare

    For Each f In fso.GetFolder(dossier).Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            nom = fso.GetBaseName(f.Name)
            presents(nom) = True
            code = NettoyerEntete(LireUtf8(f.Path))
            Set comp = Composant(wb, nom)

            If comp Is Nothing Then
                Select Case ext
                    Case "bas"
                        Set comp = wb.VBProject.VBComponents.Add(CT_STD)
                        comp.Name = nom
                    Case "cls"
                        Set comp = wb.VBProject.VBComponents.Add(CT_CLASS)
                        comp.Name = nom
                    Case "frm"
                        avertissements = avertissements & vbCrLf & "- " & nom & _
                            " : UserForm absent du classeur (creez-le dans l'editeur VBA puis relancez)."
                End Select
            ElseIf Not TypeCompatible(comp.Type, ext) Then
                avertissements = avertissements & vbCrLf & "- " & f.Name & _
                    " : type different de celui du classeur, ignore."
                Set comp = Nothing
            End If

            If Not comp Is Nothing Then
                RemplacerCode comp, code
                nbImportes = nbImportes + 1
            End If
        End If
    Next f

    ' Modules du classeur qui n'existent plus dans src\
    Set aSupprimer = New Collection
    For Each comp In wb.VBProject.VBComponents
        If comp.Type = CT_STD Or comp.Type = CT_CLASS Then
            If Not presents.Exists(comp.Name) Then
                aSupprimer.Add comp.Name
                liste = liste & vbCrLf & "- " & comp.Name
            End If
        End If
    Next comp
    If aSupprimer.Count > 0 Then
        If MsgBox("Ces modules ne sont plus dans src\ :" & liste & vbCrLf & vbCrLf & _
                  "Les supprimer du classeur ?", vbYesNo + vbQuestion, "VbaSync") = vbYes Then
            For Each item In aSupprimer
                wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(item)
            Next item
        End If
    End If

    MsgBox nbImportes & " module(s) importe(s) dans " & wb.Name & "." & _
           IIf(avertissements <> "", vbCrLf & vbCrLf & "Attention :" & avertissements, "") & _
           vbCrLf & vbCrLf & ResultatCompilation(wb) & _
           vbCrLf & vbCrLf & "Pensez a enregistrer le classeur apres vos tests.", _
           vbInformation, "VbaSync"
End Sub

' -----------------------------------------------------------------------------
' Compilation (Debogage > Compiler)
' -----------------------------------------------------------------------------
Public Sub CompilerVBA()
    Dim wb As Workbook
    If Not ClasseurCible(wb) Then Exit Sub
    MsgBox ResultatCompilation(wb), vbInformation, "VbaSync"
End Sub

' Compile le projet du classeur. Apres une compilation reussie, la commande
' Compiler est grisee : c'est ce qui permet de savoir si elle a reussi. En cas
' d'erreur, l'editeur VBA affiche lui-meme le message et la ligne en cause.
Private Function ResultatCompilation(wb As Workbook) As String
    Dim ctl As Object
    On Error GoTo erreur
    Set Application.VBE.ActiveVBProject = wb.VBProject
    Set ctl = Application.VBE.CommandBars.FindControl(ID:=ID_COMPILER)
    If ctl Is Nothing Then
        ResultatCompilation = "Compilation non verifiee : lancez Debogage > Compiler dans l'editeur VBA."
        Exit Function
    End If
    If ctl.Enabled Then ctl.Execute
    If ctl.Enabled Then
        ResultatCompilation = "ERREUR DE COMPILATION : voir la ligne surlignee dans l'editeur VBA (Alt+F11)."
    Else
        ResultatCompilation = "Compilation OK."
    End If
    Exit Function
erreur:
    ResultatCompilation = "Compilation non verifiee (" & Err.Description & ")."
End Function

' -----------------------------------------------------------------------------
' classeur actif -> src\
' -----------------------------------------------------------------------------
Public Sub ExporterVBA()
    Dim wb As Workbook, dossier As String
    Dim fso As Object, f As Object, comp As Object
    Dim ext As String, code As String
    Dim attendus As Object, obsoletes As Collection, item As Variant
    Dim nbExportes As Long

    If Not ClasseurCible(wb) Then Exit Sub
    dossier = wb.Path & "\" & DOSSIER_SRC
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossier) Then fso.CreateFolder dossier

    Set attendus = CreateObject("Scripting.Dictionary")
    attendus.CompareMode = vbTextCompare

    For Each comp In wb.VBProject.VBComponents
        ext = Extension(comp.Type)
        If ext <> "" Then
            code = CodeDuModule(comp)
            If Len(code) > 0 Then
                EcrireUtf8 dossier & "\" & comp.Name & ext, code & vbCrLf
                attendus(comp.Name & ext) = True
                nbExportes = nbExportes + 1
            End If
        End If
    Next comp

    ' Fichiers de src\ qui ne correspondent plus a aucun module
    Set obsoletes = New Collection
    For Each f In fso.GetFolder(dossier).Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            If Not attendus.Exists(f.Name) Then obsoletes.Add f.Path
        End If
    Next f
    For Each item In obsoletes
        fso.DeleteFile item
    Next item

    MsgBox nbExportes & " module(s) exporte(s) vers " & dossier & _
           IIf(obsoletes.Count > 0, vbCrLf & obsoletes.Count & " fichier(s) obsolete(s) supprime(s).", ""), _
           vbInformation, "VbaSync"
End Sub

' -----------------------------------------------------------------------------
' Utilitaires
' -----------------------------------------------------------------------------
Private Function ClasseurCible(ByRef wb As Workbook) As Boolean
    Dim n As Long
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Aucun classeur actif.", vbExclamation, "VbaSync"
        Exit Function
    End If
    If wb Is ThisWorkbook Then
        MsgBox "Activez le classeur a synchroniser (pas celui qui contient VbaSync).", _
               vbExclamation, "VbaSync"
        Exit Function
    End If
    If wb.Path = "" Then
        MsgBox "Enregistrez d'abord le classeur dans le dossier du depot.", vbExclamation, "VbaSync"
        Exit Function
    End If
    On Error Resume Next
    n = wb.VBProject.VBComponents.Count
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Acces au projet VBA refuse." & vbCrLf & vbCrLf & _
               "Options Excel > Centre de gestion de la confidentialite > Parametres des macros >" & vbCrLf & _
               "cochez 'Acces approuve au modele d'objet du projet VBA'." & vbCrLf & _
               "(Le projet VBA ne doit pas etre protege par mot de passe.)", vbExclamation, "VbaSync"
        Exit Function
    End If
    On Error GoTo 0
    ClasseurCible = True
End Function

Private Function Composant(wb As Workbook, nom As String) As Object
    On Error Resume Next
    Set Composant = wb.VBProject.VBComponents(nom)
    On Error GoTo 0
End Function

Private Function Extension(typeComp As Long) As String
    Select Case typeComp
        Case CT_STD: Extension = ".bas"
        Case CT_CLASS, CT_DOC: Extension = ".cls"
        Case CT_FORM: Extension = ".frm"
    End Select
End Function

Private Function TypeCompatible(typeComp As Long, ext As String) As Boolean
    TypeCompatible = (Extension(typeComp) = "." & ext)
End Function

Private Function CodeDuModule(comp As Object) As String
    Dim s As String
    With comp.CodeModule
        If .CountOfLines > 0 Then s = .Lines(1, .CountOfLines)
    End With
    CodeDuModule = SansSautsFinaux(s)
End Function

Private Sub RemplacerCode(comp As Object, code As String)
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(code) > 0 Then .AddFromString code
    End With
End Sub

' Retire l'en-tete d'un fichier exporte par Excel (VERSION/BEGIN...END et
' lignes "Attribute VB_..."), pour accepter aussi des fichiers exportes a la main.
Private Function NettoyerEntete(texte As String) As String
    Dim lignes() As String, i As Long, debut As Long, resultat As String

    texte = Replace(Replace(texte, vbCrLf, vbLf), vbCr, vbLf)
    lignes = Split(texte, vbLf)
    debut = LBound(lignes)

    If UBound(lignes) >= debut Then
        If Left$(lignes(debut), 8) = "VERSION " Then
            Do While debut <= UBound(lignes)
                If Left$(lignes(debut), 17) = "Attribute VB_Name" Then Exit Do
                debut = debut + 1
            Loop
        End If
    End If
    Do While debut <= UBound(lignes)
        If Left$(lignes(debut), 13) <> "Attribute VB_" Then Exit Do
        debut = debut + 1
    Loop

    For i = debut To UBound(lignes)
        resultat = resultat & lignes(i) & IIf(i < UBound(lignes), vbCrLf, "")
    Next i
    NettoyerEntete = SansSautsFinaux(resultat)
End Function

Private Function SansSautsFinaux(s As String) As String
    Do While Len(s) > 0
        Select Case Right$(s, 1)
            Case vbCr, vbLf: s = Left$(s, Len(s) - 1)
            Case Else: Exit Do
        End Select
    Loop
    SansSautsFinaux = s
End Function

Private Function LireUtf8(chemin As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2            ' texte
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile chemin
    LireUtf8 = st.ReadText(-1)
    st.Close
End Function

' Ecrit en UTF-8 sans BOM
Private Sub EcrireUtf8(chemin As String, texte As String)
    Dim st As Object, bin As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText texte
    st.Position = 3        ' saute le BOM
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1           ' binaire
    bin.Open
    st.CopyTo bin
    bin.SaveToFile chemin, 2
    bin.Close
    st.Close
End Sub
```

## Annexe B : `outils/extraire_vba.py`

```python
"""Extrait le code VBA d'un classeur .xlsm vers src/ sans passer par Excel.

Produit le meme format que la macro ExporterVBA (outils/VbaSync.bas) :
un fichier par module, en UTF-8, sans les en-tetes "Attribute VB_...".

Usage : python outils/extraire_vba.py VBA_maquette_SGO.xlsm [dossier_src]
Prerequis : pip install oletools
"""
import pathlib
import sys

from oletools.olevba import VBA_Parser

EXTENSIONS = {".bas", ".cls", ".frm"}


def nettoyer(code: str) -> str:
    lignes = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    if lignes and lignes[0].startswith("VERSION "):
        while i < len(lignes) and not lignes[i].startswith("Attribute VB_Name"):
            i += 1
    while i < len(lignes) and lignes[i].startswith("Attribute VB_"):
        i += 1
    return "\n".join(lignes[i:]).rstrip("\r\n")


def extraire(classeur: pathlib.Path, dossier: pathlib.Path) -> None:
    parser = VBA_Parser(str(classeur))
    if not parser.detect_vba_macros():
        sys.exit(f"Aucun code VBA trouve dans {classeur}")

    dossier.mkdir(parents=True, exist_ok=True)
    ecrits = set()
    for _, _, nom_fichier, code in parser.extract_macros():
        if pathlib.Path(nom_fichier).suffix.lower() not in EXTENSIONS:
            continue
        if isinstance(code, bytes):
            code = code.decode("cp1252", errors="replace")
        code = nettoyer(code)
        if not code.strip() or nom_fichier in ecrits:
            continue
        (dossier / nom_fichier).write_text(code + "\n", encoding="utf-8")
        ecrits.add(nom_fichier)
        print(f"  {nom_fichier}")
    parser.close()

    for f in dossier.iterdir():
        if f.suffix.lower() in EXTENSIONS and f.name not in ecrits:
            f.unlink()
            print(f"  (supprime) {f.name}")
    print(f"{len(ecrits)} module(s) extrait(s) vers {dossier}/")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    extraire(pathlib.Path(sys.argv[1]),
             pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else "src"))
```

## Annexe C : `menu.bat` (racine du dépôt ; ASCII, fins de ligne Windows)

```bat
@echo off
rem ==========================================================================
rem  Menu Git du projet : a placer a la racine du depot, a cote du classeur.
rem  Double-clic pour l'ouvrir (ou Excel : Complements > VbaSync > Menu Git).
rem  Il travaille toujours dans SON dossier : impossible de se tromper de projet.
rem  Fichier en ASCII (pas d'accents) avec fins de ligne Windows.
rem ==========================================================================
setlocal
cd /d "%~dp0"
title Menu Git - %~dp0

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    echo Ce dossier n'est pas un depot Git : %CD%
    pause
    exit /b 1
)

:menu
cls
set BRANCHE=
for /f "delims=" %%b in ('git branch --show-current') do set BRANCHE=%%b
set DEPOT=
for /f "delims=" %%r in ('git config --get remote.origin.url') do set DEPOT=%%r
echo ==========================================================================
echo   Dossier : %CD%
echo   Depot   : %DEPOT%
echo   Branche : %BRANCHE%
echo ==========================================================================
echo.
echo   1 - Recuperer les mises a jour (git pull)
echo   2 - Etat : fichiers modifies
echo   3 - Envoyer mes modifications (commit + push)
echo   4 - Historique (10 dernieres versions)
echo   5 - Ouvrir le classeur
echo   6 - Console Git dans ce dossier
echo   Q - Quitter
echo.
choice /c 123456Q /n /m "Votre choix : "
if errorlevel 7 goto :eof
if errorlevel 6 goto console
if errorlevel 5 goto ouvrir
if errorlevel 4 goto historique
if errorlevel 3 goto envoyer
if errorlevel 2 goto etat
if errorlevel 1 goto maj
goto menu

rem --------------------------------------------------------------------------
:maj
echo.
call :excel_ferme || goto pause
rem Le classeur est souvent modifie par un simple import du code (ImporterVBA) :
rem on propose d'annuler ces changements pour que le pull puisse passer.
git diff --quiet -- "*.xlsm"
if errorlevel 1 (
    echo Le classeur .xlsm a ete modifie depuis la derniere mise a jour.
    echo  - si c'est seulement un import du code : repondez O, le code sera reimporte apres ;
    echo  - si vous avez modifie des feuilles a garder : repondez N, puis utilisez le choix 3.
    choice /c ON /m "Annuler les modifications du classeur"
    if errorlevel 2 goto pause
    git checkout -- "*.xlsm"
)
git pull
if errorlevel 1 goto pause
echo.
echo Version recue :
git log --oneline -1
echo.
echo Ensuite : ouvrez le classeur (choix 5), puis Complements ^> VbaSync ^> Importer le code.
goto pause

rem --------------------------------------------------------------------------
:etat
echo.
git status
goto pause

rem --------------------------------------------------------------------------
:envoyer
echo.
call :excel_ferme || goto pause
git status -s
echo.
set MSG=
set /p MSG=Message decrivant vos modifications (vide = annuler) :
if "%MSG%"=="" goto menu
git add -A
git commit -m "%MSG%"
git push
goto pause

rem --------------------------------------------------------------------------
:historique
echo.
git log --oneline -10
goto pause

rem --------------------------------------------------------------------------
:ouvrir
for %%f in (*.xlsm) do (
    start "" "%%f"
    goto menu
)
echo Aucun classeur .xlsm dans ce dossier.
goto pause

rem --------------------------------------------------------------------------
:console
start "Git - %CD%" cmd /k "git status"
goto menu

rem --------------------------------------------------------------------------
:pause
echo.
pause
goto menu

rem --------------------------------------------------------------------------
rem Verifie qu'aucun classeur .xlsm du dossier n'est ouvert dans Excel
rem (un fichier ouvert dans Excel ne peut pas etre ouvert en ecriture).
:excel_ferme
set OUVERT=
for %%f in (*.xlsm) do (
    2>nul ( >>"%%f" (call ) ) || set OUVERT=%%f
)
if defined OUVERT (
    echo Le classeur %OUVERT% est ouvert dans Excel : enregistrez-le et fermez-le d'abord.
    exit /b 1
)
exit /b 0
```

## Annexe D : `outils/verifier_vba.py`

```python
"""Vérificateur statique approximatif du VBA de src/ (pour Claude : le VBA n'est pas compilable hors d'Excel).

Usage : python outils/verifier_vba.py src [fichiers à contrôler pour les identifiants...]

- encodage : caractères hors cp1252 (Excel lit les fichiers en ANSI)
- blocs : If/End If, For/Next, Do/Loop, Select/End Select, With/End With, Sub/Function
- continuations de ligne (max 24)
- déclarations en double dans une procédure (sans casse)
- identifiants non déclarés (heuristique)
"""
import re, sys, os, glob

SRC = sys.argv[1]
FICHIERS = sorted(glob.glob(os.path.join(SRC, "*.bas")) + glob.glob(os.path.join(SRC, "*.cls")) + glob.glob(os.path.join(SRC, "*.frm")))
CIBLES = set(sys.argv[2:])  # fichiers à vérifier pour les identifiants

KEYWORDS = set("""
and as boolean byref byval call case const date dim do double each else elseif end enum
error exit explicit false for function get global gosub goto if in integer is let like long loop
me mod new next not nothing null object on option optional or paramarray preserve private property
public redim resume select set single static step string sub then to true type until variant wend
while with withevents xor currency byte decimal empty friend implements raiseevent attribute
lbound ubound collection
""".split())

BUILTINS = set(x.lower() for x in """
Len Mid Mid$ InStr InStrRev Replace Split Trim Trim$ LCase LCase$ UCase UCase$ Format Format$ CStr CLng CInt CDbl CSng CBool
Val IsArray Array ChrW Chr Chr$ IIf Now Timer MsgBox Dir CreateObject GetSetting SaveSetting DeleteSetting Shell CallByName VbLet VbGet VbMethod VbSet IsDate CDate IsEmpty
IsError Int StrComp Left Left$ Right Right$ RGB Err Application ThisWorkbook Workbooks ActiveSheet ActiveWorkbook
Worksheets Worksheet Workbook Range VBA UserForms Debug DoEvents IsNumeric IsObject TypeName VarType Abs Round
vbCrLf vbCr vbLf vbTab vbTextCompare vbBinaryCompare vbYes vbNo vbYesNo vbYesNoCancel vbOKCancel vbCancel vbQuestion
vbInformation vbExclamation vbCritical vbBoolean vbObjectError vbNullString vbString
xlSheetHidden xlSheetVisible xlSheetVeryHidden xlUp xlWhole xlWBATWorksheet xlOpenXMLWorkbook
MSForms ListBox TabStrip ComboBox TextBox OptionButton Label Frame MultiPage Page CheckBox ReturnInteger ReturnBoolean
Controls Cells Rows Columns Unload Load Hide Show Space String
UserForm_Initialize UserForm_QueryClose Cancel CloseMode KeyCode Shift
Kill FileLen Environ Choose ActiveWindow MkDir RmDir FileCopy
""".split())

def nettoyer(ligne):
    """Retire chaînes et commentaire."""
    out, i, dans = [], 0, False
    while i < len(ligne):
        ch = ligne[i]
        if ch == '"':
            if dans and i + 1 < len(ligne) and ligne[i + 1] == '"':
                i += 2; continue
            dans = not dans; out.append('""' if not dans else ''); i += 1; continue
        if not dans and ch == "'":
            break
        if not dans:
            out.append(ch)
        i += 1
    return "".join(out)

def lignes_logiques(texte):
    """Joint les continuations ; renvoie (n° ligne, ligne nettoyée, nb continuations)."""
    res, acc, debut, nb = [], "", None, 0
    for n, l in enumerate(texte.splitlines(), 1):
        c = nettoyer(l).rstrip()
        if debut is None:
            debut = n
        if c.endswith(" _"):
            acc += c[:-2] + " "; nb += 1
            continue
        acc += c
        res.append((debut, acc, nb))
        acc, debut, nb = "", None, 0
    return res

def instructions(ligne):
    """Découpe une ligne logique sur les ':' hors chaîne (chaînes déjà vidées)."""
    # attention aux étiquettes "erreur:" : traitées comme instruction vide
    return [p.strip() for p in re.split(r":(?!=)", ligne)]

erreurs = []
decl_globales = {}   # nom -> fichier
membres_classe = {}

def module_nom(f):
    return os.path.splitext(os.path.basename(f))[0]

# --- 1er passage : déclarations publiques et noms de modules
for f in FICHIERS:
    texte = open(f, encoding="utf-8").read()
    decl_globales[module_nom(f).lower()] = f
    for (n, l, nb) in lignes_logiques(texte):
        s = l.strip()
        m = re.match(r"(?i)^(public|global)?\s*(private\s+)?(sub|function|property\s+(get|let|set))\s+(\w+)", s)
        if m and not re.match(r"(?i)^private", s):
            decl_globales[m.group(5).lower()] = f
        m = re.match(r"(?i)^(public|global)\s+(const\s+)?(withevents\s+)?(.*)$", s)
        if m and not re.match(r"(?i)^public\s+(sub|function|property)", s):
            for part in m.group(4).split(","):
                mm = re.match(r"\s*(\w+)", part)
                if mm:
                    decl_globales[mm.group(1).lower()] = f

for f in FICHIERS:
    texte = open(f, encoding="utf-8").read()
    nomf = os.path.basename(f)
    # encodage
    for n, l in enumerate(texte.splitlines(), 1):
        try:
            l.encode("cp1252")
        except UnicodeEncodeError as e:
            erreurs.append(f"{nomf}:{n}: caractère hors cp1252 : {l[e.start:e.end]!r}")
    ll = lignes_logiques(texte)
    for (n, l, nb) in ll:
        if nb > 24:
            erreurs.append(f"{nomf}:{n}: {nb} continuations (max 24)")
        if len(l) > 1023:
            erreurs.append(f"{nomf}:{n}: ligne logique trop longue ({len(l)})")
    # blocs
    pile = []
    module_vars = set()
    proc = None
    locaux = {}
    for (n, l, nb) in ll:
        for ins in instructions(l):
            s = ins.strip()
            if not s:
                continue
            low = s.lower()
            low = re.sub(r"^(public|private|friend|static)\s+", "", low)
            if re.match(r"^(sub|function|property\s+(get|let|set))\s+\w+", low):
                if proc:
                    erreurs.append(f"{nomf}:{n}: procédure ouverte dans {proc}")
                proc = re.match(r"^(?:sub|function|property\s+\w+)\s+(\w+)", low).group(1)
                pile = []
                locaux = {}
                # paramètres
                mpar = re.search(r"\((.*)\)", s)
                if mpar:
                    for p in mpar.group(1).split(","):
                        mm = re.match(r"(?i)\s*(optional\s+)?(byval\s+|byref\s+)?(paramarray\s+)?(\w+)", p)
                        if mm and mm.group(4):
                            k = mm.group(4).lower()
                            if k in locaux:
                                erreurs.append(f"{nomf}:{n}: paramètre en double {k}")
                            locaux[k] = n
                continue
            if re.match(r"^end\s+(sub|function|property)\b", low):
                if pile:
                    erreurs.append(f"{nomf}:{n}: bloc(s) non fermé(s) dans {proc} : {pile}")
                proc = None
                continue
            if re.match(r"^exit\s+(sub|function|property|for|do)", low):
                continue
            # déclarations
            m = re.match(r"^(dim|static|const|redim(\s+preserve)?)\s+(.*)$", low) if proc else \
                re.match(r"^(dim|const|withevents|private|public|global)?\s*(const\s+|withevents\s+)?(.*)$", low)
            if proc and m and m.group(1).startswith(("dim", "static", "const")):
                for part in re.split(r",(?![^(]*\))", m.group(3)):
                    mm = re.match(r"\s*(withevents\s+)?(\w+)", part)
                    if mm:
                        k = mm.group(2)
                        if k in locaux:
                            erreurs.append(f"{nomf}:{n}: déclaration en double « {k} » dans {proc}")
                        locaux[k] = n
            low0 = s.lower()
            if re.match(r"^(private|public|friend)?\s*(sub|function|property)\b", low0):
                pass
            if not proc and re.match(r"^(dim|private|public|global|const)\b", low0) and not re.search(r"\b(sub|function|property|type|enum)\b", low0.split("(")[0]):
                body = re.sub(r"^(dim|private|public|global)\s+", "", low0)
                body = re.sub(r"^(const|withevents)\s+", "", body)
                for part in re.split(r",(?![^(]*\))", body):
                    mm = re.match(r"\s*(withevents\s+)?(\w+)", part)
                    if mm:
                        module_vars.add(mm.group(2))
            # blocs
            if re.match(r"^if\b.*\bthen\s*$", low):
                pile.append("If")
            elif re.match(r"^elseif\b", low) or low == "else":
                if not pile or pile[-1] != "If":
                    erreurs.append(f"{nomf}:{n}: {s.split()[0]} hors If")
            elif re.match(r"^end\s+if\b", low):
                if not pile or pile.pop() != "If":
                    erreurs.append(f"{nomf}:{n}: End If sans If")
            elif re.match(r"^for\b", low):
                pile.append("For")
            elif re.match(r"^next\b", low):
                if not pile or pile.pop() != "For":
                    erreurs.append(f"{nomf}:{n}: Next sans For")
            elif re.match(r"^do\b", low):
                pile.append("Do")
            elif re.match(r"^loop\b", low):
                if not pile or pile.pop() != "Do":
                    erreurs.append(f"{nomf}:{n}: Loop sans Do")
            elif re.match(r"^while\b", low):
                pile.append("While")
            elif re.match(r"^wend\b", low):
                if not pile or pile.pop() != "While":
                    erreurs.append(f"{nomf}:{n}: Wend sans While")
            elif re.match(r"^select\s+case\b", low):
                pile.append("Select")
            elif re.match(r"^end\s+select\b", low):
                if not pile or pile.pop() != "Select":
                    erreurs.append(f"{nomf}:{n}: End Select sans Select")
            elif re.match(r"^with\b", low):
                pile.append("With")
            elif re.match(r"^end\s+with\b", low):
                if not pile or pile.pop() != "With":
                    erreurs.append(f"{nomf}:{n}: End With sans With")
            elif re.match(r"^if\b.*\bthen\b", low):
                # If sur une ligne : vérifier les sous-instructions (rien à empiler)
                pass
    if proc:
        erreurs.append(f"{nomf}: procédure {proc} non fermée")

    for (n, l, nb) in ll:
        mm = re.match(r"(?i)^\s*(?:private\s+|public\s+|friend\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(\w+)", l)
        if mm:
            module_vars.add(mm.group(1).lower())
        mm = re.match(r"^(\w+):\s*$", l.strip())
        if mm:
            module_vars.add(mm.group(1).lower())
    # identifiants non déclarés (fichiers ciblés)
    if CIBLES and os.path.basename(f) not in CIBLES:
        continue
    proc = None
    locaux = set()
    for (n, l, nb) in ll:
        s = l.strip()
        low = re.sub(r"^(public|private|friend|static)\s+", "", s.lower())
        if re.match(r"^(sub|function|property\s+(get|let|set))\s+\w+", low):
            proc = True
            locaux = set()
            mpar = re.search(r"\((.*)\)", s)
            if mpar:
                for p in mpar.group(1).split(","):
                    mm = re.match(r"(?i)\s*(optional\s+)?(byval\s+|byref\s+)?(paramarray\s+)?(\w+)", p)
                    if mm and mm.group(4):
                        locaux.add(mm.group(4).lower())
            nom = re.match(r"^(?:sub|function|property\s+\w+)\s+(\w+)", low).group(1)
            locaux.add(nom)
            continue
        if re.match(r"^end\s+(sub|function|property)\b", low):
            proc = None
            continue
        if not proc:
            continue
        for ins in instructions(l):
            li = ins.strip().lower()
            m = re.match(r"^(dim|static|const)\s+(.*)$", li)
            if m:
                for part in re.split(r",(?![^(]*\))", m.group(2)):
                    mm = re.match(r"\s*(\w+)", part)
                    if mm:
                        locaux.add(mm.group(1))
                # type names after As are fine; skip checking
                continue
            if re.match(r"^\w+:$", ins.strip() + ":") and re.match(r"^\w+$", ins.strip()):
                # étiquette ou appel sans argument
                pass
            # retirer les noms d'arguments nommés (Filename:=)
            texte_ins = re.sub(r"\w+\s*:=", "", ins)
            texte_ins = re.sub(r"&H[0-9A-Fa-f]+&?", "0", texte_ins)
            for m in re.finditer(r"(?<![\w.$])([A-Za-z_]\w*)\$?", texte_ins):
                nom = m.group(1).lower()
                debut = m.start(1)
                avant = texte_ins[:debut].rstrip()
                if avant.endswith("."):
                    continue
                if nom in KEYWORDS or nom in BUILTINS or nom in locaux or nom in module_vars or nom in decl_globales:
                    continue
                if re.match(r"(xl|vb|mso|fm)[A-Z]", m.group(1)):      # constantes Excel / VBA / Office / MSForms
                    continue
                # "As Type"
                if re.search(r"(?i)\bas\s+(new\s+)?$", avant):
                    continue
                # étiquettes GoTo
                if re.search(r"(?i)\b(goto|resume)\s+$", avant):
                    continue
                erreurs.append(f"{nomf}:{n}: identifiant non déclaré ? « {m.group(1)} »")

# Noms de module (procédures, variables, constantes) qui ne diffèrent que par la
# casse : VBA ne distingue pas majuscules et minuscules (« nom ambigu »)
import collections
noms = collections.defaultdict(set)
for f in FICHIERS:
    mod = os.path.basename(f)
    for l in open(f, encoding="utf-8"):
        l = nettoyer(l.rstrip("\n"))
        mp = re.match(r"\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+\w+)\s+(\w+)", l)
        mm = re.match(r"(?:Public|Private|Global|Dim)\s+(?:Const\s+|WithEvents\s+)?(.*)$", l)
        if mp:
            noms[mp.group(1).lower()].add((mod, mp.group(1)))
        elif mm and not l.startswith((" ", "\t")):
            for part in re.split(r",(?![^(]*\))", mm.group(1)):
                x = re.match(r"\s*(\w+)", part)
                if x:
                    noms[x.group(1).lower()].add((mod, x.group(1)))
for k, v in noms.items():
    if len({n for _, n in v}) > 1 and len({m for m, _ in v}) < len(v) + 1:
        modules = {m for m, _ in v}
        publics_ou_meme_module = len(modules) == 1 or any(m.endswith(".bas") for m in modules)
        if publics_ou_meme_module:
            erreurs.append(f"nom ambigu (casse différente) : {sorted(v)}")

for e in erreurs:
    print(e)
print(f"{len(erreurs)} remarque(s)")
```

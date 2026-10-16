# Méthode : développer du VBA Excel avec Git et Claude

*Guide réutilisable, d'un projet à l'autre. Il s'adresse à Claude au début d'une nouvelle session, et à l'utilisateur.*
*Il ne couvre que la chaîne « code VBA ↔ Git ↔ Excel » ; le contenu métier de chaque projet reste dans son propre dépôt.*
*Les sources complètes des deux outils sont en annexe : ce fichier suffit pour mettre en place la méthode dans un nouveau dépôt.*

---

## 1. Le problème et l'idée

Un classeur `.xlsm` est un **fichier binaire** (un zip qui contient lui-même un fichier binaire `vbaProject.bin`). Git ne sait ni afficher ses différences ni fusionner deux versions. Claude ne peut ni l'ouvrir dans Excel, ni compiler ou exécuter le VBA.

**L'idée** : le code VBA vit **en texte**, un fichier par module, dans un dossier `src/` du dépôt. C'est **la référence**. Le classeur n'est qu'un « conteneur » dans lequel on recharge ce code.

```
        Claude (session cloud)                     Utilisateur (poste Windows)
        ─────────────────────                      ───────────────────────────
        modifie src/*.bas, *.cls, *.frm            git pull
        vérifie, commit, git push      ──GitHub──► Excel : macro ImporterVBA
                                                     → le code de src/ remplace celui du classeur
                                                   teste dans Excel
                                                   (s'il modifie du code dans Excel :
                                                     macro ExporterVBA → src/)
                                       ◄─GitHub──  enregistre le classeur, git add/commit/push
```

## 2. Les pièces du dispositif

| Élément | Rôle |
|---|---|
| `src/` | Le code VBA en texte : `Module.bas` (modules standards), `Classe.cls` (classes, mais aussi `ThisWorkbook` et les feuilles), `Formulaire.frm` (code des UserForms). **UTF-8, sans les en-têtes `Attribute VB_…`**, fins de ligne Windows. |
| `outils/VbaSync.bas` | Deux macros, installées **une seule fois par poste** dans le classeur de macros personnel (`PERSONAL.XLSB`) : `ImporterVBA` (src → classeur actif) et `ExporterVBA` (classeur actif → src). Source en annexe A. |
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

`ExporterVBA` fait l'inverse : chaque composant est écrit dans `src\` (UTF-8 sans BOM, sans en-têtes), et les fichiers de `src\` qui ne correspondent plus à aucun module sont supprimés.

**Pourquoi « remplacer le code » plutôt que « supprimer puis réimporter le fichier »** (la méthode habituelle avec `VBComponents.Import`) :
- un module de document (`ThisWorkbook`, `Feuil1`) ne peut pas être supprimé ;
- un UserForm supprimé perd sa mise en page (le fichier `.frx` binaire) ;
- `Import` lit le fichier en ANSI : les accents d'un fichier UTF-8 seraient abîmés.

## 4. Le cycle de travail

### Mise en place sur un poste (une seule fois)
1. Excel : *Fichier > Options > Centre de gestion de la confidentialité > Paramètres… > Paramètres des macros* → cocher **« Accès approuvé au modèle d'objet du projet VBA »**.
2. Créer le classeur de macros personnel s'il n'existe pas (*Développeur > Enregistrer une macro* → « Classeur de macros personnel », puis arrêter l'enregistrement).
3. `Alt+F11` → clic droit sur `VBAProject (PERSONAL.XLSB)` → *Importer un fichier…* → `outils/VbaSync.bas` → enregistrer.
4. Git : `git clone …`, et de préférence `git config --global core.editor notepad` (évite l'éditeur Vim lors des fusions).

### Récupérer le travail de Claude
1. `git pull` (classeur **fermé**, ou au moins sans modification non enregistrée).
2. Ouvrir le classeur → `Alt+F8` → **ImporterVBA**.
3. Tester.
4. Si c'est bon : enregistrer le classeur, puis `git add -A`, `git commit -m "…"`, `git push`. Le `.xlsm` du dépôt contient alors le même code que `src/`.

### Envoyer ses propres modifications
1. Modifier le code dans l'éditeur VBA (et/ou les feuilles, les UserForms).
2. `Alt+F8` → **ExporterVBA** (met `src/` à jour).
3. Enregistrer le classeur, puis `git add -A`, `git commit -m "…"`, `git push`.

### Revenir en arrière
`git log --oneline` pour trouver la version, puis `git checkout <commit> -- src/ Classeur.xlsm` (ou `git revert <commit>`), et **ImporterVBA**.

## 5. Bonnes pratiques

**Organisation**
- **`src/` fait foi pour le code.** On ne corrige jamais le code directement dans le `.xlsm` sans faire ensuite `ExporterVBA`. Claude ne modifie jamais le `.xlsm` pour changer du code.
- **Une seule personne modifie le classeur à la fois** : Git ne fusionne pas les `.xlsm`. Toujours `git pull` **avant** d'ouvrir le classeur. En cas de conflit sur le `.xlsm` : garder une des deux versions (`git checkout --theirs Classeur.xlsm` ou `--ours`), puis **ImporterVBA** pour réappliquer le code de `src/`.
- **Commiter le `.xlsm` après chaque import réussi** : on garde ainsi des points de retour qui fonctionnent.
- **Travailler par petites itérations** : un objectif, un livrable testable, un commit. Tenir un journal (`docs/iterations.md`) qui sert aussi de point de reprise entre deux sessions de Claude.
- **Aucune donnée sensible dans le dépôt** : pas de mot de passe dans le code, pas de classeur réel contenant des identifiants ou des données personnelles. Les données de travail (fichiers produits par l'outil) se rangent hors du dépôt.

**Écriture du code (pour Claude, qui ne peut pas compiler)**
- `Option Explicit` partout.
- **Caractères** : l'éditeur VBA stocke le code en ANSI (cp1252). Les accents français passent ; tout autre caractère (symboles ✓ ◐ ⊘, flèches…) doit être écrit avec `ChrW(&H2713)`. `outils/VbaSync.bas` lui-même doit rester **en ASCII pur** (il est importé tel quel par l'éditeur VBA).
- **VBA ne distingue pas majuscules et minuscules** : une variable `uiFond` et une fonction `UIFond` provoquent « nom ambigu » ; `wc` et `wC` sont la même variable.
- Paramètres objets passés `ByVal` (évite « type d'argument ByRef incompatible ») ; pas de `CStr` d'un nombre décimal dans une chaîne lue par Excel (virgule décimale en français) : utiliser `CLng` ou `Format$`.
- Ne pas dépendre de la langue d'Excel : `CStr(True)` vaut « Vrai » en français.
- **Avant chaque commit**, Claude relit son code et le passe dans un vérificateur statique (équilibre des blocs `If/For/Do/Select/With`, déclarations en double, identifiants non déclarés, noms ambigus, caractères hors cp1252). Il en écrit un si le projet n'en a pas.
- Les règles métier se testent dans Excel avec une macro de tests (valeurs attendues générées hors Excel quand c'est possible), plutôt qu'à la main.

**UserForms**
- La **mise en page** d'un UserForm (fichier `.frx`) n'est pas dans `src/` : un nouveau UserForm se crée dans Excel, puis son code est géré dans `src/`.
- Alternative pratique : des UserForms **vides** dont tous les contrôles sont créés par le code (`Me.Controls.Add`) ; tout est alors dans `src/`, et une petite macro peut créer le formulaire vide (`VBProject.VBComponents.Add(3)`).
- Dans un événement d'un contrôle, **ne jamais supprimer ce contrôle** (Excel peut planter) : mettre à jour sur place, ou fermer et rouvrir la fenêtre.
- Pas de contrôles ActiveX (TreeView, ProgressBar…) : fragiles en 64 bits et sur postes verrouillés.

## 6. Mise en place dans un nouveau projet (checklist pour Claude)

1. Créer `outils/VbaSync.bas` et `outils/extraire_vba.py` à partir des **annexes A et B** (copie exacte).
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
'   ImporterVBA : src\  -> classeur  (apres un git pull)
'   ExporterVBA : classeur -> src\   (avant un git commit)
'
' Installation : importer ce fichier dans PERSONAL.XLSB (voir README.md).
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
           vbCrLf & vbCrLf & "Pensez a enregistrer le classeur apres vos tests.", _
           vbInformation, "VbaSync"
End Sub

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

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

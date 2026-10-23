Option Explicit

' =============================================================================
' Version de test pour le chef de projet (sans VBA ni Git).
'
'   Alt+F8 > PreparerVersionTest : à lancer par le développeur, une fois, sur le
'            classeur à jour (ImporterVBA, ImporterReferentiel, TesterMoteur faits).
'            Crée la page « Démarrer » (3 boutons) et la feuille « Remarques »,
'            masque les feuilles techniques, puis enregistre une copie à envoyer.
'
' Les boutons appellent SGONouveauDossier, SGOOuvrirDossier et NoterRemarque.
' =============================================================================

Public Const VERSION_TEST As String = "0.7"
Private Const FEUILLE_DEMARRER As String = "Démarrer"
Private Const FEUILLE_REMARQUES As String = "Remarques"

Public Sub PreparerVersionTest()
    Dim ws As Worksheet, manque As String, chemin As Variant, versionRef As String

    ' --- Vérifications
    If Not ReferentielPresent() Then manque = manque & vbCrLf & "- le référentiel (lancez ImporterReferentiel)"
    If Not FormulaireExiste("frmGarde") Or Not FormulaireExiste("frmSaisie") Then
        manque = manque & vbCrLf & "- les formulaires frmGarde et frmSaisie (lancez CreerFormulairesSGO puis ImporterVBA)"
    End If
    If manque <> "" Then
        MsgBox "La version de test ne peut pas être préparée. Il manque :" & manque, vbExclamation, "SGO"
        Exit Sub
    End If
    On Error Resume Next
    versionRef = CStr(ThisWorkbook.Worksheets("R_Version").Range("A2").Value)
    On Error GoTo erreur

    Application.ScreenUpdating = False
    ConstruireDemarrer versionRef
    ConstruireRemarques
    ' Seules « Démarrer » et « Remarques » restent visibles
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> FEUILLE_DEMARRER And ws.Name <> FEUILLE_REMARQUES Then
            If ws.Visible = xlSheetVisible Then ws.Visible = xlSheetHidden
        End If
    Next ws
    ThisWorkbook.Worksheets(FEUILLE_DEMARRER).Activate
    Application.ScreenUpdating = True

    chemin = Application.GetSaveAsFilename(InitialFileName:="SGO_test_CDP_v" & VERSION_TEST & ".xlsm", _
                FileFilter:="Classeur Excel prenant en charge les macros (*.xlsm), *.xlsm", _
                Title:="Enregistrer la version de test (hors du dossier Git)")
    If VarType(chemin) = vbBoolean Then
        MsgBox "Page « Démarrer » créée. Aucune copie enregistrée.", vbInformation, "SGO"
        Exit Sub
    End If
    ThisWorkbook.SaveCopyAs CStr(chemin)
    MsgBox "Version de test enregistrée :" & vbCrLf & chemin & vbCrLf & vbCrLf & _
           "Envoyez-la au chef de projet avec la fiche docs/fiche_test_CDP.pdf.", vbInformation, "SGO"
    Exit Sub
erreur:
    Application.ScreenUpdating = True
    MsgBox "Préparation impossible : " & Err.Description, vbExclamation, "SGO"
End Sub

Private Function FormulaireExiste(ByVal nom As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(nom)
    If Err.Number <> 0 Then
        FormulaireExiste = True          ' accès au projet VBA refusé : on ne peut pas vérifier
    Else
        FormulaireExiste = Not comp Is Nothing
    End If
End Function

' -----------------------------------------------------------------------------
' Page « Démarrer »
' -----------------------------------------------------------------------------
Private Sub ConstruireDemarrer(ByVal versionRef As String)
    Dim ws As Worksheet, y As Single
    Set ws = FeuilleVisible(FEUILLE_DEMARRER, True)
    ws.Unprotect
    ws.Cells.Clear
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Interior.Color = RGB(243, 245, 248)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:L").ColumnWidth = 12

    With ws.Range("A1:L3")
        .Interior.Color = RGB(22, 54, 92)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("B2").Value = "SGO · Dossier de conception de l'offre"
    ws.Range("B2").Font.Size = 18: ws.Range("B2").Font.Bold = True
    ws.Range("J2").Value = "Version de test " & VERSION_TEST
    ws.Range("J2").Font.Color = RGB(200, 216, 234)

    y = ws.Range("B5").Top
    Bouton ws, "btnNouveau", "+   Créer un nouveau dossier", y, "SGONouveauDossier", RGB(0, 99, 177), RGB(255, 255, 255)
    Legende ws, y, "Page de garde, puis saisie des questions."
    y = y + 56
    Bouton ws, "btnOuvrir", "Ouvrir un dossier existant", y, "SGOOuvrirDossier", RGB(255, 255, 255), RGB(33, 37, 41)
    Legende ws, y, "Reprendre un dossier .xlsx déjà créé."
    y = y + 56
    Bouton ws, "btnRemarque", "Noter une remarque", y, "NoterRemarque", RGB(255, 249, 230), RGB(122, 90, 0)
    Legende ws, y, "Vos retours pour l'équipe projet (feuille « Remarques »)."

    ws.Range("B15").Value = "Bon à savoir"
    ws.Range("B15").Font.Bold = True: ws.Range("B15").Font.Color = RGB(22, 54, 92)
    ws.Range("B16").Value = "· À la création, vous choisissez où enregistrer le dossier (un fichier .xlsx par offre)."
    ws.Range("B17").Value = "· Les réponses sont enregistrées automatiquement à chaque changement de question."
    ws.Range("B18").Value = "· Le dossier s'ouvre uniquement avec ce classeur : double-cliquer dessus n'affiche qu'une page d'accueil."
    ws.Range("B19").Value = "· Une copie de sauvegarde est faite à chaque ouverture (sous-dossier « Sauvegardes SGO »)."
    ws.Range("B16:B19").Font.Color = RGB(80, 86, 94)
    ws.Range("B21").Value = "Référentiel version " & versionRef & " · version de test : les réponses saisies servent uniquement à tester l'outil."
    ws.Range("B21").Font.Size = 9: ws.Range("B21").Font.Color = RGB(150, 156, 164)

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Range("A1").Select
    On Error GoTo 0
    ws.Protect DrawingObjects:=True, Contents:=True
    ws.EnableSelection = xlNoSelection
End Sub

Private Sub Bouton(ws As Worksheet, nom As String, texte As String, ByVal y As Single, macro As String, _
                   ByVal fond As Long, ByVal coul As Long)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(5, ws.Range("B1").Left, y, 280, 42)     ' 5 : rectangle à coins arrondis
    s.Name = nom
    s.Fill.ForeColor.RGB = fond
    s.Line.ForeColor.RGB = RGB(170, 178, 189)
    s.Line.Visible = (fond = RGB(255, 255, 255))
    s.Shadow.Visible = False
    With s.TextFrame
        .Characters.Text = texte
        .Characters.Font.Name = "Segoe UI"
        .Characters.Font.Size = 13
        .Characters.Font.Bold = True
        .Characters.Font.Color = coul
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    s.OnAction = macro                           ' sans nom de classeur : la copie envoyée garde des boutons valides
End Sub

Private Sub Legende(ws As Worksheet, ByVal y As Single, texte As String)
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(1, ws.Range("B1").Left + 296, y + 10, 420, 24)
    s.Fill.Visible = False
    s.Line.Visible = False
    s.TextFrame.Characters.Text = texte
    s.TextFrame.Characters.Font.Name = "Segoe UI"
    s.TextFrame.Characters.Font.Size = 11
    s.TextFrame.Characters.Font.Color = RGB(100, 108, 118)
End Sub

' -----------------------------------------------------------------------------
' Feuille « Remarques »
' -----------------------------------------------------------------------------
Private Sub ConstruireRemarques()
    Dim ws As Worksheet, creee As Boolean
    creee = Not FeuilleExiste(FEUILLE_REMARQUES)
    Set ws = FeuilleVisible(FEUILLE_REMARQUES, False)
    If Not creee Then Exit Sub                        ' on ne touche pas aux remarques déjà saisies
    ws.Cells.Font.Name = "Segoe UI"
    ws.Range("A1:G1").Value = Array("Date", "Écran", "Question (El…)", "Remarque", "Gravité", "Réponse de l'équipe", "Traité")
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92)
    End With
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 70
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 40
    ws.Columns("G").ColumnWidth = 8
    ws.Range("A2:G500").WrapText = True
    ws.Range("A2:G500").VerticalAlignment = xlTop
    ws.Range("A2:A500").NumberFormat = "dd/mm/yyyy"
    With ws.Range("B2:B500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
             Formula1:="Page de garde,Écran de saisie,Navigation,Recherche,Enregistrement,Autre"
    End With
    With ws.Range("E2:E500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="Bloquant,Gênant,Suggestion"
    End With
    ws.Range("A2").Select
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' Bouton « Noter une remarque » : nouvelle ligne datée
Public Sub NoterRemarque()
    Dim ws As Worksheet, n As Long
    Set ws = FeuilleVisible(FEUILLE_REMARQUES, False)
    n = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row + 1
    If n < 2 Then n = 2
    ws.Activate
    ws.Cells(n, 1).Value = Date
    ws.Cells(n, 2).Select
End Sub

' -----------------------------------------------------------------------------
Private Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    FeuilleExiste = Not ws Is Nothing
End Function

' Feuille visible (créée si besoin), en première position ou à la suite
Private Function FeuilleVisible(ByVal nom As String, ByVal enPremier As Boolean) As Worksheet
    Dim ws As Worksheet
    If FeuilleExiste(nom) Then
        Set ws = ThisWorkbook.Worksheets(nom)
    ElseIf enPremier Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = nom
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(FEUILLE_DEMARRER))
        ws.Name = nom
    End If
    ws.Visible = xlSheetVisible
    Set FeuilleVisible = ws
End Function

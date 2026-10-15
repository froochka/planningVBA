Option Explicit

' =============================================================================
' Dossier d'une offre : un classeur .xlsx SANS macro (cadrage P1 et P4).
'   Feuille Offre     : page de garde, version du référentiel, dates
'   Feuille Reponses  : une ligne par question renseignée (réponse, différence,
'                       justification, commentaire...)
'   Feuille Archives  : réponses retirées (réinitialisation de la page de garde)
' Le fichier reste lisible sans l'outil. Il est enregistré automatiquement à
' chaque changement de question et à la fermeture de l'écran de saisie.
' =============================================================================

Public gWbDossier As Workbook

Public Const CHAMPS_OFFRE As String = "numero_procedure;libelle_procedure;acheteur;assistant;chef_produit;" & _
    "juriste;pilote;sponsor;direction_achat;standard_direction;standard_univers;version_referentiel;cree_le;modifie_le"
Private Const LIBELLES_OFFRE As String = "N° de procédure;Libellé de la procédure;Nom de l'acheteur;" & _
    "Assistant(e) support achat;Nom du chef produit;Nom du juriste;Pilote de l'offre (DPO / RPO / CPA / Leader);" & _
    "Nom du sponsor;Direction achat;Standard de direction;Standard d'univers;Version du référentiel;Créé le;Modifié le"
Private Const ENTETES_REP As String = "question;libelle;reponse;reponse_saisie;difference;justification;" & _
    "commentaire;standard_modifie;modifie_le"
Private Const NB_COL_REP As Long = 9

' -----------------------------------------------------------------------------
' Création, ouverture
' -----------------------------------------------------------------------------
Public Function CreerDossier(ByVal chemin As String, valeurs As Object) As Boolean
    Dim wb As Workbook, ws As Worksheet, champs As Variant, libs As Variant, i As Long

    On Error GoTo erreur
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Offre"
    ws.Cells.NumberFormat = "@"
    champs = Split(CHAMPS_OFFRE, ";")
    libs = Split(LIBELLES_OFFRE, ";")
    ws.Range("A1:C1").Value = Array("champ", "libelle", "valeur")
    For i = 0 To UBound(champs)
        ws.Cells(i + 2, 1).Value = champs(i)
        ws.Cells(i + 2, 2).Value = libs(i)
        If valeurs.Exists(champs(i)) Then ws.Cells(i + 2, 3).Value = TexteCellule(CStr(valeurs(champs(i))))
    Next i
    MiseEnFormeTableau ws, 3
    ws.Columns(1).Hidden = True

    Set ws = wb.Worksheets.Add(After:=ws)
    ws.Name = "Reponses"
    ws.Cells.NumberFormat = "@"
    ws.Range("A1").Resize(1, NB_COL_REP).Value = Split(ENTETES_REP, ";")
    MiseEnFormeTableau ws, NB_COL_REP

    Set ws = wb.Worksheets.Add(After:=ws)
    ws.Name = "Archives"
    ws.Cells.NumberFormat = "@"
    ws.Range("A1").Resize(1, NB_COL_REP + 2).Value = Split("archive_le;motif;" & ENTETES_REP, ";")
    MiseEnFormeTableau ws, NB_COL_REP + 2

    wb.Worksheets("Offre").Activate
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=chemin, FileFormat:=51        ' xlOpenXMLWorkbook (.xlsx, sans macro)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Set gWbDossier = wb
    DossierEcrire "version_referentiel", gVersionRef
    DossierEcrire "cree_le", Format$(Now, "yyyy-mm-dd hh:nn")
    SauverDossier
    CreerDossier = True
    Exit Function
erreur:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Création du dossier impossible : " & Err.Description, vbExclamation, "SGO"
End Function

Private Sub MiseEnFormeTableau(ws As Worksheet, ByVal nbCol As Long)
    With ws.Range("A1").Resize(1, nbCol)
        .Font.Bold = True
        .Interior.Color = RGB(232, 236, 241)
    End With
    ws.Columns("A:" & Chr(64 + nbCol)).ColumnWidth = 22
End Sub

Public Function OuvrirDossier(ByVal chemin As String) As Boolean
    Dim wb As Workbook, nom As String
    On Error Resume Next
    nom = Mid$(chemin, InStrRev(chemin, "\") + 1)
    Set wb = Workbooks(nom)
    On Error GoTo erreur
    If wb Is Nothing Then Set wb = Workbooks.Open(chemin)
    If Not FeuilleExiste(wb, "Offre") Or Not FeuilleExiste(wb, "Reponses") Then
        MsgBox "Ce classeur n'est pas un dossier SGO (feuilles Offre et Reponses absentes).", vbExclamation, "SGO"
        Exit Function
    End If
    If Not FeuilleExiste(wb, "Archives") Then
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "Archives"
        wb.Worksheets("Archives").Range("A1").Resize(1, NB_COL_REP + 2).Value = _
            Split("archive_le;motif;" & ENTETES_REP, ";")
    End If
    Set gWbDossier = wb
    OuvrirDossier = True
    Exit Function
erreur:
    MsgBox "Ouverture du dossier impossible : " & Err.Description, vbExclamation, "SGO"
End Function

Public Function DossierOuvert() As Boolean
    On Error Resume Next
    If Not gWbDossier Is Nothing Then DossierOuvert = (Len(gWbDossier.Name) > 0)
End Function

Private Function FeuilleExiste(wb As Workbook, nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nom)
    FeuilleExiste = Not ws Is Nothing
End Function

' -----------------------------------------------------------------------------
' Feuille Offre
' -----------------------------------------------------------------------------
Public Function DossierValeur(ByVal champ As String) As String
    Dim r As Range
    Set r = gWbDossier.Worksheets("Offre").Columns(1).Find(What:=champ, LookAt:=xlWhole, MatchCase:=False)
    If Not r Is Nothing Then DossierValeur = CStr(r.Offset(0, 2).Value)
End Function

Public Sub DossierEcrire(ByVal champ As String, ByVal valeur As String)
    Dim ws As Worksheet, r As Range, n As Long
    Set ws = gWbDossier.Worksheets("Offre")
    Set r = ws.Columns(1).Find(What:=champ, LookAt:=xlWhole, MatchCase:=False)
    If r Is Nothing Then
        n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(n, 1).Value = champ
        Set r = ws.Cells(n, 1)
    End If
    r.Offset(0, 2).NumberFormat = "@"
    r.Offset(0, 2).Value = TexteCellule(valeur)
End Sub

' Toutes les valeurs de la page de garde : champ -> valeur
Public Function ValeursOffre() As Object
    Dim d As Object, champs As Variant, i As Long
    Set d = NouveauDico()
    champs = Split(CHAMPS_OFFRE, ";")
    For i = 0 To UBound(champs)
        d(champs(i)) = DossierValeur(CStr(champs(i)))
    Next i
    Set ValeursOffre = d
End Function

' -----------------------------------------------------------------------------
' Feuille Reponses
' -----------------------------------------------------------------------------
' Charge les saisies du dossier dans les questions. Renvoie le nombre de lignes
' dont la question n'existe plus dans le référentiel (elles restent dans le fichier).
Public Function ChargerReponses() As Long
    Dim ws As Worksheet, t As Variant, i As Long, n As Long, q As cQuestionSGO, inconnues As Long

    Set ws = gWbDossier.Worksheets("Reponses")
    n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If n < 2 Then Exit Function
    t = ws.Range("A1").Resize(n, NB_COL_REP).Value
    For i = 2 To n
        If gQuestions.Exists(CStr(t(i, 1))) Then
            Set q = gQuestions(CStr(t(i, 1)))
            q.Saisie = (CStr(t(i, 4)) = "Oui")
            q.ReponseSaisie = CStr(t(i, 3))
            q.Difference = CStr(t(i, 5))
            q.Justification = CStr(t(i, 6))
            q.Commentaire = CStr(t(i, 7))
            q.StandardModifie = (CStr(t(i, 8)) = "Oui")
            If IsDate(t(i, 9)) Then q.ModifieLe = CDate(t(i, 9))
            q.LigneDossier = i
        ElseIf CStr(t(i, 1)) <> "" Then
            inconnues = inconnues + 1
        End If
    Next i
    ChargerReponses = inconnues
End Function

' Écrit la ligne d'une question dans la feuille Reponses (sans enregistrer le fichier)
Public Sub EnregistrerQuestion(ByVal q As cQuestionSGO)
    Dim ws As Worksheet, v As Variant, j As Long
    If q.LigneDossier = 0 And Not q.AQuelqueChose Then Exit Sub
    Set ws = gWbDossier.Worksheets("Reponses")
    If q.LigneDossier = 0 Then q.LigneDossier = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    q.ModifieLe = Now
    v = Array(q.Id, q.Libelle, q.ReponseSaisie, IIf(q.Saisie, "Oui", "Non"), q.Difference, q.Justification, _
              q.Commentaire, IIf(q.StandardModifie, "Oui", "Non"), Format$(q.ModifieLe, "yyyy-mm-dd hh:nn"))
    For j = 0 To UBound(v)
        v(j) = TexteCellule(CStr(v(j)))
    Next j
    With ws.Cells(q.LigneDossier, 1).Resize(1, NB_COL_REP)
        .NumberFormat = "@"
        .Value = v
    End With
End Sub

Public Sub SauverDossier()
    If Not DossierOuvert() Then Exit Sub
    DossierEcrire "modifie_le", Format$(Now, "yyyy-mm-dd hh:nn")
    Application.DisplayAlerts = False
    gWbDossier.Save
    Application.DisplayAlerts = True
End Sub

' Déplace toutes les réponses dans la feuille Archives et remet les questions à
' leur état initial (réponses pré-remplies avec le standard)
Public Sub ArchiverReponses(ByVal motif As String)
    Dim wsR As Worksheet, wsA As Worksheet, n As Long, dest As Long, k As Variant, q As cQuestionSGO

    Set wsR = gWbDossier.Worksheets("Reponses")
    Set wsA = gWbDossier.Worksheets("Archives")
    n = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    If n >= 2 Then
        dest = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row + 1
        wsA.Cells(dest, 1).Resize(n - 1, 2).NumberFormat = "@"
        wsA.Cells(dest, 1).Resize(n - 1, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn")
        wsA.Cells(dest, 2).Resize(n - 1, 1).Value = motif
        wsA.Cells(dest, 3).Resize(n - 1, NB_COL_REP).NumberFormat = "@"
        wsA.Cells(dest, 3).Resize(n - 1, NB_COL_REP).Value = wsR.Range("A2").Resize(n - 1, NB_COL_REP).Value
        wsR.Range("A2").Resize(n - 1, NB_COL_REP).Delete Shift:=xlUp
    End If
    For Each k In gQuestions.Keys
        Set q = gQuestions(k)
        q.Saisie = False: q.ReponseSaisie = "": q.Difference = "": q.Justification = ""
        q.Commentaire = "": q.StandardModifie = False: q.LigneDossier = 0
    Next k
End Sub

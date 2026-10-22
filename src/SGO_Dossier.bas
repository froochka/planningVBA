Option Explicit

' =============================================================================
' Dossier d'une offre : un classeur .xlsx SANS macro (cadrage P1 et P4),
' lu et modifié uniquement par l'outil (décision du 22/10/2026).
'   Feuille Accueil   : seule feuille visible ; résumé de l'offre et message
'                       « Ce fichier s'ouvre avec l'outil SGO »
'   Feuille Offre     : page de garde, version du référentiel, dates   } très
'   Feuille Reponses  : une ligne par question renseignée              } masquées
'   Feuille Archives  : réponses retirées                              }
' La structure du classeur est protégée. C'est une protection contre les
' erreurs (ajout de colonnes, suppression de feuilles...), pas un secret : le
' mot de passe est dans ce code.
' Lecture robuste : les colonnes sont retrouvées par leur en-tête, et un
' contrôle est fait à l'ouverture. Une copie de sauvegarde est faite à chaque
' ouverture (sous-dossier « Sauvegardes SGO », 5 dernières copies).
' =============================================================================

Public gWbDossier As Workbook
Public gRapportOuverture As String       ' anomalies relevées à l'ouverture (affichées par InitialiserSession)

Public Const CHAMPS_OFFRE As String = "numero_procedure;libelle_procedure;acheteur;assistant;chef_produit;" & _
    "juriste;pilote;sponsor;direction_achat;standard_direction;standard_univers;version_referentiel;cree_le;modifie_le;modifie_par"
Private Const LIBELLES_OFFRE As String = "N° de procédure;Libellé de la procédure;Nom de l'acheteur;" & _
    "Assistant(e) support achat;Nom du chef produit;Nom du juriste;Pilote de l'offre (DPO / RPO / CPA / Leader);" & _
    "Nom du sponsor;Direction achat;Standard de direction;Standard d'univers;Version du référentiel;Créé le;Modifié le;Modifié par"
Private Const ENTETES_REP As String = "question;libelle;reponse;reponse_saisie;difference;justification;" & _
    "commentaire;standard_modifie;modifie_le"

Private Const FORMAT_DOSSIER As String = "2"           ' 1 : feuilles visibles (itération 6) ; 2 : protégé
Private Const PROPRIETE As String = "SGO_Dossier"      ' marqueur d'un dossier SGO (propriété du classeur)
Private Const MDP_DOSSIER As String = "SGO-dossier"    ' dissuasif uniquement (voir l'en-tête)
Private Const NB_SAUVEGARDES As Long = 5
Private Const DOSSIER_SAUVEGARDES As String = "Sauvegardes SGO"

' -----------------------------------------------------------------------------
' Création, ouverture
' -----------------------------------------------------------------------------
Public Function CreerDossier(ByVal chemin As String, valeurs As Object) As Boolean
    Dim wb As Workbook, ws As Worksheet, champs As Variant, libs As Variant, i As Long

    On Error GoTo erreur
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "Accueil"

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
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

    Set ws = wb.Worksheets.Add(After:=ws)
    ws.Name = "Reponses"
    ws.Cells.NumberFormat = "@"
    ws.Range("A1").Resize(1, UBound(Split(ENTETES_REP, ";")) + 1).Value = Split(ENTETES_REP, ";")

    Set ws = wb.Worksheets.Add(After:=ws)
    ws.Name = "Archives"
    ws.Cells.NumberFormat = "@"
    ws.Range("A1").Resize(1, UBound(Split(ENTETES_REP, ";")) + 3).Value = Split("archive_le;motif;" & ENTETES_REP, ";")

    Set gWbDossier = wb
    DossierEcrire "version_referentiel", gVersionRef
    DossierEcrire "cree_le", Format$(Now, "yyyy-mm-dd hh:nn")
    MettreAuFormat wb
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=chemin, FileFormat:=51        ' xlOpenXMLWorkbook (.xlsx, sans macro)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SauverDossier
    CreerDossier = True
    Exit Function
erreur:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Création du dossier impossible : " & Err.Description, vbExclamation, "SGO"
End Function

Public Function OuvrirDossier(ByVal chemin As String) As Boolean
    Dim wb As Workbook, nom As String
    gRapportOuverture = ""
    On Error Resume Next
    nom = Mid$(chemin, InStrRev(chemin, "\") + 1)
    Set wb = Workbooks(nom)
    On Error GoTo erreur
    If wb Is Nothing Then Set wb = Workbooks.Open(chemin)
    If Not EstDossierSGO(wb) Then
        MsgBox "Ce classeur n'est pas un dossier SGO.", vbExclamation, "SGO"
        Exit Function
    End If
    Sauvegarder wb
    Application.ScreenUpdating = False
    MettreAuFormat wb                                 ' anciens dossiers : protection ajoutée
    Application.ScreenUpdating = True
    Set gWbDossier = wb
    OuvrirDossier = True
    Exit Function
erreur:
    Application.ScreenUpdating = True
    MsgBox "Ouverture du dossier impossible : " & Err.Description, vbExclamation, "SGO"
End Function

Public Function DossierOuvert() As Boolean
    On Error Resume Next
    If Not gWbDossier Is Nothing Then DossierOuvert = (Len(gWbDossier.Name) > 0)
End Function

' Dossier SGO : marqueur interne, ou feuilles Offre et Reponses (dossiers du format 1)
Public Function EstDossierSGO(wb As Workbook) As Boolean
    Dim p As Object
    On Error Resume Next
    Set p = wb.CustomDocumentProperties(PROPRIETE)
    On Error GoTo 0
    If Not p Is Nothing Then
        EstDossierSGO = True
    Else
        EstDossierSGO = FeuilleExiste(wb, "Offre") And FeuilleExiste(wb, "Reponses")
    End If
End Function

Private Function FeuilleExiste(wb As Workbook, nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nom)
    FeuilleExiste = Not ws Is Nothing
End Function

' Format 2 : feuilles complètes, en-têtes réparés, accueil, feuilles de données
' très masquées, marqueur, structure protégée
Private Sub MettreAuFormat(wb As Workbook)
    Dim ws As Worksheet, nom As Variant, manquants As String
    Deproteger wb
    If Not FeuilleExiste(wb, "Accueil") Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = "Accueil"
    End If
    If Not FeuilleExiste(wb, "Archives") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Archives"
        ws.Cells.NumberFormat = "@"
        ws.Range("A1").Resize(1, UBound(Split(ENTETES_REP, ";")) + 3).Value = Split("archive_le;motif;" & ENTETES_REP, ";")
    End If
    ' En-têtes de Reponses : une colonne manquante est recréée à la fin
    manquants = ReparerEntetes(wb.Worksheets("Reponses"))
    If manquants <> "" Then gRapportOuverture = gRapportOuverture & vbCrLf & _
        "- Colonnes absentes de la feuille Reponses, recréées : " & manquants
    ConstruireAccueil wb
    wb.Worksheets("Accueil").Visible = xlSheetVisible
    For Each nom In Array("Offre", "Reponses", "Archives")
        wb.Worksheets(CStr(nom)).Visible = xlSheetVeryHidden
    Next nom
    Marquer wb
    Proteger wb
End Sub

Private Sub Marquer(wb As Workbook)
    On Error Resume Next
    wb.CustomDocumentProperties(PROPRIETE).Delete
    wb.CustomDocumentProperties.Add Name:=PROPRIETE, LinkToContent:=False, Type:=4, Value:=FORMAT_DOSSIER  ' 4 : texte
End Sub

Private Sub Proteger(wb As Workbook)
    On Error Resume Next
    wb.Protect Password:=MDP_DOSSIER, Structure:=True, Windows:=False
End Sub

Private Sub Deproteger(wb As Workbook)
    On Error Resume Next
    wb.Unprotect Password:=MDP_DOSSIER
End Sub

' -----------------------------------------------------------------------------
' Feuille d'accueil (seule feuille visible)
' -----------------------------------------------------------------------------
Private Sub ConstruireAccueil(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Accueil")
    ws.Unprotect Password:=MDP_DOSSIER
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 11
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C:H").ColumnWidth = 16
    With ws.Range("B2:H3")
        .Interior.Color = RGB(22, 54, 92)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("B2").Value = "Dossier de conception de l'offre (DCO)"
    ws.Range("B2").Font.Size = 18: ws.Range("B2").Font.Bold = True
    ws.Range("B3").Value = "Élaboration de l'offre · outil SGO"
    ws.Range("B3").Font.Color = RGB(200, 216, 234)
    ws.Range("B5:B9").Value = Application.WorksheetFunction.Transpose(Array("Offre", "Direction achat", _
                                  "Standards", "Avancement", "Dernière modification"))
    ws.Range("B5:B9").Font.Color = RGB(100, 108, 118)
    ws.Range("C5:C9").Font.Bold = True
    With ws.Range("B11:H14")
        .Interior.Color = RGB(255, 249, 230)
        .Borders.LineStyle = xlNone
    End With
    ws.Range("B11:B14").Borders(xlEdgeLeft).Color = RGB(214, 158, 0)
    ws.Range("B11:B14").Borders(xlEdgeLeft).Weight = xlThick
    ws.Range("B11").Value = "Ce fichier s'ouvre avec l'outil SGO."
    ws.Range("B11").Font.Bold = True: ws.Range("B11").Font.Size = 13
    ws.Range("B12").Value = "Dans Excel : Alt+F8 › SGO › « Non : ouvrir un dossier existant », puis choisissez ce fichier."
    ws.Range("B13").Value = "Ses données sont protégées : seules les modifications faites dans l'outil sont prises en compte."
    ws.Range("B14").Value = "Pour lire ou diffuser le dossier, utilisez les exports de l'outil (PDF, HTML)."
    ws.Range("B12:B14").Font.Color = RGB(80, 70, 40)
    ws.Range("B16").Font.Size = 9
    ws.Range("B16").Font.Color = RGB(150, 156, 164)
    On Error Resume Next
    wb.Activate
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Range("A1").Select
    On Error GoTo 0
    MettreAJourAccueil wb
End Sub

' Résumé de l'offre, mis à jour à chaque enregistrement
Private Sub MettreAJourAccueil(wb As Workbook)
    Dim ws As Worksheet, a As Variant, num As String, lib As String, avt As String, u As String
    On Error GoTo fin
    Set ws = wb.Worksheets("Accueil")
    ws.Unprotect Password:=MDP_DOSSIER
    num = ValeurDe(wb, "numero_procedure"): lib = ValeurDe(wb, "libelle_procedure")
    ws.Range("C5").Value = TexteCellule(IIf(num = "", "(sans numéro)", num) & IIf(lib = "", "", " · " & lib))
    ws.Range("C6").Value = TexteCellule(ValeurDe(wb, "direction_achat"))
    u = ValeurDe(wb, "standard_univers")
    ws.Range("C7").Value = TexteCellule(LibelleCode(ValeurDe(wb, "standard_direction")) & IIf(u = "", "", " · univers " & LibelleCode(u)))
    avt = "-"
    If Not gChapitres Is Nothing Then
        a = AvancementGlobal()
        If a(1) > 0 Then avt = Int(100 * a(0) / a(1)) & " %  (" & a(0) & " / " & a(1) & " questions traitées)"
    End If
    ws.Range("C8").Value = TexteCellule(avt)
    ws.Range("C9").Value = TexteCellule(ValeurDe(wb, "modifie_le") & IIf(ValeurDe(wb, "modifie_par") = "", "", " par " & ValeurDe(wb, "modifie_par")))
    ws.Range("B16").Value = TexteCellule("Format du dossier " & FORMAT_DOSSIER & " · référentiel version " & ValeurDe(wb, "version_referentiel") & _
                            " · créé le " & ValeurDe(wb, "cree_le"))
fin:
    On Error Resume Next
    ws.Protect Password:=MDP_DOSSIER, DrawingObjects:=True, Contents:=True
    ws.EnableSelection = xlNoSelection
End Sub

' -----------------------------------------------------------------------------
' Sauvegarde : copie du dossier à l'ouverture, 5 dernières conservées
' -----------------------------------------------------------------------------
Private Sub Sauvegarder(wb As Workbook)
    Dim dossier As String, base As String, f As String, noms() As String, n As Long, i As Long, j As Long, t As String
    On Error GoTo erreur
    If wb.Path = "" Or LCase$(Left$(wb.Path, 4)) = "http" Then
        gRapportOuverture = gRapportOuverture & vbCrLf & _
            "- Pas de copie de sauvegarde : le dossier n'est pas dans un dossier du poste (Teams / SharePoint en ligne)."
        Exit Sub
    End If
    dossier = wb.Path & "\" & DOSSIER_SAUVEGARDES
    If Dir(dossier, vbDirectory) = "" Then MkDir dossier
    base = Left$(wb.Name, InStrRev(wb.Name, ".") - 1)
    wb.SaveCopyAs dossier & "\" & base & "_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsx"
    ' Ne garder que les dernières copies (le nom contient la date : l'ordre alphabétique suffit)
    ReDim noms(0 To 0)
    f = Dir(dossier & "\" & base & "_????-??-??_??????.xlsx")
    Do While f <> ""
        If n > 0 Then ReDim Preserve noms(0 To n)
        noms(n) = f
        n = n + 1
        f = Dir()
    Loop
    If n <= NB_SAUVEGARDES Then Exit Sub
    For i = 0 To n - 2                               ' tri croissant (petite liste)
        For j = i + 1 To n - 1
            If noms(j) < noms(i) Then t = noms(i): noms(i) = noms(j): noms(j) = t
        Next j
    Next i
    For i = 0 To n - 1 - NB_SAUVEGARDES
        Kill dossier & "\" & noms(i)
    Next i
    Exit Sub
erreur:
    gRapportOuverture = gRapportOuverture & vbCrLf & "- Copie de sauvegarde impossible : " & Err.Description
End Sub

' -----------------------------------------------------------------------------
' Feuille Offre
' -----------------------------------------------------------------------------
Public Function DossierValeur(ByVal champ As String) As String
    DossierValeur = ValeurDe(gWbDossier, champ)
End Function

Private Function ValeurDe(wb As Workbook, ByVal champ As String) As String
    Dim r As Range
    Set r = wb.Worksheets("Offre").Columns(1).Find(What:=champ, LookAt:=xlWhole, MatchCase:=False)
    If Not r Is Nothing Then ValeurDe = CStr(r.Offset(0, 2).Value)
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
' Feuille Reponses : colonnes retrouvées par leur en-tête
' -----------------------------------------------------------------------------
' En-tête -> numéro de colonne (ligne 1)
Private Function ColonnesRep(ws As Worksheet) As Object
    Dim d As Object, j As Long, n As Long, h As String
    Set d = NouveauDico()
    n = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To n
        h = Trim$(CStr(ws.Cells(1, j).Value))
        If h <> "" And Not d.Exists(h) Then d(h) = j
    Next j
    Set ColonnesRep = d
End Function

' Recrée à la fin les colonnes attendues absentes ; renvoie leurs noms
Private Function ReparerEntetes(ws As Worksheet) As String
    Dim col As Object, h As Variant, n As Long, manquants As String
    Set col = ColonnesRep(ws)
    n = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If n = 1 And Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then n = 0
    For Each h In Split(ENTETES_REP, ";")
        If Not col.Exists(h) Then
            n = n + 1
            ws.Cells(1, n).NumberFormat = "@"
            ws.Cells(1, n).Value = h
            ws.Columns(n).NumberFormat = "@"
            manquants = manquants & IIf(manquants = "", "", ", ") & h
        End If
    Next h
    ReparerEntetes = manquants
End Function

' Charge les saisies du dossier dans les questions et contrôle la feuille.
' Renvoie le nombre de lignes dont la question n'existe plus dans le référentiel
' (elles restent dans le fichier) ; les autres anomalies vont dans gRapportOuverture.
Public Function ChargerReponses() As Long
    Dim ws As Worksheet, col As Object, t As Variant, i As Long, n As Long, nc As Long, q As cQuestionSGO
    Dim inconnues As Long, vus As Object, doublons As String, id As String, cQ As Long

    Set ws = gWbDossier.Worksheets("Reponses")
    Set col = ColonnesRep(ws)
    If Not col.Exists("question") Then
        gRapportOuverture = gRapportOuverture & vbCrLf & "- Feuille Reponses illisible (colonne « question » absente) : aucune réponse chargée."
        Exit Function
    End If
    cQ = col("question")
    n = ws.Cells(ws.Rows.Count, cQ).End(xlUp).Row
    If n < 2 Then Exit Function
    nc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    t = ws.Range("A1").Resize(n, nc).Value
    Set vus = NouveauDico()
    For i = 2 To n
        id = Trim$(CStr(t(i, cQ)))
        If id = "" Then
            ' ligne vide : ignorée
        ElseIf vus.Exists(id) Then
            doublons = doublons & IIf(doublons = "", "", ", ") & id     ' la première ligne est gardée
        ElseIf gQuestions.Exists(id) Then
            vus(id) = True
            Set q = gQuestions(id)
            q.Saisie = (Champ(t, i, col, "reponse_saisie") = "Oui")
            q.ReponseSaisie = Champ(t, i, col, "reponse")
            q.Difference = Champ(t, i, col, "difference")
            q.Justification = Champ(t, i, col, "justification")
            q.Commentaire = Champ(t, i, col, "commentaire")
            q.StandardModifie = (Champ(t, i, col, "standard_modifie") = "Oui")
            If IsDate(Champ(t, i, col, "modifie_le")) Then q.ModifieLe = CDate(Champ(t, i, col, "modifie_le"))
            q.LigneDossier = i
        Else
            vus(id) = True
            inconnues = inconnues + 1
        End If
    Next i
    If doublons <> "" Then gRapportOuverture = gRapportOuverture & vbCrLf & _
        "- Questions présentes plusieurs fois (la première ligne est gardée) : " & doublons
    ChargerReponses = inconnues
End Function

Private Function Champ(t As Variant, ByVal i As Long, col As Object, ByVal nom As String) As String
    If col.Exists(nom) Then
        If Not IsError(t(i, col(nom))) Then Champ = CStr(t(i, col(nom)))
    End If
End Function

' Écrit la ligne d'une question dans la feuille Reponses (sans enregistrer le fichier)
Public Sub EnregistrerQuestion(ByVal q As cQuestionSGO)
    Dim ws As Worksheet, col As Object, noms As Variant, valeurs As Variant, j As Long
    If q.LigneDossier = 0 And Not q.AQuelqueChose Then Exit Sub
    Set ws = gWbDossier.Worksheets("Reponses")
    Set col = ColonnesRep(ws)
    If q.LigneDossier = 0 Then q.LigneDossier = ws.Cells(ws.Rows.Count, col("question")).End(xlUp).Row + 1
    q.ModifieLe = Now
    noms = Split(ENTETES_REP, ";")
    valeurs = Array(q.Id, q.Libelle, q.ReponseSaisie, IIf(q.Saisie, "Oui", "Non"), q.Difference, q.Justification, _
                    q.Commentaire, IIf(q.StandardModifie, "Oui", "Non"), Format$(q.ModifieLe, "yyyy-mm-dd hh:nn"))
    For j = 0 To UBound(noms)
        With ws.Cells(q.LigneDossier, col(noms(j)))
            .NumberFormat = "@"
            .Value = TexteCellule(CStr(valeurs(j)))
        End With
    Next j
End Sub

Public Sub SauverDossier()
    If Not DossierOuvert() Then Exit Sub
    DossierEcrire "modifie_le", Format$(Now, "yyyy-mm-dd hh:nn")
    DossierEcrire "modifie_par", Application.UserName
    MettreAJourAccueil gWbDossier
    Application.DisplayAlerts = False
    gWbDossier.Save
    Application.DisplayAlerts = True
End Sub

' Déplace toutes les réponses dans la feuille Archives et remet les questions à
' leur état initial (réponses pré-remplies avec le standard)
Public Sub ArchiverReponses(ByVal motif As String)
    Dim wsR As Worksheet, wsA As Worksheet, colR As Object, colA As Object, n As Long, dest As Long
    Dim h As Variant, i As Long, k As Variant, q As cQuestionSGO

    Set wsR = gWbDossier.Worksheets("Reponses")
    Set wsA = gWbDossier.Worksheets("Archives")
    Set colR = ColonnesRep(wsR)
    ReparerEntetes wsA
    Set colA = ColonnesRep(wsA)
    If Not colA.Exists("archive_le") Or Not colA.Exists("motif") Then
        ' Archives d'un ancien format : on ajoute les deux colonnes à la fin
        For Each h In Array("archive_le", "motif")
            If Not colA.Exists(h) Then wsA.Cells(1, wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column + 1).Value = h
        Next h
        Set colA = ColonnesRep(wsA)
    End If
    n = wsR.Cells(wsR.Rows.Count, colR("question")).End(xlUp).Row
    If n >= 2 Then
        dest = wsA.Cells(wsA.Rows.Count, colA("question")).End(xlUp).Row + 1
        For i = 2 To n
            wsA.Cells(dest, colA("archive_le")).Value = TexteCellule(Format$(Now, "yyyy-mm-dd hh:nn"))
            wsA.Cells(dest, colA("motif")).Value = TexteCellule(motif)
            For Each h In Split(ENTETES_REP, ";")
                wsA.Cells(dest, colA(h)).NumberFormat = "@"
                wsA.Cells(dest, colA(h)).Value = TexteCellule(CStr(wsR.Cells(i, colR(h)).Value))
            Next h
            dest = dest + 1
        Next i
        wsR.Range(wsR.Rows(2), wsR.Rows(n)).Delete
    End If
    For Each k In gQuestions.Keys
        Set q = gQuestions(k)
        q.Saisie = False: q.ReponseSaisie = "": q.Difference = "": q.Justification = ""
        q.Commentaire = "": q.StandardModifie = False: q.LigneDossier = 0
    Next k
End Sub

Option Explicit

' =============================================================================
' Référentiel du SGO : questions, structure, standards, conditions, directions.
'
'   Alt+F8 > ImporterReferentiel : copie les fichiers referentiel\R_*.csv (à côté
'            du classeur) dans des feuilles cachées R_... du classeur. À relancer
'            après chaque mise à jour du référentiel (git pull).
'
' ChargerReferentiel lit ensuite ces feuilles et prépare les données en mémoire
' pour le moteur et les écrans. Dans le futur complément SGO.xlam, les feuilles
' R_... seront embarquées dans le complément (cadrage P2).
' =============================================================================

Public Const SEP_CHOIX As String = "|"
Private Const DOSSIER_REF As String = "referentiel"
Private Const FEUILLES_REF As String = "R_Version;R_Categories;R_Questions;R_Conditions;R_Standards;R_Directions"

Public gRefCharge As Boolean
Public gVersionRef As String
Public gQuestions As Object          ' id -> cQuestionSGO (toutes, dans l'ordre ; sans casse)
Public gOrdre As Collection          ' cQuestionSGO hors questions système, dans l'ordre
Public gCategories As Object         ' id -> cCategorieSGO
Public gChapitres As Collection      ' cCategorieSGO de type CHAPITRE, dans l'ordre
Public gStd As Object                ' "question|type|cible" -> Array(niveau, valeur)
Public gDirStd As Object             ' standard de direction -> Array(a des standards, libellé)
Public gUnivers As Object            ' univers -> Array(standard de direction parent, a des standards)
Public gDirAchat As Object           ' direction achat -> libellé
Public gConditions As Object         ' fille (question ou catégorie) -> Collection de Array(mère, opérateur, critère)
Public gFilles As Object             ' mère -> texte « El78a, El78b » (questions et catégories conditionnées)

' -----------------------------------------------------------------------------
' Import des CSV dans les feuilles R_...
' -----------------------------------------------------------------------------
Public Sub ImporterReferentiel()
    Dim dossier As String, noms As Variant, i As Long, t As Variant, ws As Worksheet
    Dim bilan As String, feuilleActive As Object

    dossier = ThisWorkbook.Path & "\" & DOSSIER_REF
    If Dir(dossier & "\R_Questions.csv") = "" Then
        MsgBox "Fichiers du référentiel introuvables dans :" & vbCrLf & dossier & vbCrLf & vbCrLf & _
               "Faites un git pull, puis relancez la macro.", vbExclamation, "SGO"
        Exit Sub
    End If
    On Error GoTo erreur
    Application.ScreenUpdating = False
    Set feuilleActive = ActiveSheet
    noms = Split(FEUILLES_REF, ";")
    For i = 0 To UBound(noms)
        t = LireCsv(dossier & "\" & noms(i) & ".csv")
        Set ws = FeuilleTechnique(ThisWorkbook, CStr(noms(i)))
        ws.Cells.Clear
        EcrireTableau ws, t
        bilan = bilan & vbCrLf & "- " & noms(i) & " : " & (UBound(t, 1) - 1) & " lignes"
    Next i
    If Not feuilleActive Is Nothing Then feuilleActive.Activate
    Application.ScreenUpdating = True
    gRefCharge = False
    MsgBox "Référentiel importé dans le classeur :" & bilan & vbCrLf & vbCrLf & _
           "Pensez à enregistrer le classeur.", vbInformation, "SGO"
    Exit Sub
erreur:
    Application.ScreenUpdating = True
    MsgBox "Import du référentiel impossible : " & Err.Description, vbExclamation, "SGO"
End Sub

' Feuille cachée du classeur, créée si besoin
Public Function FeuilleTechnique(wb As Workbook, nom As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nom
    End If
    If wb Is ThisWorkbook Then ws.Visible = xlSheetHidden
    Set FeuilleTechnique = ws
End Function

' Écrit un tableau 2D (1..n, 1..m) en texte, à partir de A1
Public Sub EcrireTableau(ws As Worksheet, t As Variant)
    Dim i As Long, j As Long, v As String
    For i = 1 To UBound(t, 1)
        For j = 1 To UBound(t, 2)
            v = CStr(t(i, j))
            t(i, j) = TexteCellule(v)
        Next j
    Next i
    With ws.Range("A1").Resize(UBound(t, 1), UBound(t, 2))
        .NumberFormat = "@"
        .Value = t
    End With
End Sub

' Une valeur commençant par = + - @ serait lue comme une formule : on la protège
Public Function TexteCellule(ByVal v As String) As String
    If Len(v) > 0 Then
        If InStr("=+-@", Left$(v, 1)) > 0 Then v = "'" & v
    End If
    TexteCellule = v
End Function

' -----------------------------------------------------------------------------
' Lecture d'un CSV UTF-8 (séparateur « ; », guillemets, retours à la ligne
' dans les champs) -> tableau 2D (1..lignes, 1..colonnes)
' -----------------------------------------------------------------------------
Public Function LireCsv(chemin As String) As Variant
    Dim s As String, n As Long, p As Long, q As Long, debut As Long
    Dim ch As String, champ As String
    Dim lignes As Collection, ligne As Collection

    s = LireTexteUtf8(chemin)
    If Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    n = Len(s): p = 1
    Set lignes = New Collection
    Set ligne = New Collection
    Do While p <= n
        If Mid$(s, p, 1) = """" Then
            ' Champ entre guillemets : "" vaut un guillemet
            debut = p + 1: champ = ""
            Do
                q = InStr(debut, s, """")
                If q = 0 Then q = n + 1
                champ = champ & Mid$(s, debut, q - debut)
                If Mid$(s, q + 1, 1) = """" Then
                    champ = champ & """"
                    debut = q + 2
                Else
                    p = q + 1
                    Exit Do
                End If
            Loop
        Else
            q = p
            Do While q <= n
                ch = Mid$(s, q, 1)
                If ch = ";" Or ch = vbCr Or ch = vbLf Then Exit Do
                q = q + 1
            Loop
            champ = Mid$(s, p, q - p)
            p = q
        End If
        ligne.Add champ
        ch = Mid$(s, p, 1)
        If ch = ";" Then
            p = p + 1
            If p > n Then ligne.Add ""
        Else
            If ch = vbCr Then p = p + 1
            If Mid$(s, p, 1) = vbLf Then p = p + 1
            lignes.Add ligne
            Set ligne = New Collection
        End If
    Loop
    If ligne.Count > 0 Then lignes.Add ligne
    LireCsv = VersTableau(lignes)
End Function

Private Function VersTableau(lignes As Collection) As Variant
    Dim t() As Variant, nCol As Long, nLig As Long, ligne As Collection, j As Long
    For Each ligne In lignes
        If ligne.Count > nCol Then nCol = ligne.Count
    Next ligne
    ReDim t(1 To lignes.Count, 1 To nCol)
    For Each ligne In lignes
        If Not (ligne.Count = 1 And ligne(1) = "") Then      ' ignore les lignes vides
            nLig = nLig + 1
            For j = 1 To nCol
                If j <= ligne.Count Then t(nLig, j) = ligne(j) Else t(nLig, j) = ""
            Next j
        End If
    Next ligne
    VersTableau = Redimensionner(t, nLig, nCol)
End Function

Private Function Redimensionner(t As Variant, nLig As Long, nCol As Long) As Variant
    Dim r() As Variant, i As Long, j As Long
    ReDim r(1 To nLig, 1 To nCol)
    For i = 1 To nLig
        For j = 1 To nCol
            r(i, j) = t(i, j)
        Next j
    Next i
    Redimensionner = r
End Function

Public Function LireTexteUtf8(chemin As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' texte
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile chemin
    LireTexteUtf8 = st.ReadText(-1)
    st.Close
End Function

' -----------------------------------------------------------------------------
' Chargement en mémoire
' -----------------------------------------------------------------------------
Public Sub ChargerReferentiel()
    Dim t As Variant, col As Object, i As Long, cle As String
    Dim q As cQuestionSGO, c As cCategorieSGO, k As Variant, cnd As Collection

    Set gQuestions = NouveauDico()
    Set gCategories = NouveauDico()
    Set gStd = NouveauDico()
    Set gDirStd = NouveauDico()
    Set gUnivers = NouveauDico()
    Set gDirAchat = NouveauDico()
    Set gConditions = NouveauDico()
    Set gFilles = NouveauDico()
    Set gOrdre = New Collection
    Set gChapitres = New Collection

    ' Version
    t = TableauFeuille("R_Version")
    gVersionRef = CStr(t(2, 1))

    ' Catégories (dans l'ordre du fichier)
    t = TableauFeuille("R_Categories"): Set col = Colonnes(t)
    For i = 2 To UBound(t, 1)
        Set c = New cCategorieSGO
        c.Id = t(i, col("id"))
        c.TypeCat = t(i, col("type"))
        c.Libelle = t(i, col("libelle"))
        c.Parent = t(i, col("parent"))
        c.Ordre = Val(t(i, col("ordre")))
        c.Visible = True
        gCategories.Add c.Id, c
    Next i
    For Each k In gCategories.Keys
        Set c = gCategories(k)
        If c.Parent = "" Then
            gChapitres.Add c
        Else
            gCategories(c.Parent).Enfants.Add c
        End If
    Next k

    ' Questions
    t = TableauFeuille("R_Questions"): Set col = Colonnes(t)
    For i = 2 To UBound(t, 1)
        Set q = New cQuestionSGO
        q.Id = t(i, col("id"))
        q.Libelle = t(i, col("libelle"))
        q.SousTheme = t(i, col("sous_theme"))
        q.Ordre = Val(t(i, col("ordre")))
        q.Mode = t(i, col("mode"))
        If t(i, col("choix")) = "" Then
            q.Choix = Array()
        Else
            q.Choix = Split(t(i, col("choix")), SEP_CHOIX)
        End If
        q.Operateur = t(i, col("operateur"))
        q.AlimenteCCAP = (t(i, col("alimente_ccap")) = "Oui")
        q.Systeme = (t(i, col("systeme")) = "Oui")
        If t(i, col("active")) = "Oui" Then
            gQuestions.Add q.Id, q
            If Not q.Systeme Then
                gOrdre.Add q
                gCategories(q.SousTheme).Questions.Add q
            End If
        End If
    Next i

    ' Standards : une ligne par question et par cible
    t = TableauFeuille("R_Standards"): Set col = Colonnes(t)
    For i = 2 To UBound(t, 1)
        cle = LCase$(t(i, col("question")) & "|" & t(i, col("type")) & "|" & t(i, col("cible")))
        gStd(cle) = Array(CStr(t(i, col("niveau"))), CStr(t(i, col("valeur"))))
    Next i

    ' Directions et univers
    t = TableauFeuille("R_Directions"): Set col = Colonnes(t)
    For i = 2 To UBound(t, 1)
        Select Case t(i, col("type"))
            Case "ACHAT"
                gDirAchat(CStr(t(i, col("code")))) = CStr(t(i, col("libelle")))
            Case "STANDARD"
                gDirStd(CStr(t(i, col("code")))) = (t(i, col("a_standards")) = "Oui")
            Case "UNIVERS"
                gUnivers(CStr(t(i, col("code")))) = Array(CStr(t(i, col("parent"))), (t(i, col("a_standards")) = "Oui"))
        End Select
    Next i

    ' Conditions (une fille peut en avoir plusieurs : toutes doivent être vraies)
    t = TableauFeuille("R_Conditions"): Set col = Colonnes(t)
    For i = 2 To UBound(t, 1)
        cle = LCase$(t(i, col("fille")))
        If Not gConditions.Exists(cle) Then gConditions.Add cle, New Collection
        Set cnd = gConditions(cle)
        cnd.Add Array(LCase$(t(i, col("mere"))), CStr(t(i, col("operateur"))), CStr(t(i, col("critere"))))
        cle = LCase$(t(i, col("mere")))
        If gFilles.Exists(cle) Then
            gFilles(cle) = gFilles(cle) & ", " & t(i, col("fille"))
        Else
            gFilles(cle) = CStr(t(i, col("fille")))
        End If
    Next i

    gRefCharge = True
End Sub

' Vrai si les feuilles du référentiel sont présentes dans le classeur
Public Function ReferentielPresent() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("R_Questions")
    ReferentielPresent = Not ws Is Nothing
End Function

Public Function NouveauDico() As Object
    Set NouveauDico = CreateObject("Scripting.Dictionary")
    NouveauDico.CompareMode = vbTextCompare
End Function

Private Function TableauFeuille(nom As String) As Variant
    Dim ws As Worksheet, t As Variant, i As Long, j As Long
    Set ws = ThisWorkbook.Worksheets(nom)
    t = ws.Range("A1", ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, _
                               ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)).Value
    For i = 1 To UBound(t, 1)
        For j = 1 To UBound(t, 2)
            If IsEmpty(t(i, j)) Or IsError(t(i, j)) Then t(i, j) = "" Else t(i, j) = CStr(t(i, j))
        Next j
    Next i
    TableauFeuille = t
End Function

' En-têtes de la ligne 1 -> numéro de colonne
Private Function Colonnes(t As Variant) As Object
    Dim d As Object, j As Long
    Set d = NouveauDico()
    For j = 1 To UBound(t, 2)
        d(CStr(t(1, j))) = j
    Next j
    Set Colonnes = d
End Function

' Libellé d'affichage d'un code (DAG_SOFI -> DAG SOFI)
Public Function LibelleCode(ByVal code As String) As String
    LibelleCode = Replace(code, "_", " ")
End Function

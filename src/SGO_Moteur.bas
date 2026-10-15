Option Explicit

' =============================================================================
' Moteur de règles du SGO : standard applicable, réponse pré-remplie,
' conformité, questions conditionnelles, statuts et avancement.
'
' Les règles de standard et de conformité reprennent exactement les formules
' de l'outil d'origine (feuille 150-Modele SGO) : voir docs/analyse_existant.md.
' Copie conforme : outils/moteur.py (utilisé pour la preuve d'équivalence).
' Toute règle modifiée ici doit l'être aussi dans moteur.py.
' =============================================================================

' Textes de l'outil d'origine (repris dans les synthèses et l'extraction GINO)
Public Const MSG_STANDARD_INCOMPLET As String = "Standards incomplets pour cette question, contacter votre administrateur!"
Public Const MSG_REPONSE_NON_SAISIE As String = "Renseigner la réponse DCO"
Public Const MSG_STANDARD_NON_APPLICABLE As String = "Standard non applicable pour cette offre au regard des réponses précédentes apportées"
Public Const NC_OFFRE As String = "Non concerné pour l'offre"
Public Const NC_DIRECTION As String = "Non concerné pour la direction/univers"
Public Const CONF_CONFORME As String = "Conforme au standard"
Public Const CONF_NON_CONFORME As String = "Non conforme au standard"
Public Const CONF_SPECIFIQUE As String = "Spécifique Offre"
Public Const CONF_NON_CONCERNE As String = "Non concerné"
Public Const DIFF_IDENTIQUE As String = "Identique"
Public Const DIFF_DIFFERENTE As String = "Différente"
Public Const DIFF_NC As String = "Non concerné"

' Statuts d'une question
Public Const ST_MASQUEE As Long = 0
Public Const ST_NON_REPONDUE As Long = 1
Public Const ST_CONFORME As Long = 2
Public Const ST_SPECIFIQUE As Long = 3
Public Const ST_A_JUSTIFIER As Long = 4
Public Const ST_JUSTIFIEE As Long = 5
Public Const ST_NON_CONCERNEE As Long = 6
Public Const ST_INCOMPLET As Long = 7

' Contexte de l'offre (page de garde)
Public gDirection As String          ' standard de direction (DAV, DAG_SOFI...)
Public gUniversChoisi As String      ' standard d'univers (VI, Logiciels...)

Private gVisibles As Object          ' id -> Boolean, calculé par EvaluerDossier

' -----------------------------------------------------------------------------
' Standard
' -----------------------------------------------------------------------------
' Niveau de standard : E, D, U, Sp, Titre, ou "" (standards incomplets)
Public Function NiveauStd(ByVal q As String, ByVal d As String, ByVal u As String) As String
    Dim n As String, vu As String
    If Not gDirStd.Exists(d) Then
        NiveauStd = "Sp": Exit Function
    End If
    If Not gDirStd(d) Then
        NiveauStd = "Sp": Exit Function             ' direction sans colonne de standards
    End If
    n = ValeurTable(q, "D", d, 0)
    If n = "D" Or n = "U" Then
        If UniversAStandards(u) Then
            vu = ValeurTable(q, "U", u, 1)
            If LCase$(vu) = "sp" Then n = "Sp"
        End If
    End If
    NiveauStd = n
End Function

Public Function ValeurStd(ByVal q As String, ByVal d As String, ByVal u As String, ByVal niv As String) As String
    Select Case niv
        Case "E": ValeurStd = ValeurTable(q, "E", "", 1)
        Case "D": ValeurStd = ValeurTable(q, "D", d, 1)
        Case "U": ValeurStd = ValeurTable(q, "U", u, 1)
        Case Else: ValeurStd = ""
    End Select
End Function

' Valeur brute de R_Standards (i = 0 : niveau, 1 : valeur)
Public Function ValeurTable(ByVal q As String, ByVal typ As String, ByVal cible As String, ByVal i As Long) As String
    Dim cle As String
    cle = LCase$(q & "|" & typ & "|" & cible)
    If gStd.Exists(cle) Then ValeurTable = gStd(cle)(i)
End Function

Private Function UniversAStandards(ByVal u As String) As Boolean
    If gUnivers.Exists(u) Then UniversAStandards = gUnivers(u)(1)
End Function

' Réponse pré-remplie au chargement (décision ES-6)
Public Function ReponseInitiale(ByVal niv As String, ByVal std As String, ByVal mode As String) As String
    If InStr(1, mode, "annex", vbTextCompare) > 0 Then
        If LCase$(Left$(mode, 7)) = "annexe " Then
            ReponseInitiale = Mid$(mode, 8)
        Else
            ReponseInitiale = mode
        End If
    ElseIf niv = "E" Or niv = "D" Or niv = "U" Then
        ReponseInitiale = std
    End If
End Function

' Conformité : formule de l'outil d'origine (colonne « Conformité au standard »)
Public Function CalculerConformite(ByVal niv As String, ByVal std As String, ByVal rep As String, _
                                   ByVal mode As String, ByVal op As String) As String
    Dim annexe As Boolean
    annexe = (InStr(1, mode, "annex", vbTextCompare) > 0)
    If (niv = "Titre" Or LCase$(Left$(rep, 11)) = "non concern" Or annexe) And niv <> "Sp" Then
        If std = "" Or LCase$(Left$(std, 11)) = "non concern" Then
            CalculerConformite = CONF_NON_CONCERNE
        Else
            CalculerConformite = CONF_NON_CONFORME
        End If
    ElseIf niv = "Sp" And rep <> "" Then
        CalculerConformite = CONF_SPECIFIQUE
    ElseIf (niv = "" Or std = "") And niv <> "Sp" Then
        CalculerConformite = MSG_STANDARD_INCOMPLET
    ElseIf rep = "" Then
        CalculerConformite = MSG_REPONSE_NON_SAISIE
    ElseIf EstConforme(rep, std, op) Then
        CalculerConformite = CONF_CONFORME
    ElseIf rep = MSG_STANDARD_NON_APPLICABLE Then
        CalculerConformite = CONF_NON_CONCERNE
    Else
        CalculerConformite = CONF_NON_CONFORME
    End If
End Function

' Opérateur « = » : égalité (sans tenir compte des majuscules, comme Excel) ;
' sinon la réponse doit contenir le standard
Private Function EstConforme(ByVal rep As String, ByVal std As String, ByVal op As String) As Boolean
    If op = "=" Then
        EstConforme = (StrComp(rep, std, vbTextCompare) = 0)
    Else
        EstConforme = (InStr(1, rep, std, vbTextCompare) > 0)
    End If
End Function

' Standard qui rend la question sans objet : « Non concerné pour la direction/univers »
' ou « ... non applicable ... ». Renvoie le motif du masquage, ou "".
Public Function MasqueeParStandard(ByVal std As String) As String
    If InStr(1, std, "direction/", vbTextCompare) > 0 Then
        MasqueeParStandard = NC_DIRECTION
    ElseIf InStr(1, std, "non applicable", vbTextCompare) > 0 Then
        MasqueeParStandard = MSG_STANDARD_NON_APPLICABLE
    End If
End Function

' -----------------------------------------------------------------------------
' Conditions : « <> » = la réponse contient le critère ; « = » = égale au critère
' -----------------------------------------------------------------------------
Public Function ConditionVraie(ByVal op As String, ByVal critere As String, ByVal rep As String) As Boolean
    If op = "=" Then
        ConditionVraie = (StrComp(rep, critere, vbTextCompare) = 0)
    Else
        ConditionVraie = (InStr(1, rep, critere, vbTextCompare) > 0)
    End If
End Function

' -----------------------------------------------------------------------------
' Évaluation complète du dossier (après chargement ou après chaque réponse)
' -----------------------------------------------------------------------------
Public Sub EvaluerDossier()
    Dim k As Variant, q As cQuestionSGO

    Set gVisibles = NouveauDico()
    ' 1. Standard et réponse en vigueur de chaque question
    For Each k In gQuestions.Keys
        Set q = gQuestions(k)
        If q.Systeme Then
            If LCase$(q.Id) = "el393" Then q.Reponse = gDirection Else q.Reponse = gUniversChoisi
        Else
            q.Niveau = NiveauStd(q.Id, gDirection, gUniversChoisi)
            q.Standard = ValeurStd(q.Id, gDirection, gUniversChoisi, q.Niveau)
            If q.Saisie Then
                q.Reponse = q.ReponseSaisie
            Else
                q.Reponse = ReponseInitiale(q.Niveau, q.Standard, q.Mode)
            End If
        End If
    Next k
    ' 2. Visibilité (dans l'ordre du référentiel), conformité et statut
    For Each k In gQuestions.Keys
        Set q = gQuestions(k)
        If q.Systeme Then
            q.Motif = ""
        Else
            q.Motif = MasqueeParStandard(q.Standard)
            If q.Motif = "" Then
                If Not ElementVisible(q.Id, "") Then q.Motif = NC_OFFRE
            End If
        End If
        gVisibles(LCase$(q.Id)) = (q.Motif = "")
        If q.Motif <> "" Or q.Systeme Then
            q.Conformite = ""
            q.Statut = ST_MASQUEE
        Else
            q.Conformite = CalculerConformite(q.Niveau, q.Standard, q.Reponse, q.Mode, q.Operateur)
            q.Statut = CalculerStatut(q)
        End If
    Next k
    CalculerAvancement
End Sub

' Visible si toutes ses conditions sont vraies (avec des questions mères visibles)
' et si sa catégorie parente est visible (cadrage P5). Une boucle rend masqué.
Private Function ElementVisible(ByVal id As String, ByVal pile As String) As Boolean
    Dim cle As String, cnd As Variant, mere As cQuestionSGO, parent As String, estCategorie As Boolean

    cle = LCase$(id)
    estCategorie = gCategories.Exists(cle)
    If estCategorie And gVisibles.Exists(cle) Then
        ElementVisible = gVisibles(cle)
        Exit Function
    End If
    If InStr(pile, "|" & cle & "|") > 0 Then Exit Function
    pile = pile & "|" & cle & "|"

    If gConditions.Exists(cle) Then
        For Each cnd In gConditions(cle)
            If Not gQuestions.Exists(cnd(0)) Then GoTo fin
            Set mere = gQuestions(cnd(0))
            If Not QuestionVisible(mere, pile) Then GoTo fin
            If Not ConditionVraie(cnd(1), cnd(2), mere.Reponse) Then GoTo fin
        Next cnd
    End If
    If estCategorie Then parent = gCategories(cle).Parent Else parent = gQuestions(cle).SousTheme
    If parent = "" Then
        ElementVisible = True
    Else
        ElementVisible = ElementVisible(parent, pile)
    End If
fin:
    If estCategorie Then gVisibles(cle) = ElementVisible
End Function

Private Function QuestionVisible(ByVal q As cQuestionSGO, ByVal pile As String) As Boolean
    Dim cle As String
    cle = LCase$(q.Id)
    If gVisibles.Exists(cle) Then
        QuestionVisible = gVisibles(cle)
    ElseIf q.Systeme Then
        QuestionVisible = True
    ElseIf MasqueeParStandard(q.Standard) <> "" Then
        QuestionVisible = False
    Else
        QuestionVisible = ElementVisible(q.Id, pile)
    End If
End Function

' Différence avec l'offre précédente : « Identique » par défaut pour les
' questions spécifiques (Sp), « Non concerné » sinon (outil d'origine)
Public Function DifferenceEnVigueur(ByVal q As cQuestionSGO) As String
    If q.Difference <> "" Then
        DifferenceEnVigueur = q.Difference
    ElseIf q.Niveau = "Sp" Then
        DifferenceEnVigueur = DIFF_IDENTIQUE
    Else
        DifferenceEnVigueur = DIFF_NC
    End If
End Function

' Statut d'une question visible (les 4 cas du mode opératoire)
Public Function CalculerStatut(ByVal q As cQuestionSGO) As Long
    If q.Conformite = MSG_STANDARD_INCOMPLET Then
        CalculerStatut = ST_INCOMPLET
    ElseIf Trim$(q.Reponse) = "" Then
        CalculerStatut = ST_NON_REPONDUE
    ElseIf q.Conformite = CONF_NON_CONFORME Then
        CalculerStatut = IIf(Trim$(q.Justification) = "", ST_A_JUSTIFIER, ST_JUSTIFIEE)       ' cas 2
    ElseIf q.Conformite = CONF_SPECIFIQUE Then
        If DifferenceEnVigueur(q) = DIFF_DIFFERENTE Then
            CalculerStatut = IIf(Trim$(q.Justification) = "", ST_A_JUSTIFIER, ST_JUSTIFIEE)   ' cas 4
        Else
            CalculerStatut = ST_SPECIFIQUE                                                    ' cas 3
        End If
    ElseIf q.Conformite = CONF_NON_CONCERNE Then
        CalculerStatut = ST_NON_CONCERNEE
    Else
        CalculerStatut = ST_CONFORME                                                          ' cas 1
    End If
End Function

Public Function EstTraitee(ByVal q As cQuestionSGO) As Boolean
    Select Case q.Statut
        Case ST_CONFORME, ST_SPECIFIQUE, ST_JUSTIFIEE, ST_NON_CONCERNEE
            EstTraitee = True
    End Select
End Function

' Avancement de chaque catégorie (questions visibles uniquement)
Private Sub CalculerAvancement()
    Dim c As Variant
    For Each c In gChapitres
        CompterCategorie c, True
    Next c
End Sub

Private Sub CompterCategorie(ByVal c As cCategorieSGO, ByVal parentVisible As Boolean)
    Dim e As Variant, q As Variant
    c.Visible = parentVisible And ElementVisible(c.Id, "")
    c.NbVisibles = 0: c.NbTraitees = 0: c.NbAJustifier = 0
    For Each q In c.Questions
        If q.Statut <> ST_MASQUEE Then
            c.NbVisibles = c.NbVisibles + 1
            If EstTraitee(q) Then c.NbTraitees = c.NbTraitees + 1
            If q.Statut = ST_A_JUSTIFIER Then c.NbAJustifier = c.NbAJustifier + 1
        End If
    Next q
    For Each e In c.Enfants
        CompterCategorie e, c.Visible
        c.NbVisibles = c.NbVisibles + e.NbVisibles
        c.NbTraitees = c.NbTraitees + e.NbTraitees
        c.NbAJustifier = c.NbAJustifier + e.NbAJustifier
    Next e
End Sub

' Avancement global : Array(traitées, visibles)
Public Function AvancementGlobal() As Variant
    Dim c As Variant, nT As Long, nV As Long
    For Each c In gChapitres
        nT = nT + c.NbTraitees
        nV = nV + c.NbVisibles
    Next c
    AvancementGlobal = Array(nT, nV)
End Function

' Pourquoi un élément est masqué : texte court pour l'affichage (info-bulles)
Public Function RaisonMasquage(ByVal id As String) As String
    Dim cle As String, cnd As Variant, txt As String, q As cQuestionSGO
    cle = LCase$(id)
    If gQuestions.Exists(cle) Then
        Set q = gQuestions(cle)
        If q.Motif = NC_DIRECTION Then
            RaisonMasquage = "Sans standard pour votre direction/univers"
            Exit Function
        ElseIf q.Motif = MSG_STANDARD_NON_APPLICABLE Then
            RaisonMasquage = "Standard non applicable pour votre direction/univers"
            Exit Function
        End If
    End If
    If gConditions.Exists(cle) Then
        For Each cnd In gConditions(cle)
            If txt <> "" Then txt = txt & " et de "
            txt = txt & gQuestions(cnd(0)).Id
        Next cnd
        RaisonMasquage = "Dépend de la réponse à " & txt
    ElseIf gQuestions.Exists(cle) Then
        RaisonMasquage = "Sa rubrique est sans objet : " & RaisonMasquage(gQuestions(cle).SousTheme)
    ElseIf gCategories.Exists(cle) Then
        If gCategories(cle).Parent <> "" Then
            If Not gCategories(gCategories(cle).Parent).Visible Then
                RaisonMasquage = RaisonMasquage(gCategories(cle).Parent)
                Exit Function
            End If
        End If
        RaisonMasquage = "Aucune question applicable"
    End If
End Function

' Pourquoi un élément est affiché, s'il dépend d'une condition
Public Function RaisonAffichage(ByVal id As String) As String
    Dim cnd As Variant, txt As String, op As String
    If Not gConditions.Exists(LCase$(id)) Then Exit Function
    For Each cnd In gConditions(LCase$(id))
        If txt <> "" Then txt = txt & " et "
        If cnd(1) = "=" Then op = " vaut « " Else op = " contient « "
        txt = txt & gQuestions(cnd(0)).Id & op & cnd(2) & " »"
    Next cnd
    RaisonAffichage = "affichée car " & txt
End Function

Option Explicit

' =============================================================================
' Écran de saisie du DCO : vraies questions, vrais standards, moteur branché.
' Formulaire vide dans le classeur : tout est construit par ce code, selon la
' taille de l'écran et la taille de texte choisie (A- / A+).
'
' Navigation : chapitres (liste) > thèmes (onglets) > sous-thèmes (boutons)
' > questions > détail (Réponse / Standards / Commentaire).
' Les réponses sont écrites dans le dossier de l'offre (.xlsx) à chaque
' changement de question, et le fichier est enregistré (décision ES-13).
' =============================================================================

' --- Contrôles avec événements ------------------------------------------------
Private WithEvents lstChap As MSForms.ListBox
Private WithEvents tsThemes As MSForms.TabStrip
Private WithEvents tsSST As MSForms.TabStrip
Private WithEvents cboFiltre As MSForms.ComboBox
Private WithEvents lstQ As MSForms.ListBox
Private WithEvents lstChoix As MSForms.ListBox
Private WithEvents cboChoix As MSForms.ComboBox
Private WithEvents txtRep As MSForms.TextBox
Private WithEvents optDiffNC As MSForms.OptionButton
Private WithEvents optDiffId As MSForms.OptionButton
Private WithEvents optDiffDiff As MSForms.OptionButton
Private WithEvents txtJustif As MSForms.TextBox
Private WithEvents txtCom As MSForms.TextBox
Private WithEvents txtRecherche As MSForms.TextBox
Private WithEvents lstRes As MSForms.ListBox
Private WithEvents btnPageGarde As MSForms.Label
Private WithEvents btnRech As MSForms.Label
Private WithEvents btnRetablir As MSForms.Label
Private WithEvents btnReprendre As MSForms.Label
Private WithEvents btnNC As MSForms.Label
Private WithEvents btnValSuiv As MSForms.Label
Private WithEvents btnPrec As MSForms.Label
Private WithEvents btnSuiv As MSForms.Label
Private WithEvents btnSuivTraiter As MSForms.Label
Private WithEvents btnAmoins As MSForms.Label
Private WithEvents btnAplus As MSForms.Label
Private WithEvents btnFermer As MSForms.Label
Private WithEvents btnFermerRes As MSForms.Label

Private mp As MSForms.MultiPage
Private fraRes As MSForms.Frame         ' résultats de recherche (créé une fois, masqué ou affiché)

' --- État -----------------------------------------------------------------------
Private curChap As cCategorieSGO
Private curTheme As cCategorieSGO
Private curSST As cCategorieSGO
Private curQ As cQuestionSGO
Private listeQ As Collection            ' questions affichées dans lstQ
Private themesAff As Collection         ' thèmes du chapitre (onglets)
Private sstAff As Collection            ' sous-thèmes du thème (boutons)
Private resultats As Collection         ' résultats de recherche
Private chargement As Boolean           ' ignore les événements pendant un remplissage par le code
Private sale As Boolean                 ' la question courante a changé depuis le dernier enregistrement

' --- Disposition ------------------------------------------------------------------
Private M As Single, W As Single, H As Single      ' marge, largeur et hauteur utiles
Private wDet As Single, hDet As Single  ' taille utile des pages du détail

Private Const FILTRES As String = "Toutes les questions;Sans réponse;Justification manquante;Standard modifié;Non conformes au standard"

' =============================================================================
' Ouverture, fermeture
' =============================================================================
Private Sub UserForm_Initialize()
    On Error GoTo erreur
    UIInit
    Construire
    If gQuestionDepart Is Nothing Then
        AllerAQuestion PremiereQuestionVisible()
    Else
        AllerAQuestion gQuestionDepart
        Set gQuestionDepart = Nothing
    End If
    Exit Sub
erreur:
    MsgBox "Erreur à l'ouverture de l'écran de saisie : " & Err.Description, vbExclamation, "SGO"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Valider
    SauverDossier
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub

' =============================================================================
' Construction de la fenêtre
' =============================================================================
Private Sub Construire()
    Dim hEntete As Single, hPied As Single, yCorps As Single, hCorps As Single, wNav As Single
    Dim x0 As Single, wZone As Single, xIn As Single, yIn As Single, wIn As Single, yBas As Single
    Dim wQ As Single, y2 As Single, xD As Single, wR As Single, y As Single, c As Object

    chargement = True
    M = UI_MARGE
    Me.Caption = "SGO - Élaboration du DCO   ·   " & gWbDossier.Name
    UIDimensionner Me, 0.95, 760, 440, 1600, 1000
    Me.BackColor = uiCoulFond
    W = Me.InsideWidth: H = Me.InsideHeight
    hEntete = 56: hPied = 44
    yCorps = hEntete + M
    hCorps = H - yCorps - hPied
    wNav = UIBorne(W * 0.2, 180, 260)
    x0 = M + wNav + M
    wZone = W - M - x0

    ' --- Bandeau ---------------------------------------------------------------
    UIFond Me, "lblBandeau", 0, 0, W, hEntete, uiBandeau
    UILbl Me, "lblTitre", "Élaboration du DCO", M + 4, 6, 260, UILigneH(4), True, 4, uiBlanc
    wR = UIBorne(W * 0.3, 240, 340)
    UILbl Me, "lblContexte", TexteContexte(), M + 4, 33, W - wR - 3 * M - 130, UILigneH(-1), False, -1, RGB(200, 216, 234)
    Me.Controls("lblContexte").WordWrap = False
    Set btnPageGarde = UIBouton(Me, "btnPageGarde", "Page de garde", W - M - wR - M - 124, 30, 124, 20, 2, ICO_CRAYON)
    Set txtRecherche = UITxt(Me, "txtRecherche", "", W - M - wR, 7, wR - 30, 20)
    txtRecherche.ControlTipText = "Identifiant (El84) ou mot-clé, puis Entrée"
    Set btnRech = UIBouton(Me, "btnRech", "", W - M - 28, 7, 28, 20, 1, ICO_RECHERCHE)
    UIFond Me, "lblBarreFond", W - M - wR, 37, wR - 130, 8, RGB(12, 36, 64)
    UIFond Me, "lblBarre", W - M - wR, 37, 1, 8, RGB(76, 175, 80)
    Set c = UILbl(Me, "lblAvancement", "", W - M - 124, 32, 124, UILigneH(-1), True, -1, uiBlanc)
    c.TextAlign = 3

    ' --- Chapitres ------------------------------------------------------------------
    UILbl Me, "lblChapTitre", "CHAPITRES", M + 4, yCorps, wNav - 8, UILigneH(-2), True, -2, uiGris
    Set lstChap = UIListe(Me, "lstChap", M, yCorps + 18, wNav, hCorps - 18 - 48, CLng(wNav - 46) & ";36")
    lstChap.Font.Name = UI_POLICE_SYMB
    lstChap.BackColor = uiOngletFond
    Set c = UILbl(Me, "lblLegende", ChrW(&H2713) & " terminé  " & ChrW(&H25D0) & " en cours  " & ChrW(&H25CB) & " à faire  " & _
                  "! à justifier  " & ChrW(&H21BB) & " standard modifié  " & ChrW(&H25C6) & " spécifique  " & ChrW(&H2298) & " sans objet", _
                  M + 2, yCorps + hCorps - 44, wNav - 4, 44, False, -2, uiGris)
    c.Font.Name = UI_POLICE_SYMB
    c.ControlTipText = "Sans objet : dépend de la réponse à une question mère, ou sans standard pour votre direction/univers"

    ' --- Thèmes (onglets) et sous-thèmes (boutons) -----------------------------------
    Set tsThemes = UICtl(Me, "Forms.TabStrip.1", "tsThemes", x0, yCorps, wZone, hCorps)
    tsThemes.Font.Name = UI_POLICE_SYMB
    tsThemes.TabFixedHeight = UILigneH() + 8
    Set c = UILbl(Me, "lblThemesTitre", "Thèmes", x0 + wZone - 96, yCorps + 5, 90, UILigneH(-2), False, -2, uiGris)
    c.TextAlign = 3
    UIFond Me, "lblZoneBlanche", x0 + 2, yCorps + UILigneH() + 12, wZone - 4, hCorps - UILigneH() - 14, uiBlanc
    xIn = x0 + M + 2: yIn = yCorps + UILigneH() + 12 + M: wIn = wZone - 2 * M - 4
    yBas = yCorps + hCorps - M

    Set tsSST = UICtl(Me, "Forms.TabStrip.1", "tsSST", xIn, yIn, wIn, UILigneH() + 12)
    tsSST.Font.Name = UI_POLICE_SYMB
    tsSST.Style = 1                                  ' boutons
    tsSST.TabFixedHeight = UILigneH() + 8
    Set c = UILbl(Me, "lblSSTTitre", "Sous-thèmes", xIn + wIn - 96, yIn + 6, 90, UILigneH(-2), False, -2, uiGris)
    c.TextAlign = 3
    Set c = UILbl(Me, "lblMasque", "", xIn + 2, yIn + UILigneH() + 16, wIn, UILigneH(-2), False, -2, uiGris)
    c.Font.Name = UI_POLICE_SYMB

    ' --- Questions ----------------------------------------------------------------------
    y2 = yIn + UILigneH() + 16 + UILigneH(-2) + 8
    wQ = UIBorne(wIn * 0.4, 230, 400)
    UISection Me, "secQuestions", "QUESTIONS", xIn, y2, wQ
    UILbl Me, "lblFiltre", "Afficher", xIn, y2 + 27, 52, UILigneH()
    Set cboFiltre = UICombo(Me, "cboFiltre", xIn + 54, y2 + 24, wQ - 54, UILigneH() + 8, Split(FILTRES, ";"))
    cboFiltre.ListIndex = 0
    UILbl Me, "lblCompteur", "", xIn, y2 + 24 + UILigneH() + 12, wQ, UILigneH(-2), False, -2, uiGris
    y = y2 + 24 + UILigneH() + 12 + UILigneH(-2) + 4
    Set lstQ = UIListe(Me, "lstQ", xIn, y, wQ, yBas - 36 - y, "16;48;" & CLng(wQ - 16 - 48 - 22))
    lstQ.Font.Name = UI_POLICE_SYMB
    Set btnRetablir = UIBouton(Me, "btnRetablir", "Rétablir les standards du sous-thème", xIn, yBas - 28, wQ, 28, 0, ICO_COCHE)

    ' --- Détail -----------------------------------------------------------------------
    xD = xIn + wQ + 2 * M
    Set mp = UICtl(Me, "Forms.MultiPage.1", "mpDetail", xD, y2, wIn - wQ - 2 * M, yBas - 38 - y2)
    mp.TabFixedHeight = UILigneH() + 8
    Do While mp.Pages.Count < 3
        mp.Pages.Add
    Loop
    mp.Pages(0).Caption = "Réponse"
    mp.Pages(1).Caption = "Standards"
    mp.Pages(2).Caption = "Commentaire"
    mp.Value = 0
    wDet = mp.Width - 6
    hDet = mp.Height - mp.TabFixedHeight - 10
    ' Boutons d'action de la question : hors des pages, car les pages sont reconstruites
    Set btnValSuiv = UIBouton(Me, "btnValSuiv", "Valider et suivante", xD + mp.Width - 196, yBas - 30, 196, 30, 1, ICO_SUIVANT)
    Set btnNC = UIBouton(Me, "btnNC", "Non concerné", xD + mp.Width - 196 - 8 - 128, yBas - 30, 128, 30, 0)
    btnNC.ControlTipText = "Répond « " & NC_OFFRE & " » (même règle que l'outil d'origine)"

    ' --- Pied de page --------------------------------------------------------------------
    y = H - hPied + 8
    Set btnPrec = UIBouton(Me, "btnPrec", "Précédente", M, y, 112, 28, 0, ICO_PRECEDENT)
    Set btnSuiv = UIBouton(Me, "btnSuiv", "Suivante", M + 120, y, 104, 28, 0, ICO_SUIVANT)
    Set btnSuivTraiter = UIBouton(Me, "btnSuivTraiter", "Suivante à traiter", M + 232, y, 170, 28, 0, ICO_SUIVANT)
    btnSuivTraiter.ControlTipText = "Question suivante sans réponse, à justifier ou au standard modifié"
    Set btnAmoins = UIBouton(Me, "btnAmoins", "A-", M + 410, y, 34, 28, 0)
    btnAmoins.ControlTipText = "Réduire la taille du texte"
    Set btnAplus = UIBouton(Me, "btnAplus", "A+", M + 448, y, 34, 28, 0)
    btnAplus.ControlTipText = "Agrandir la taille du texte"
    Set btnFermer = UIBouton(Me, "btnFermer", "Fermer", W - M - 126, y, 126, 28, 0)
    Set c = UILbl(Me, "lblEnregistre", "", M + 490, y + 7, W - M - 126 - 8 - (M + 490), UILigneH(-2), False, -2, uiVert)
    c.Font.Name = UI_POLICE_SYMB: c.TextAlign = 3

    ConstruireResultats
    RemplirChapitres
    AfficherAvancement
    chargement = False
End Sub

Private Function TexteContexte() As String
    Dim num As String, lib As String
    num = DossierValeur("numero_procedure"): lib = DossierValeur("libelle_procedure")
    TexteContexte = "Offre " & IIf(num = "", "(sans numéro)", num) & IIf(lib = "", "", " · " & lib) & "   |   " & _
                    DossierValeur("direction_achat") & " · Standard " & LibelleCode(gDirection) & _
                    IIf(gUniversChoisi = "", "", " · Univers " & LibelleCode(gUniversChoisi))
End Function

' Tout reconstruire (taille du texte ou page de garde modifiée) : la fenêtre est
' fermée puis rouverte sur la même question par SGO_Lancement.AfficherSaisie.
' (On ne supprime jamais un contrôle pendant son propre événement.)
Private Sub Rouvrir()
    Set gQuestionDepart = curQ
    gRouvrirSaisie = True
    Unload Me
End Sub

Private Sub LibererDetail()
    Set lstChoix = Nothing: Set cboChoix = Nothing: Set txtRep = Nothing
    Set optDiffNC = Nothing: Set optDiffId = Nothing: Set optDiffDiff = Nothing
    Set txtJustif = Nothing: Set btnReprendre = Nothing
End Sub

' =============================================================================
' Navigation : chapitres, thèmes, sous-thèmes, questions
' =============================================================================
Private Sub RemplirChapitres()
    Dim c As Variant, i As Long, sel As Long
    sel = lstChap.ListIndex
    lstChap.Clear
    For Each c In gChapitres
        UIAjouterLigne lstChap, UIPictoCategorie(c) & "  " & c.Libelle, IIf(c.Grisee(), "", c.Pourcentage() & " %")
    Next c
    If sel >= 0 And sel < lstChap.ListCount Then lstChap.ListIndex = sel
End Sub

' Met à jour les pictogrammes et pourcentages sans reconstruire la navigation
Private Sub RafraichirNavigation()
    Dim c As Variant, i As Long, etaitEnChargement As Boolean
    etaitEnChargement = chargement
    chargement = True
    i = 0
    For Each c In gChapitres
        lstChap.List(i, 0) = UIPictoCategorie(c) & "  " & c.Libelle
        lstChap.List(i, 1) = IIf(c.Grisee(), "", c.Pourcentage() & " %")
        i = i + 1
    Next c
    EtatOnglets tsThemes, themesAff
    EtatOnglets tsSST, sstAff
    RemplirQuestions
    AfficherAvancement
    chargement = etaitEnChargement
End Sub

Private Sub EtatOnglets(ts As MSForms.TabStrip, cats As Collection)
    Dim i As Long, c As cCategorieSGO
    If cats Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To cats.Count
        Set c = cats(i)
        ts.Tabs(i - 1).Caption = UIPictoCategorie(c) & " " & c.Libelle
        ts.Tabs(i - 1).Enabled = Not c.Grisee()
        ts.Tabs(i - 1).ControlTipText = IIf(c.Grisee(), RaisonMasquage(c.Id), c.Pourcentage() & " % traité")
    Next i
End Sub

Private Sub AfficherAvancement()
    Dim a As Variant, pct As Long, wMax As Single
    a = AvancementGlobal()
    If a(1) > 0 Then pct = Int(100 * a(0) / a(1))
    wMax = Me.Controls("lblBarreFond").Width
    Me.Controls("lblBarre").Width = UIBorne(wMax * pct / 100, 1, wMax)
    Me.Controls("lblAvancement").Caption = pct & " % · " & a(0) & " / " & a(1)
End Sub

Private Sub lstChap_Change()
    If chargement Or lstChap.ListIndex < 0 Then Exit Sub
    Valider
    AfficherChapitre gChapitres(lstChap.ListIndex + 1), Nothing
End Sub

Private Sub tsThemes_Change()
    If chargement Or tsThemes.Value < 0 Then Exit Sub
    Valider
    AfficherTheme themesAff(tsThemes.Value + 1), Nothing
End Sub

Private Sub tsSST_Change()
    If chargement Or tsSST.Value < 0 Then Exit Sub
    Valider
    AfficherSousTheme sstAff(tsSST.Value + 1), Nothing
End Sub

Private Sub cboFiltre_Change()
    If chargement Then Exit Sub
    Valider
    chargement = True
    RemplirQuestions
    chargement = False
    If listeQ.Count > 0 Then
        lstQ.ListIndex = 0
    Else
        AfficherQuestion Nothing
    End If
End Sub

' Affiche un chapitre ; si q est donnée, va jusqu'à elle
Private Sub AfficherChapitre(ByVal chap As cCategorieSGO, ByVal q As cQuestionSGO)
    Dim c As Variant, i As Long, cible As Long, th As cCategorieSGO
    Set curChap = chap
    chargement = True
    Set themesAff = New Collection
    tsThemes.Tabs.Clear
    cible = -1
    For Each c In chap.Enfants
        themesAff.Add c
        tsThemes.Tabs.Add "t" & themesAff.Count, UIPictoCategorie(c) & " " & c.Libelle
        If Not q Is Nothing Then
            If StrComp(gCategories(q.SousTheme).Parent, c.Id, vbTextCompare) = 0 Then cible = themesAff.Count - 1
        End If
        If cible < 0 And q Is Nothing And Not c.Grisee() Then cible = themesAff.Count - 1
    Next c
    EtatOnglets tsThemes, themesAff
    If cible < 0 Then cible = 0
    For i = 1 To gChapitres.Count
        If gChapitres(i) Is chap Then lstChap.ListIndex = i - 1
    Next i
    If themesAff.Count > 0 Then tsThemes.Value = cible
    chargement = False
    If themesAff.Count > 0 Then
        Set th = themesAff(cible + 1)
        AfficherTheme th, q
    End If
End Sub

Private Sub AfficherTheme(ByVal th As cCategorieSGO, ByVal q As cQuestionSGO)
    Dim c As Variant, cible As Long, plusieurs As Boolean
    Set curTheme = th
    chargement = True
    Set sstAff = New Collection
    tsSST.Tabs.Clear
    cible = -1
    For Each c In th.Enfants
        sstAff.Add c
        tsSST.Tabs.Add "s" & sstAff.Count, UIPictoCategorie(c) & " " & c.Libelle
        If Not q Is Nothing Then
            If StrComp(q.SousTheme, c.Id, vbTextCompare) = 0 Then cible = sstAff.Count - 1
        End If
        If cible < 0 And q Is Nothing And Not c.Grisee() Then cible = sstAff.Count - 1
    Next c
    EtatOnglets tsSST, sstAff
    If cible < 0 Then cible = 0
    plusieurs = (sstAff.Count > 1)
    tsSST.Visible = plusieurs
    Me.Controls("lblSSTTitre").Visible = plusieurs
    If sstAff.Count > 0 Then tsSST.Value = cible
    chargement = False
    If sstAff.Count > 0 Then AfficherSousTheme sstAff(cible + 1), q
End Sub

Private Sub AfficherSousTheme(ByVal sst As cCategorieSGO, ByVal q As cQuestionSGO)
    Dim i As Long, cible As Long
    Set curSST = sst
    If sst.Grisee() Then
        Me.Controls("lblMasque").Caption = ChrW(&H2298) & " « " & sst.Libelle & " » est sans objet : " & RaisonMasquage(sst.Id) & "."
    Else
        Me.Controls("lblMasque").Caption = ""
    End If
    chargement = True
    RemplirQuestions
    cible = -1
    If Not q Is Nothing Then
        For i = 1 To listeQ.Count
            If listeQ(i) Is q Then cible = i - 1
        Next i
        If cible < 0 Then                            ' filtrée : on revient à « Toutes »
            cboFiltre.ListIndex = 0
            RemplirQuestions
            For i = 1 To listeQ.Count
                If listeQ(i) Is q Then cible = i - 1
            Next i
        End If
    End If
    If cible < 0 Then
        For i = 1 To listeQ.Count
            If listeQ(i).Statut <> ST_MASQUEE Then cible = i - 1: Exit For
        Next i
    End If
    If cible < 0 And listeQ.Count > 0 Then cible = 0
    If cible >= 0 Then lstQ.ListIndex = cible
    chargement = False
    If cible >= 0 Then
        AfficherQuestion listeQ(cible + 1)
    Else
        AfficherQuestion Nothing
    End If
End Sub

' Liste des questions du sous-thème, selon le filtre
Private Sub RemplirQuestions()
    Dim q As Variant, lib As String, nMasq As Long, nTot As Long, sel As Long, i As Long
    If curSST Is Nothing Then Exit Sub
    Set listeQ = New Collection
    lstQ.Clear
    For Each q In curSST.Questions
        nTot = nTot + 1
        If q.Statut = ST_MASQUEE Then nMasq = nMasq + 1
        If PasseFiltre(q) Then
            listeQ.Add q
            lib = Replace(Replace(q.Libelle, vbCr, " "), vbLf, " ")
            If q.Statut = ST_MASQUEE Then lib = "(" & RaisonMasquage(q.Id) & ") " & lib
            UIAjouterLigne lstQ, UIPicto(q.Statut, q.StandardModifie), q.Id, lib
        End If
    Next q
    If cboFiltre.ListIndex <= 0 Then
        Me.Controls("lblCompteur").Caption = nTot & " question" & IIf(nTot > 1, "s", "") & _
            IIf(nMasq > 0, ", dont " & nMasq & " sans objet (grisée" & IIf(nMasq > 1, "s", "") & ", non modifiable" & IIf(nMasq > 1, "s", "") & ")", "")
    Else
        Me.Controls("lblCompteur").Caption = listeQ.Count & " sur " & nTot & " question" & IIf(nTot > 1, "s", "")
    End If
    sel = -1
    If Not curQ Is Nothing Then
        For i = 1 To listeQ.Count
            If listeQ(i) Is curQ Then sel = i - 1
        Next i
    End If
    If sel >= 0 Then lstQ.ListIndex = sel
End Sub

Private Function PasseFiltre(ByVal q As cQuestionSGO) As Boolean
    Select Case cboFiltre.ListIndex
        Case 1: PasseFiltre = (q.Statut = ST_NON_REPONDUE)
        Case 2: PasseFiltre = (q.Statut = ST_A_JUSTIFIER)
        Case 3: PasseFiltre = q.StandardModifie
        Case 4: PasseFiltre = (q.Conformite = CONF_NON_CONFORME)
        Case Else: PasseFiltre = True
    End Select
End Function

Private Sub lstQ_Change()
    If chargement Or lstQ.ListIndex < 0 Then Exit Sub
    If listeQ(lstQ.ListIndex + 1) Is curQ Then Exit Sub
    Valider
    AfficherQuestion listeQ(lstQ.ListIndex + 1)
End Sub

' Va jusqu'à une question : chapitre, thème, sous-thème, sélection
Private Sub AllerAQuestion(ByVal q As cQuestionSGO)
    Dim th As cCategorieSGO, chap As cCategorieSGO
    If q Is Nothing Then Exit Sub
    Set th = gCategories(gCategories(q.SousTheme).Parent)
    Set chap = gCategories(th.Parent)
    AfficherChapitre chap, q
End Sub

Private Function PremiereQuestionVisible() As cQuestionSGO
    Dim q As Variant
    For Each q In gOrdre
        If q.Statut <> ST_MASQUEE Then
            Set PremiereQuestionVisible = q
            Exit Function
        End If
    Next q
    If gOrdre.Count > 0 Then Set PremiereQuestionVisible = gOrdre(1)
End Function

' Question visible suivante (ou précédente) dans l'ordre du référentiel
Private Function Voisine(ByVal sens As Long, ByVal aTraiter As Boolean) As cQuestionSGO
    Dim i As Long, n As Long, depart As Long, k As Long, q As cQuestionSGO
    n = gOrdre.Count
    depart = 1
    If Not curQ Is Nothing Then
        For i = 1 To n
            If gOrdre(i) Is curQ Then depart = i: Exit For
        Next i
    End If
    If Not aTraiter Then
        ' Précédente / Suivante : on s'arrête à la première ou à la dernière question
        i = depart + sens
        Do While i >= 1 And i <= n
            If gOrdre(i).Statut <> ST_MASQUEE Then
                Set Voisine = gOrdre(i)
                Exit Function
            End If
            i = i + sens
        Loop
        Exit Function
    End If
    ' Suivante à traiter : on fait le tour complet
    For k = 1 To n
        i = ((depart - 1 + sens * k) Mod n + n) Mod n + 1
        Set q = gOrdre(i)
        If q.Statut <> ST_MASQUEE Then
            If q.Statut = ST_NON_REPONDUE Or q.Statut = ST_A_JUSTIFIER Or q.StandardModifie Then
                Set Voisine = q: Exit Function
            End If
        End If
    Next k
End Function

Private Sub btnPrec_Click()
    Aller Voisine(-1, False), "Première question atteinte."
End Sub

Private Sub btnSuiv_Click()
    Aller Voisine(1, False), "Dernière question atteinte."
End Sub

Private Sub btnSuivTraiter_Click()
    Aller Voisine(1, True), "Toutes les questions visibles sont traitées."
End Sub

Private Sub Aller(ByVal q As cQuestionSGO, ByVal messageSiRien As String)
    Valider
    If q Is Nothing Then
        MsgBox messageSiRien, vbInformation, "SGO"
    Else
        If cboFiltre.ListIndex > 0 And Not PasseFiltre(q) Then
            chargement = True: cboFiltre.ListIndex = 0: chargement = False
        End If
        AllerAQuestion q
    End If
End Sub

' =============================================================================
' Détail d'une question
' =============================================================================
Private Sub AfficherQuestion(ByVal q As cQuestionSGO)
    Set curQ = q
    sale = False
    chargement = True
    ConstruirePageReponse
    ConstruirePageStandards
    ConstruirePageCommentaire
    chargement = False
    UIBoutonActif Me, "btnRetablir", Not curSST Is Nothing
    MettreAJourBoutons
End Sub

Private Sub MettreAJourBoutons()
    Dim modifiable As Boolean
    If Not curQ Is Nothing Then modifiable = (curQ.Statut <> ST_MASQUEE And Not curQ.EstAnnexe)
    UIBoutonActif Me, "btnNC", modifiable
    UIBoutonActif Me, "btnValSuiv", Not curQ Is Nothing
    If Not btnReprendre Is Nothing Then UIBoutonActif mp.Pages(0), "btnReprendre", curQ.Saisie
End Sub

Private Sub ConstruirePageReponse()
    Dim pg As MSForms.Page, iw As Single, y As Single, c As Object, hz As Single, txt As String
    Dim masquee As Boolean, i As Long, sp As Boolean, avecStd As Boolean, lesFilles As String

    Set pg = mp.Pages(0)
    LibererDetail
    pg.Controls.Clear
    pg.ScrollBars = 0: pg.ScrollTop = 0
    iw = UIBorne(wDet - 2 * M - 14, 200, 660)
    If curQ Is Nothing Then
        UILbl pg, "lblVide", IIf(cboFiltre.ListIndex > 0, "Aucune question ne correspond au filtre.", "Aucune question."), _
              M, 12, iw, UILigneH(), False, 0, uiGris
        Exit Sub
    End If
    masquee = (curQ.Statut = ST_MASQUEE)
    sp = (curQ.Niveau = "Sp")
    avecStd = (curQ.Niveau = "E" Or curQ.Niveau = "D" Or curQ.Niveau = "U")

    ' Identifiant, mode, condition, questions conditionnées
    txt = curQ.Id & "  ·  " & curQ.Mode
    If RaisonAffichage(curQ.Id) <> "" And Not masquee Then txt = txt & "  ·  " & RaisonAffichage(curQ.Id)
    If gFilles.Exists(curQ.Id) Then
        lesFilles = gFilles(curQ.Id)
        txt = txt & "  ·  Conditionne : " & lesFilles
    End If
    hz = LignesTexte(txt, iw, -2) * UILigneH(-2)
    Set c = UILbl(pg, "lblTypeQ", txt, M, 8, iw, hz, False, -2, uiGris)
    If lesFilles <> "" Then c.ControlTipText = "Votre réponse fait apparaître ou masque : " & lesFilles
    y = 8 + hz + 4
    hz = UIBorne(LignesTexte(curQ.Libelle, iw, 2), 1, 6) * UILigneH(2)
    UILbl pg, "lblLibelle", curQ.Libelle, M, y, iw, hz, True, 2
    y = y + hz + 10

    ' Encadré : standard, spécifique, sans objet, annexe ou standard incomplet
    If masquee Then
        y = Encadre(pg, y, iw, ChrW(&H2298) & "  SANS OBJET POUR CETTE OFFRE", _
                    RaisonMasquage(curQ.Id) & ". La question reste visible mais n'est pas modifiable." & _
                    IIf(curQ.AQuelqueChose, " Les réponses déjà saisies sont conservées et seront reprises si la question redevient applicable.", ""), _
                    uiCoulFond, uiGris, uiGris)
    ElseIf curQ.EstAnnexe Then
        y = Encadre(pg, y, iw, "RÉPONSE DANS UNE ANNEXE", _
                    "Cette question se renseigne dans l'annexe « " & curQ.Reponse & " ». Les annexes seront intégrées dans une prochaine itération.", _
                    uiCoulFond, uiGris, uiGris)
    ElseIf avecStd Then
        y = Encadre(pg, y, iw, "STANDARD APPLICABLE  ·  " & LibelleNiveau(curQ.Niveau), curQ.Standard, _
                    uiStdFond, uiStdBarre, uiOrange, True)
        If curQ.StandardModifie Then
            Set c = UILbl(pg, "lblStdModifie", ChrW(&H21BB) & "  Le standard a changé depuis votre saisie (page de garde modifiée) : " & _
                          "vérifiez la réponse, puis cliquez sur « Valider et suivante ».", _
                          M, y - 6, iw, 2 * UILigneH(-2), True, -2, uiOrange)
            c.Font.Name = UI_POLICE_SYMB
            y = y + 2 * UILigneH(-2)
        End If
    ElseIf sp Then
        y = Encadre(pg, y, iw, "AUCUN STANDARD  ·  Spécifique à l'offre (Sp)", _
                    "Réponse propre à cette offre. Indiquez si elle diffère de l'offre précédente.", uiCoulFond, uiGris, uiGris)
    Else
        y = Encadre(pg, y, iw, "STANDARDS INCOMPLETS", _
                    "Aucun standard n'est défini pour cette question avec votre direction et votre univers : contactez votre administrateur.", _
                    uiStdFond, uiStdBarre, uiOrange)
    End If

    ' Réponse
    UISection pg, "secReponse", "RÉPONSE", M, y, iw
    y = y + 26
    If curQ.EstAnnexe Then
        UILbl pg, "lblRepAnnexe", curQ.Reponse, M, y, iw, UILigneH(), True
        y = y + UILigneH() + 12
    ElseIf curQ.EstListe And curQ.NbChoix > 0 And InStr(1, curQ.Mode, "texte", vbTextCompare) = 0 Then
        hz = UIBorne(curQ.NbChoix, 2, 8) * UILigneH() + 8
        Set lstChoix = UIListe(pg, "lstChoix", M, y, iw, hz, "")
        lstChoix.ListStyle = 1                       ' cases à cocher / boutons d'option
        lstChoix.MultiSelect = IIf(curQ.EstChoixMultiple, 1, 0)
        For i = 0 To UBound(curQ.Choix)
            lstChoix.AddItem curQ.Choix(i)
        Next i
        SelectionnerChoix
        lstChoix.Enabled = Not masquee
        y = y + hz + 4
        hz = UIBorne(LignesTexte("Réponse retenue : " & curQ.Reponse, iw, -1), 1, 4) * UILigneH(-1)
        UILbl pg, "lblRepComplete", "Réponse retenue : " & IIf(curQ.Reponse = "", "(aucune)", curQ.Reponse), _
              M, y, iw, hz, False, -1, uiGris
        y = y + hz + 10
    ElseIf curQ.EstListe And curQ.NbChoix > 0 Then
        Set cboChoix = UICombo(pg, "cboChoix", M, y, iw, UILigneH() + 8, curQ.Choix, True)
        cboChoix.Text = curQ.Reponse
        cboChoix.Enabled = Not masquee
        cboChoix.ControlTipText = "Choisissez dans la liste ou saisissez un texte"
        y = y + UILigneH() + 20
    Else
        hz = UIBorne(LignesTexte(curQ.Reponse, iw, 0) + 1, 3, 8) * UILigneH() + 8
        Set txtRep = UITxt(pg, "txtRep", curQ.Reponse, M, y, iw, hz)
        txtRep.MultiLine = True: txtRep.WordWrap = True: txtRep.ScrollBars = 2
        txtRep.EnterKeyBehavior = True
        txtRep.Enabled = Not masquee
        y = y + hz + 14
    End If

    If Not masquee Then
        ' Conformité calculée
        y = BandeauConformite(pg, y, iw)
        ' Différence avec l'offre précédente
        If sp Then
            UISection pg, "secDiff", "DIFFÉRENCE AVEC L'OFFRE PRÉCÉDENTE", M, y, iw
            y = y + 26
            Set optDiffId = UIOpt(pg, "optDiffId", DIFF_IDENTIQUE, M, y, 120, UILigneH() + 6, "grpDiff")
            Set optDiffDiff = UIOpt(pg, "optDiffDiff", DIFF_DIFFERENTE, M + 130, y, 120, UILigneH() + 6, "grpDiff")
            UILbl pg, "lblDiffInfo", "Par défaut « Identique » : à vérifier.", M + 260, y + 3, iw - 260, UILigneH(-2), False, -2, uiGris
            y = y + UILigneH() + 18
        ElseIf avecStd Then
            UILbl pg, "lblDiffFac", "Différence avec l'offre précédente (facultatif)", M, y, iw, UILigneH(-2), False, -2, uiGris
            Set optDiffNC = UIOpt(pg, "optDiffNC", DIFF_NC, M, y + UILigneH(-2) + 2, 120, UILigneH() + 4, "grpDiff")
            Set optDiffId = UIOpt(pg, "optDiffId", DIFF_IDENTIQUE, M + 126, y + UILigneH(-2) + 2, 100, UILigneH() + 4, "grpDiff")
            Set optDiffDiff = UIOpt(pg, "optDiffDiff", DIFF_DIFFERENTE, M + 232, y + UILigneH(-2) + 2, 110, UILigneH() + 4, "grpDiff")
            y = y + UILigneH(-2) + UILigneH() + 16
        End If
        Select Case DifferenceEnVigueur(curQ)
            Case DIFF_IDENTIQUE: If Not optDiffId Is Nothing Then optDiffId.Value = True
            Case DIFF_DIFFERENTE: If Not optDiffDiff Is Nothing Then optDiffDiff.Value = True
            Case Else: If Not optDiffNC Is Nothing Then optDiffNC.Value = True
        End Select

        ' Justification (obligatoire pour les cas 2 et 4)
        UISection pg, "secJustif", IIf(sp, "JUSTIFICATION DU CAS DIFFÉRENCIANT", "JUSTIFICATION DE LA DÉROGATION AU STANDARD"), M, y, iw
        y = y + 26
        hz = UIBorne(LignesTexte(curQ.Justification, iw, 0) + 1, 3, 8) * UILigneH() + 8
        Set txtJustif = UITxt(pg, "txtJustif", curQ.Justification, M, y, iw, hz)
        txtJustif.MultiLine = True: txtJustif.WordWrap = True: txtJustif.ScrollBars = 2
        txtJustif.EnterKeyBehavior = True
        If curQ.Statut = ST_A_JUSTIFIER Then txtJustif.BorderColor = uiStdBarre
        y = y + hz + 12
    End If

    If y > hDet Then
        pg.ScrollBars = 2
        pg.ScrollHeight = y + M
    End If
End Sub

' Encadré coloré (titre + texte) ; renvoie la position sous l'encadré
Private Function Encadre(ByVal pg As Object, ByVal y As Single, ByVal iw As Single, titre As String, texte As String, _
                         ByVal fond As Long, ByVal barre As Long, ByVal coulTitre As Long, _
                         Optional ByVal boutonReprendre As Boolean = False) As Single
    Dim hz As Single, wTexte As Single, c As Object
    wTexte = iw - 28 - IIf(boutonReprendre, 160, 0)
    hz = UIBorne(LignesTexte(texte, wTexte, 0), 1, 8) * UILigneH()
    UIFond pg, "lblEncFond", M, y, iw, hz + UILigneH(-2) + 22, fond
    UIFond pg, "lblEncBarre", M, y, 4, hz + UILigneH(-2) + 22, barre
    Set c = UILbl(pg, "lblEncTitre", titre, M + 14, y + 7, iw - 28, UILigneH(-2), True, -2, coulTitre)
    c.Font.Name = UI_POLICE_SYMB
    Set c = UILbl(pg, "lblEncTexte", texte, M + 14, y + 10 + UILigneH(-2), wTexte, hz, boutonReprendre Or barre = uiStdBarre)
    If boutonReprendre Then
        Set btnReprendre = UIBouton(pg, "btnReprendre", "Reprendre le standard", M + iw - 152, y + 10 + UILigneH(-2), 144, UILigneH() + 10, 0)
        btnReprendre.ControlTipText = "Remplace votre réponse par la valeur du standard"
    End If
    Encadre = y + hz + UILigneH(-2) + 22 + 16
End Function

Private Function BandeauConformite(ByVal pg As Object, ByVal y As Single, ByVal iw As Single) As Single
    Dim txt As String, fond As Long, coul As Long, info As String, c As Object
    TexteBandeau txt, fond, coul
    UIFond pg, "lblConfFond", M, y, iw, UILigneH() + 10, fond
    Set c = UILbl(pg, "lblConformite", txt, M + 8, y + 5, iw - 16, UILigneH(-1), True, -1, coul)
    c.Font.Name = UI_POLICE_SYMB
    If curQ.Niveau = "Sp" Then
        info = "Statut calculé automatiquement : aucun standard pour cette question."
    Else
        info = "Conformité calculée automatiquement : la réponse doit " & _
               IIf(curQ.Operateur = "=", "être égale au", "contenir le") & " standard."
    End If
    UILbl pg, "lblConfInfo", info, M, y + UILigneH() + 13, iw, UILigneH(-2), False, -2, uiGris
    BandeauConformite = y + UILigneH() + 13 + UILigneH(-2) + 14
End Function

Private Sub TexteBandeau(txt As String, fond As Long, coul As Long)
    Select Case curQ.Statut
        Case ST_CONFORME
            txt = ChrW(&H2713) & "  Conforme au standard   ·   cas " & ChrW(&H2460) & " : standard en place"
            fond = RGB(232, 245, 236): coul = uiVert
        Case ST_A_JUSTIFIER, ST_JUSTIFIEE
            If curQ.Conformite = CONF_SPECIFIQUE Then
                txt = ChrW(&H25C6) & "  Spécifique à l'offre   ·   cas " & ChrW(&H2463) & " : différent de l'offre précédente"
                fond = uiBleuClair: coul = uiAction
            Else
                txt = ChrW(&H2717) & "  Non conforme au standard   ·   cas " & ChrW(&H2461) & " : dérogation"
                fond = uiRougeClair: coul = uiRouge
            End If
            txt = txt & IIf(curQ.Statut = ST_A_JUSTIFIER, ", justification obligatoire", ", justifiée")
        Case ST_SPECIFIQUE
            txt = ChrW(&H25C6) & "  Spécifique à l'offre   ·   cas " & ChrW(&H2462) & " : identique à l'offre précédente"
            fond = uiBleuClair: coul = uiAction
        Case ST_NON_CONCERNEE
            txt = ChrW(&H2013) & "  Non concerné"
            fond = uiCoulFond: coul = uiGris
        Case ST_INCOMPLET
            txt = "?  Standards incomplets pour cette question"
            fond = uiStdFond: coul = uiOrange
        Case Else
            txt = ChrW(&H25CB) & "  Réponse à saisir"
            fond = uiCoulFond: coul = uiGris
    End Select
End Sub

Private Sub ConstruirePageStandards()
    Dim pg As MSForms.Page, iw As Single, lst As Object, c As Object, hz As Single
    Set pg = mp.Pages(1)
    pg.Controls.Clear
    pg.ScrollBars = 0
    iw = wDet - 2 * M
    If curQ Is Nothing Then Exit Sub
    UILbl pg, "lblStdInfo", "Le niveau applicable dépend du standard de direction et de l'univers choisis en page de garde.", _
          M, 8, iw, 2 * UILigneH(-1), False, -1, uiGris
    Set lst = UIListe(pg, "lstStd", M, 12 + 2 * UILigneH(-1), iw, 6 * UILigneH() + 8, "150;" & CLng(iw - 156))
    UIAjouterLigne lst, "Établissement (E)", Ligne1(ValeurTable(curQ.Id, "E", "", 1))
    UIAjouterLigne lst, "Direction " & LibelleCode(gDirection), _
                   Ligne1(ValeurTable(curQ.Id, "D", gDirection, 1)) & "   (niveau : " & IIf(ValeurTable(curQ.Id, "D", gDirection, 0) = "", "-", ValeurTable(curQ.Id, "D", gDirection, 0)) & ")"
    UIAjouterLigne lst, "Univers " & LibelleCode(gUniversChoisi), Ligne1(ValeurTable(curQ.Id, "U", gUniversChoisi, 1))
    UIAjouterLigne lst, "Niveau applicable", LibelleNiveau(curQ.Niveau)
    UIAjouterLigne lst, "Opérateur", IIf(curQ.Operateur = "=", "égal (la réponse doit être égale au standard)", "contient (la réponse doit contenir le standard)")
    UIAjouterLigne lst, "Alimente le CCAP", IIf(curQ.AlimenteCCAP, "Oui", "Non")
    hz = 12 + 2 * UILigneH(-1) + 6 * UILigneH() + 8 + 14
    UISection pg, "secStdComplet", "VALEUR COMPLÈTE DU STANDARD APPLICABLE", M, hz, iw
    Set c = UITxt(pg, "txtStdComplet", IIf(curQ.Standard = "", "(aucun)", curQ.Standard), M, hz + 26, iw, _
                  UIBorne(hDet - hz - 26 - M, 40, 400))
    c.MultiLine = True: c.WordWrap = True: c.ScrollBars = 2: c.Locked = True: c.BackColor = uiCoulFond
End Sub

Private Sub ConstruirePageCommentaire()
    Dim pg As MSForms.Page, iw As Single, c As Object
    Set pg = mp.Pages(2)
    Set txtCom = Nothing
    pg.Controls.Clear
    iw = wDet - 2 * M
    If curQ Is Nothing Then Exit Sub
    UISection pg, "secCom", "COMMENTAIRE SUR CETTE QUESTION", M, 8, iw
    Set txtCom = UITxt(pg, "txtCom", curQ.Commentaire, M, 34, iw, UIBorne(hDet - 34 - 30, 60, 600))
    txtCom.MultiLine = True: txtCom.WordWrap = True: txtCom.ScrollBars = 2
    txtCom.EnterKeyBehavior = True
    txtCom.MaxLength = 5000                         ' décision ES-14
    Set c = UILbl(pg, "lblCompteurCom", "", M + iw - 200, hDet - 22, 200, UILigneH(-2), False, -2, uiGris)
    c.TextAlign = 3
    AfficherCompteurCom
    mp.Pages(2).Caption = "Commentaire" & IIf(curQ.Commentaire <> "", " (1)", "")
End Sub

Private Sub AfficherCompteurCom()
    If txtCom Is Nothing Then Exit Sub
    mp.Pages(2).Controls("lblCompteurCom").Caption = Format$(Len(txtCom.Text), "#,##0") & " / 5 000 caractères"
End Sub

Private Function LibelleNiveau(ByVal niv As String) As String
    Select Case niv
        Case "E": LibelleNiveau = "Établissement (E)"
        Case "D": LibelleNiveau = "Direction (D)"
        Case "U": LibelleNiveau = "Univers (U)"
        Case "Sp": LibelleNiveau = "Spécifique à l'offre (Sp)"
        Case "": LibelleNiveau = "(non défini)"
        Case Else: LibelleNiveau = niv
    End Select
End Function

Private Function Ligne1(ByVal s As String) As String
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    Ligne1 = IIf(s = "", "(pas de standard)", s)
End Function

' Nombre de lignes approximatif d'un texte affiché sur une largeur donnée
Private Function LignesTexte(ByVal s As String, ByVal largeur As Single, ByVal delta As Single) As Long
    Dim p As Variant, n As Long, parLigne As Long
    parLigne = Int(largeur / ((uiTaille + delta) * 0.5))
    If parLigne < 10 Then parLigne = 10
    For Each p In Split(Replace(s, vbCr, ""), vbLf)
        n = n + Int((Len(p) + parLigne - 1) / parLigne)
        If Len(p) = 0 Then n = n + 1
    Next p
    If n < 1 Then n = 1
    LignesTexte = n
End Function

' Coche dans la liste les choix contenus dans la réponse
Private Sub SelectionnerChoix()
    Dim i As Long, morceaux As Variant, morceau As Variant
    If lstChoix Is Nothing Then Exit Sub
    If curQ.EstChoixMultiple Then
        morceaux = Split(curQ.Reponse, ";")
        For i = 0 To lstChoix.ListCount - 1
            lstChoix.Selected(i) = False
            For Each morceau In morceaux
                If StrComp(Trim$(morceau), Trim$(lstChoix.List(i)), vbTextCompare) = 0 Then lstChoix.Selected(i) = True
            Next morceau
        Next i
    Else
        lstChoix.ListIndex = -1
        For i = 0 To lstChoix.ListCount - 1
            If StrComp(Trim$(curQ.Reponse), Trim$(lstChoix.List(i)), vbTextCompare) = 0 Then lstChoix.ListIndex = i
        Next i
    End If
End Sub

' =============================================================================
' Saisie
' =============================================================================
Private Sub lstChoix_Change()
    Dim i As Long, s As String
    If chargement Then Exit Sub
    If curQ.EstChoixMultiple Then
        For i = 0 To lstChoix.ListCount - 1
            If lstChoix.Selected(i) Then s = s & IIf(s = "", "", ";") & lstChoix.List(i)
        Next i
    ElseIf lstChoix.ListIndex >= 0 Then
        s = lstChoix.List(lstChoix.ListIndex)
    End If
    ChangerReponse s, True
End Sub

Private Sub cboChoix_Change()
    If chargement Then Exit Sub
    ChangerReponse cboChoix.Text, True
End Sub

Private Sub txtRep_Change()
    If chargement Then Exit Sub
    ChangerReponse txtRep.Text, True
End Sub

Private Sub btnReprendre_Click()
    If curQ Is Nothing Then Exit Sub
    If Not curQ.Saisie Then Exit Sub
    ChangerReponse "", False
    AfficherReponseSurPlace
End Sub

Private Sub btnNC_Click()
    If curQ Is Nothing Then Exit Sub
    If curQ.Statut = ST_MASQUEE Or curQ.EstAnnexe Then Exit Sub
    ChangerReponse NC_OFFRE, True
    AfficherReponseSurPlace
End Sub

' Remet la zone de réponse à la valeur en vigueur, sans reconstruire la page
Private Sub AfficherReponseSurPlace()
    chargement = True
    If Not txtRep Is Nothing Then txtRep.Text = curQ.Reponse
    If Not cboChoix Is Nothing Then cboChoix.Text = curQ.Reponse
    If Not lstChoix Is Nothing Then SelectionnerChoix
    chargement = False
    MettreAJourDetail
End Sub

' Applique une nouvelle réponse, réévalue le dossier, demande confirmation si
' des questions déjà renseignées vont être masquées (cadrage P5)
Private Sub ChangerReponse(ByVal texte As String, ByVal estSaisie As Boolean)
    Dim ancSaisie As Boolean, ancTexte As String, avant As Object, k As Variant, q As cQuestionSGO
    Dim masquees As String, n As Long

    If curQ Is Nothing Then Exit Sub
    ancSaisie = curQ.Saisie: ancTexte = curQ.ReponseSaisie
    Set avant = NouveauDico()
    For Each k In gQuestions.Keys
        Set q = gQuestions(k)
        If q.Statut <> ST_MASQUEE And q.AQuelqueChose And Not q Is curQ Then avant(k) = True
    Next k

    curQ.Saisie = estSaisie
    curQ.ReponseSaisie = IIf(estSaisie, texte, "")
    curQ.StandardModifie = False
    EvaluerDossier

    For Each k In avant.Keys
        If gQuestions(k).Statut = ST_MASQUEE Then
            n = n + 1
            If n <= 10 Then masquees = masquees & IIf(masquees = "", "", ", ") & gQuestions(k).Id
        End If
    Next k
    If n > 0 Then
        If MsgBox("Cette réponse masque " & n & " question(s) déjà renseignée(s) (" & masquees & IIf(n > 10, "…", "") & ")." & _
                  vbCrLf & "Leurs réponses seront conservées mais ignorées ; elles réapparaîtront si la condition redevient vraie." & _
                  vbCrLf & vbCrLf & "Continuer ?", vbYesNo + vbQuestion, "SGO") = vbNo Then
            curQ.Saisie = ancSaisie: curQ.ReponseSaisie = ancTexte
            EvaluerDossier
            AfficherReponseSurPlace
            Exit Sub
        End If
    End If
    sale = True
    RafraichirApresSaisie
End Sub

' Rafraîchit ce qui dépend de la réponse, sans reconstruire la page (la saisie
' est en cours dans un de ses contrôles)
Private Sub RafraichirApresSaisie()
    RafraichirNavigation
    MettreAJourDetail
End Sub

Private Sub MettreAJourDetail()
    Dim pg As Object, txt As String, fond As Long, coul As Long
    Set pg = mp.Pages(0)
    On Error Resume Next
    pg.Controls("lblRepComplete").Caption = "Réponse retenue : " & IIf(curQ.Reponse = "", "(aucune)", curQ.Reponse)
    If Not curQ.StandardModifie Then pg.Controls("lblStdModifie").Caption = ""
    TexteBandeau txt, fond, coul
    pg.Controls("lblConfFond").BackColor = fond
    pg.Controls("lblConformite").Caption = txt
    pg.Controls("lblConformite").ForeColor = coul
    If Not txtJustif Is Nothing Then
        txtJustif.BorderColor = IIf(curQ.Statut = ST_A_JUSTIFIER, uiStdBarre, RGB(170, 178, 189))
    End If
    On Error GoTo 0
    MettreAJourBoutons
End Sub

' Reconstruit le détail de la question courante (appelé uniquement depuis un
' contrôle situé hors de la page Réponse)
Private Sub AfficherQuestionCourante()
    Dim estSale As Boolean
    If curQ Is Nothing Then Exit Sub
    estSale = sale
    AfficherQuestion curQ
    sale = estSale
End Sub

Private Sub optDiffNC_Click()
    ChangerDifference DIFF_NC
End Sub

Private Sub optDiffId_Click()
    ChangerDifference DIFF_IDENTIQUE
End Sub

Private Sub optDiffDiff_Click()
    ChangerDifference DIFF_DIFFERENTE
End Sub

Private Sub ChangerDifference(ByVal valeur As String)
    If chargement Or curQ Is Nothing Then Exit Sub
    If DifferenceEnVigueur(curQ) = valeur Then Exit Sub
    curQ.Difference = valeur
    sale = True
    EvaluerDossier
    RafraichirApresSaisie
End Sub

Private Sub txtJustif_Change()
    If chargement Or curQ Is Nothing Then Exit Sub
    curQ.Justification = txtJustif.Text
    curQ.StandardModifie = False
    sale = True
    EvaluerDossier
    RafraichirApresSaisie
End Sub

Private Sub txtCom_Change()
    If chargement Or curQ Is Nothing Then Exit Sub
    curQ.Commentaire = txtCom.Text
    sale = True
    AfficherCompteurCom
End Sub

Private Sub btnValSuiv_Click()
    If curQ Is Nothing Then Exit Sub
    If curQ.StandardModifie Then
        curQ.StandardModifie = False
        sale = True
        EvaluerDossier
    End If
    Aller Voisine(1, False), "Dernière question atteinte."
End Sub

' Écrit la question courante dans le dossier et enregistre le fichier
Private Sub Valider()
    On Error GoTo erreur
    If curQ Is Nothing Or Not sale Then Exit Sub
    EnregistrerQuestion curQ
    SauverDossier
    sale = False
    Me.Controls("lblEnregistre").Caption = ChrW(&H2714) & " Enregistré à " & Format$(Now, "hh:nn")
    Exit Sub
erreur:
    MsgBox "Enregistrement impossible : " & Err.Description & vbCrLf & _
           "Vérifiez que le dossier n'est pas ouvert en lecture seule.", vbExclamation, "SGO"
End Sub

' Questions du sous-thème remises au standard
Private Sub btnRetablir_Click()
    Dim q As Variant, n As Long
    If curSST Is Nothing Then Exit Sub
    Valider
    For Each q In curSST.Questions
        If q.Statut <> ST_MASQUEE And q.Saisie And (q.Niveau = "E" Or q.Niveau = "D" Or q.Niveau = "U") Then n = n + 1
    Next q
    If n = 0 Then
        MsgBox "Toutes les questions avec standard de ce sous-thème sont déjà au standard.", vbInformation, "SGO"
        Exit Sub
    End If
    If MsgBox(n & " réponse(s) de ce sous-thème vont être remplacées par la valeur du standard." & vbCrLf & _
              "Les justifications et commentaires sont conservés." & vbCrLf & vbCrLf & "Continuer ?", _
              vbYesNo + vbQuestion, "SGO") = vbNo Then Exit Sub
    For Each q In curSST.Questions
        If q.Statut <> ST_MASQUEE And q.Saisie And (q.Niveau = "E" Or q.Niveau = "D" Or q.Niveau = "U") Then
            q.Saisie = False: q.ReponseSaisie = "": q.StandardModifie = False
            EnregistrerQuestion q
        End If
    Next q
    EvaluerDossier
    SauverDossier
    RafraichirNavigation
    AfficherQuestionCourante
End Sub

' =============================================================================
' Page de garde, taille du texte
' =============================================================================
Private Sub btnPageGarde_Click()
    Valider
    If ModifierPageGarde() Then Rouvrir
End Sub

Private Sub btnAmoins_Click()
    Valider
    UIChangerTaille -1
    Rouvrir
End Sub

Private Sub btnAplus_Click()
    Valider
    UIChangerTaille 1
    Rouvrir
End Sub

' =============================================================================
' Recherche (identifiant ou mot-clé)
' =============================================================================
Private Sub txtRecherche_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0: Rechercher
    If KeyCode = 27 Then KeyCode = 0: FermerResultats
End Sub

Private Sub btnRech_Click()
    Rechercher
End Sub

' Cadre des résultats : créé une fois, masqué ; jamais supprimé (il contient
' la liste dont les événements ouvrent un résultat)
Private Sub ConstruireResultats()
    Dim wRes As Single
    wRes = UIBorne(W * 0.5, 380, 620)
    Set fraRes = UICtl(Me, "Forms.Frame.1", "fraRes", W - M - wRes, 29, wRes, 8 * UILigneH() + 54)
    fraRes.Caption = ""
    fraRes.BackColor = uiBlanc
    fraRes.SpecialEffect = 0: fraRes.BorderStyle = 1: fraRes.BorderColor = RGB(150, 160, 175)
    UILbl fraRes, "lblResTitre", "", M, 6, wRes - 2 * M - 30, UILigneH(-2), False, -2, uiGris
    Set btnFermerRes = UIBouton(fraRes, "btnFermerRes", "", wRes - M - 26, 3, 22, 18, 0, ICO_FERMER)
    Set lstRes = UIListe(fraRes, "lstRes", M, 26, wRes - 2 * M - 4, 8 * UILigneH() + 14, _
                         "52;140;" & CLng(wRes - 2 * M - 4 - 52 - 140 - 22))
    lstRes.BorderStyle = 0
    fraRes.Visible = False
End Sub

Private Sub Rechercher()
    Dim terme As String, q As Variant, chemin As String
    terme = Trim$(txtRecherche.Text)
    If terme = "" Then Exit Sub
    ' Identifiant exact : on y va directement
    If gQuestions.Exists(terme) Then
        If Not gQuestions(terme).Systeme Then
            FermerResultats
            Aller gQuestions(terme), ""
            Exit Sub
        End If
    End If
    Set resultats = New Collection
    For Each q In gOrdre
        If InStr(1, q.Id, terme, vbTextCompare) > 0 Or InStr(1, q.Libelle, terme, vbTextCompare) > 0 Then
            resultats.Add q
            If resultats.Count >= 200 Then Exit For
        End If
    Next q
    fraRes.Controls("lblResTitre").Caption = resultats.Count & IIf(resultats.Count >= 200, "+", "") & _
        " résultat(s) pour « " & terme & " »   ·   double-clic ou Entrée pour ouvrir, Échap pour fermer"
    lstRes.Clear
    For Each q In resultats
        chemin = Left$(gCategories(gCategories(gCategories(q.SousTheme).Parent).Parent).Libelle, 2) & " › " & _
                 gCategories(q.SousTheme).Libelle
        UIAjouterLigne lstRes, q.Id, chemin, IIf(q.Statut = ST_MASQUEE, "(sans objet) ", "") & _
                       Replace(Replace(q.Libelle, vbCr, " "), vbLf, " ")
    Next q
    If lstRes.ListCount > 0 Then lstRes.ListIndex = 0
    fraRes.Visible = True
    fraRes.ZOrder 0
    If lstRes.ListCount > 0 Then lstRes.SetFocus
End Sub

Private Sub FermerResultats()
    If Not fraRes Is Nothing Then fraRes.Visible = False
End Sub

Private Sub btnFermerRes_Click()
    FermerResultats
End Sub

Private Sub lstRes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    OuvrirResultat
End Sub

Private Sub lstRes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then KeyCode = 0: OuvrirResultat
    If KeyCode = 27 Then KeyCode = 0: FermerResultats
End Sub

Private Sub OuvrirResultat()
    Dim q As cQuestionSGO
    If lstRes.ListIndex < 0 Or resultats Is Nothing Then Exit Sub
    Set q = resultats(lstRes.ListIndex + 1)
    FermerResultats
    Aller q, ""
End Sub

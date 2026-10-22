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
Private WithEvents tsThemes As MSForms.TabStrip
Private WithEvents tsSST As MSForms.TabStrip
Private WithEvents cboFiltre As MSForms.ComboBox
Private WithEvents chkSansObjet As MSForms.CheckBox
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
Private fraChap As MSForms.Frame        ' liste des chapitres (lignes de Labels)
Private fraQ As MSForms.Frame           ' liste des questions (lignes de Labels)
Private lignes As Collection            ' cClicLigne des deux listes
Private fraRes As MSForms.Frame         ' résultats de recherche (créé une fois, masqué ou affiché)

' --- État -----------------------------------------------------------------------
Private curChap As cCategorieSGO
Private curTheme As cCategorieSGO
Private curSST As cCategorieSGO
Private curQ As cQuestionSGO
Private listeQ As Collection            ' questions affichées dans la liste
Private themesAff As Collection         ' thèmes du chapitre (onglets)
Private sstAff As Collection            ' sous-thèmes du thème (boutons)
Private resultats As Collection         ' résultats de recherche
Private chapSel As Long                 ' chapitre affiché (0 = premier)
Private qSel As Long                    ' question sélectionnée dans listeQ (0 = première, -1 : aucune)
Private chargement As Boolean           ' ignore les événements pendant un remplissage par le code
Private sale As Boolean                 ' la question courante a changé depuis le dernier enregistrement

' --- Disposition ------------------------------------------------------------------
Private M As Single, W As Single, H As Single      ' marge, largeur et hauteur utiles
Private wDet As Single, hDet As Single  ' taille utile des pages du détail
Private xIn As Single, yIn As Single, wIn As Single, yBas As Single, wQ As Single
Private hSST As Single, yZone As Single ' hauteur des onglets de sous-thèmes ; haut de la zone Questions

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

Private Sub UserForm_Terminate()
    LibererLignes "C"
    LibererLignes "Q"
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub

' =============================================================================
' Construction de la fenêtre
' =============================================================================
Private Sub Construire()
    Dim hEntete As Single, hPied As Single, yCorps As Single, hCorps As Single, wNav As Single
    Dim x0 As Single, wZone As Single
    Dim xD As Single, wR As Single, y As Single, c As Object, wDetail As Single

    chargement = True
    Set lignes = New Collection
    chapSel = -1: qSel = -1: yZone = -1
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
    UIFond Me, "lblBarreFond", W - M - wR, 36, wR - 170, 10, RGB(12, 36, 64)
    UIFond Me, "lblBarre", W - M - wR, 36, 1, 10, RGB(76, 175, 80)
    Set c = UILbl(Me, "lblAvancement", "", W - M - 164, 32, 164, UILigneH(-1), True, -1, uiBlanc)
    c.TextAlign = 3

    ' --- Chapitres ------------------------------------------------------------------
    UILbl Me, "lblChapTitre", "CHAPITRES", M + 4, yCorps, wNav - 8, UILigneH(-2), True, -2, uiGris
    Set fraChap = NouvelleListe("fraChap", M, yCorps + 18, wNav, hCorps - 18 - 48, uiOngletFond)
    Legende M + 2, yCorps + hCorps - 44, wNav - 4

    ' --- Thèmes (onglets) et sous-thèmes (boutons) -----------------------------------
    Set tsThemes = UICtl(Me, "Forms.TabStrip.1", "tsThemes", x0, yCorps, wZone, hCorps)
    tsThemes.Font.Name = UI_POLICE_SYMB
    tsThemes.TabFixedHeight = UILigneH() + 8
    Set c = UILbl(Me, "lblThemesTitre", "Thèmes", x0 + wZone - 96, yCorps + 5, 90, UILigneH(-2), False, -2, uiGris)
    c.TextAlign = 3
    UIFond Me, "lblZoneBlanche", x0 + 2, yCorps + UILigneH() + 12, wZone - 4, hCorps - UILigneH() - 14, uiBlanc
    xIn = x0 + M + 2: yIn = yCorps + UILigneH() + 12 + M: wIn = wZone - 2 * M - 4
    yBas = yCorps + hCorps - M

    hSST = UILigneH() + 12
    Set tsSST = UICtl(Me, "Forms.TabStrip.1", "tsSST", xIn, yIn, wIn, hSST)
    tsSST.Font.Name = UI_POLICE_SYMB
    tsSST.Style = 1                                  ' boutons
    tsSST.TabFixedHeight = UILigneH() + 8
    Set c = UILbl(Me, "lblSSTTitre", "Sous-thèmes", xIn + wIn - 96, yIn + 6, 90, UILigneH(-2), False, -2, uiGris)
    c.TextAlign = 3
    Set c = UILbl(Me, "lblMasque", "", xIn + 2, yIn, wIn, UILigneH(-2), False, -2, uiGris)
    c.Font.Name = UI_POLICE_SYMB

    ' --- Questions : le détail garde une largeur de lecture confortable, la liste
    '     prend le reste (les libellés passent à la ligne)
    wDetail = UIBorne(wIn * 0.5, 430, 700)
    wQ = UIBorne(wIn - 2 * M - wDetail, 230, 900)
    UISection Me, "secQuestions", "QUESTIONS", xIn, yIn, wQ
    UILbl Me, "lblFiltre", "Afficher", xIn, yIn + 27, 52, UILigneH()
    Set cboFiltre = UICombo(Me, "cboFiltre", xIn + 54, yIn + 24, UIBorne(wQ - 54 - 190, 110, 260), UILigneH() + 8, Split(FILTRES, ";"))
    cboFiltre.ListIndex = 0
    ' Afficher ou non les questions sans objet (grisées) ; réglage mémorisé, coché par défaut (ES-4)
    Set chkSansObjet = UICtl(Me, "Forms.CheckBox.1", "chkSansObjet", cboFiltre.Left + cboFiltre.Width + 10, yIn + 26, _
                             UIBorne(wQ - (cboFiltre.Left + cboFiltre.Width + 10 - xIn), 60, 200), UILigneH() + 4)
    chkSansObjet.Caption = "Questions sans objet"
    chkSansObjet.BackStyle = 0
    chkSansObjet.Value = (GetSetting("SGO", "Affichage", "SansObjet", "1") = "1")
    chkSansObjet.ControlTipText = "Afficher les questions sans objet (grisées, non modifiables)"
    UILbl Me, "lblCompteur", "", xIn, yIn + 24 + UILigneH() + 12, wQ, UILigneH(-2), False, -2, uiGris
    Set fraQ = NouvelleListe("fraQ", xIn, yIn + 60, wQ, 100, uiBlanc)
    Set btnRetablir = UIBouton(Me, "btnRetablir", "Rétablir les standards du sous-thème", xIn, yBas - 28, wQ, 28, 0, ICO_COCHE)

    ' --- Détail (fil d'Ariane au-dessus) ----------------------------------------------
    xD = xIn + wQ + 2 * M
    Set c = UILbl(Me, "lblAriane", "", xD, yIn, wIn - wQ - 2 * M, UILigneH(-1), False, -1, uiGris)
    c.WordWrap = False
    Set mp = UICtl(Me, "Forms.MultiPage.1", "mpDetail", xD, yIn, wIn - wQ - 2 * M, yBas - 38 - yIn)
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
    Set c = UILbl(Me, "lblEnregistre", "", M + 500, y + 7, W - M - 126 - 8 - (M + 500), UILigneH(-1), True, -1, uiVert)
    c.Font.Name = UI_POLICE_SYMB: c.TextAlign = 1: c.WordWrap = False

    PlacerZoneQuestions False, False
    ConstruireResultats
    RemplirChapitres
    AfficherAvancement
    chargement = False
End Sub

' Légende des pictogrammes, chacun dans sa couleur de statut
Private Sub Legende(ByVal x0 As Single, ByVal y0 As Single, ByVal largeur As Single)
    Dim items As Variant, coul As Variant, i As Long, x As Single, y As Single, w As Single, c As Object
    items = Array(ChrW(&H2713) & " conforme", "! à justifier", ChrW(&H25C6) & " spécifique", _
                  ChrW(&H25CB) & " à saisir", ChrW(&H21BB) & " standard modifié", ChrW(&H2298) & " sans objet")
    coul = Array(uiVert, uiRouge, uiBleuStatut, uiGris, uiOrange, uiGrisClair)
    x = x0: y = y0
    For i = 0 To UBound(items)
        w = Len(items(i)) * (uiTaille - 2) * 0.52 + 10
        If x + w > x0 + largeur Then x = x0: y = y + UILigneH(-2) + 2
        Set c = UILbl(Me, "lblLeg" & i, CStr(items(i)), x, y, w, UILigneH(-2), (i = 1), -2, CLng(coul(i)))
        c.Font.Name = UI_POLICE_SYMB: c.WordWrap = False
        If i = 5 Then c.ControlTipText = "Sans objet : dépend de la réponse à une question mère, ou sans standard pour votre direction/univers"
        x = x + w
    Next i
End Sub

' « 03 Montage contractuel › 03.2 Modalités d'intervention › 03.2.2 Offre non exécutée › El78b »
Private Function FilAriane() As String
    Dim sst As cCategorieSGO, th As cCategorieSGO, ch As cCategorieSGO, t As String
    If curQ Is Nothing Then Exit Function
    Set sst = gCategories(curQ.SousTheme)
    Set th = gCategories(sst.Parent)
    Set ch = gCategories(th.Parent)
    t = ch.Libelle & "  " & ChrW(&H203A) & "  "
    If th.Libelle <> sst.Libelle And Left$(th.Id, 4) <> "THM_" Then t = t & th.Libelle & "  " & ChrW(&H203A) & "  "
    If Left$(sst.Id, 4) <> "SST_" Then t = t & sst.Libelle & "  " & ChrW(&H203A) & "  "
    FilAriane = t & curQ.Id
End Function

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
    Dim c As Variant, i As Long, y As Single
    LibererLignes "C"
    fraChap.Controls.Clear
    y = 2
    For Each c In gChapitres
        y = y + AjouterLigneListe(fraChap, "C", i, y, UIPictoCategorie(c), UICouleurCategorie(c), "", c.Libelle, _
                                  IIf(c.Grisee(), "", c.Pourcentage() & " %"), 2, c.Grisee(), c.Libelle, _
                                  IIf(c.Grisee(), -1, c.Pourcentage()))
        i = i + 1
    Next c
    AjusterDefilement fraChap, y
    MarquerChapitre chapSel
End Sub

' Met à jour pictogrammes, pourcentages et grisé des chapitres, sans reconstruire
Private Sub RafraichirChapitres()
    Dim c As Variant, i As Long
    For Each c In gChapitres
        fraChap.Controls("C" & i & "_ico").Caption = UIPictoCategorie(c)
        fraChap.Controls("C" & i & "_ico").Tag = CStr(UICouleurCategorie(c))
        fraChap.Controls("C" & i & "_dr").Caption = IIf(c.Grisee(), "", c.Pourcentage() & " %")
        MiniBarre "C" & i, IIf(c.Grisee(), -1, c.Pourcentage())
        CouleursLigne fraChap, "C" & i, (i = chapSel), c.Grisee()
        i = i + 1
    Next c
End Sub

Private Sub MarquerChapitre(ByVal index As Long)
    Dim i As Long
    chapSel = index
    For i = 0 To gChapitres.Count - 1
        CouleursLigne fraChap, "C" & i, (i = index), gChapitres(i + 1).Grisee()
    Next i
End Sub

' Met à jour les pictogrammes et pourcentages sans reconstruire la navigation
Private Sub RafraichirNavigation()
    Dim c As Variant, i As Long, etaitEnChargement As Boolean
    etaitEnChargement = chargement
    chargement = True
    RafraichirChapitres
    EtatOnglets tsThemes, themesAff
    EtatOnglets tsSST, sstAff
    RemplirQuestions False
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
    Me.Controls("lblAvancement").Caption = pct & " %  ·  " & a(0) & " / " & a(1) & " traitées"
End Sub

' Clic sur une ligne d'une liste maison (appelé par cClicLigne)
Public Sub LigneCliquee(ByVal liste As String, ByVal index As Long)
    If chargement Then Exit Sub
    If liste = "C" Then
        If index = chapSel Then Exit Sub
        Valider
        AfficherChapitre gChapitres(index + 1), Nothing
    Else
        If index >= listeQ.Count Then Exit Sub
        If listeQ(index + 1) Is curQ Then Exit Sub
        Valider
        MarquerQuestion index
        AfficherQuestion listeQ(index + 1)
    End If
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
    chkSansObjet.Enabled = (cboFiltre.ListIndex <= 0)    ' les autres filtres excluent déjà les questions sans objet
    AppliquerFiltre
End Sub

Private Sub chkSansObjet_Click()
    If chargement Then Exit Sub
    SaveSetting "SGO", "Affichage", "SansObjet", IIf(chkSansObjet.Value, "1", "0")
    AppliquerFiltre
End Sub

' Reconstruit la liste ; garde la question courante si elle y est encore
Private Sub AppliquerFiltre()
    Dim i As Long, cible As Long
    Valider
    chargement = True
    RemplirQuestions True
    chargement = False
    cible = -1
    For i = 1 To listeQ.Count
        If listeQ(i) Is curQ Then cible = i - 1
    Next i
    If cible >= 0 Then Exit Sub
    If listeQ.Count > 0 Then
        MarquerQuestion 0
        AfficherQuestion listeQ(1)
    Else
        AfficherQuestion Nothing
    End If
End Sub

' Affiche une question sans objet même si la case est décochée (recherche, lien direct)
Private Sub MontrerSansObjet()
    If chkSansObjet.Value Then Exit Sub
    chargement = True
    chkSansObjet.Value = True
    chargement = False
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
        If gChapitres(i) Is chap Then MarquerChapitre i - 1
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
    PlacerZoneQuestions sstAff.Count > 1, sst.Grisee()
    chargement = True
    Set curQ = Nothing
    RemplirQuestions True
    cible = -1
    If Not q Is Nothing Then
        For i = 1 To listeQ.Count
            If listeQ(i) Is q Then cible = i - 1
        Next i
        If cible < 0 Then                            ' filtrée : on revient à « Toutes », sans objet compris
            cboFiltre.ListIndex = 0
            chkSansObjet.Enabled = True
            If q.Statut = ST_MASQUEE Then MontrerSansObjet
            chargement = True
            RemplirQuestions True
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
    MarquerQuestion cible
    chargement = False
    If cible >= 0 Then
        AfficherQuestion listeQ(cible + 1)
    Else
        AfficherQuestion Nothing
    End If
End Sub

' Liste des questions du sous-thème, selon le filtre. Sans « reconstruire », la
' liste n'est reconstruite que si ses lignes changent (sinon : mise à jour sur place)
Private Sub RemplirQuestions(ByVal reconstruire As Boolean)
    Dim q As Variant, nMasq As Long, nTot As Long, i As Long, nouvelle As Collection, y As Single
    Dim identique As Boolean, lib As String
    If curSST Is Nothing Then Exit Sub
    Set nouvelle = New Collection
    For Each q In curSST.Questions
        nTot = nTot + 1
        If q.Statut = ST_MASQUEE Then nMasq = nMasq + 1
        If PasseFiltre(q) Then nouvelle.Add q
    Next q
    If cboFiltre.ListIndex <= 0 And nMasq > 0 And Not chkSansObjet.Value Then
        Me.Controls("lblCompteur").Caption = nTot & " question" & IIf(nTot > 1, "s", "") & ", dont " & nMasq & _
            " sans objet masquée" & IIf(nMasq > 1, "s", "")
    ElseIf cboFiltre.ListIndex <= 0 Then
        Me.Controls("lblCompteur").Caption = nTot & " question" & IIf(nTot > 1, "s", "") & _
            IIf(nMasq > 0, ", dont " & nMasq & " sans objet (grisée" & IIf(nMasq > 1, "s", "") & ", non modifiable" & IIf(nMasq > 1, "s", "") & ")", "")
    Else
        Me.Controls("lblCompteur").Caption = nouvelle.Count & " sur " & nTot & " question" & IIf(nTot > 1, "s", "")
    End If

    ' Mêmes lignes, même grisé : simple mise à jour des pictogrammes
    identique = Not reconstruire And Not listeQ Is Nothing
    If identique Then identique = (nouvelle.Count = listeQ.Count)
    If identique Then
        For i = 1 To nouvelle.Count
            If Not nouvelle(i) Is listeQ(i) Or fraQ.Controls("Q" & (i - 1) & "_txt").ForeColor <> CouleurTexte(nouvelle(i)) Then
                identique = False: Exit For
            End If
        Next i
    End If
    If identique Then
        For i = 1 To listeQ.Count
            fraQ.Controls("Q" & (i - 1) & "_ico").Caption = UIPicto(listeQ(i).Statut, listeQ(i).StandardModifie)
            fraQ.Controls("Q" & (i - 1) & "_ico").Tag = CStr(UICouleurStatut(listeQ(i).Statut, listeQ(i).StandardModifie))
            CouleursLigne fraQ, "Q" & (i - 1), (i - 1 = qSel), listeQ(i).Statut = ST_MASQUEE
        Next i
        Exit Sub
    End If

    Set listeQ = nouvelle
    LibererLignes "Q"
    fraQ.Controls.Clear
    fraQ.ScrollTop = 0
    y = 2
    For i = 1 To listeQ.Count
        Set q = listeQ(i)
        lib = Replace(Replace(q.Libelle, vbCr, " "), vbLf, " ")
        If q.Statut = ST_MASQUEE Then lib = "(" & RaisonMasquage(q.Id) & ") " & lib
        y = y + AjouterLigneListe(fraQ, "Q", i - 1, y, UIPicto(q.Statut, q.StandardModifie), _
                                  UICouleurStatut(q.Statut, q.StandardModifie), q.Id, lib, "", 3, _
                                  q.Statut = ST_MASQUEE, q.Id & " · " & Replace(Replace(q.Libelle, vbCr, " "), vbLf, " "), -1)
    Next i
    AjusterDefilement fraQ, y
    qSel = -1
    If Not curQ Is Nothing Then
        For i = 1 To listeQ.Count
            If listeQ(i) Is curQ Then MarquerQuestion i - 1
        Next i
    End If
End Sub

Private Function CouleurTexte(ByVal q As cQuestionSGO) As Long
    CouleurTexte = IIf(q.Statut = ST_MASQUEE, RGB(150, 156, 164), uiTexte)
End Function

Private Sub MarquerQuestion(ByVal index As Long)
    Dim i As Long, haut As Single, bas As Single
    If qSel >= 0 And qSel < listeQ.Count Then CouleursLigne fraQ, "Q" & qSel, False, listeQ(qSel + 1).Statut = ST_MASQUEE
    qSel = index
    If index < 0 Or index >= listeQ.Count Then Exit Sub
    CouleursLigne fraQ, "Q" & index, True, listeQ(index + 1).Statut = ST_MASQUEE
    ' La ligne choisie reste visible
    haut = fraQ.Controls("Q" & index & "_fond").Top
    bas = haut + fraQ.Controls("Q" & index & "_fond").Height
    If haut < fraQ.ScrollTop Then
        fraQ.ScrollTop = haut
    ElseIf bas > fraQ.ScrollTop + fraQ.InsideHeight Then
        fraQ.ScrollTop = bas - fraQ.InsideHeight
    End If
End Sub

' -----------------------------------------------------------------------------
' Listes « maison » : lignes de Labels dans un Frame qui défile
' -----------------------------------------------------------------------------
Private Function NouvelleListe(nom As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                               ByVal h As Single, ByVal fond As Long) As MSForms.Frame
    Dim f As MSForms.Frame
    Set f = UICtl(Me, "Forms.Frame.1", nom, l, t, w, h)
    f.Caption = ""
    f.BackColor = fond
    f.SpecialEffect = 0
    f.BorderStyle = 1
    f.BorderColor = RGB(190, 197, 206)
    f.KeepScrollBarsVisible = 0
    Set NouvelleListe = f
End Function

' Ajoute une ligne (pictogramme coloré, identifiant, texte sur maxLignes lignes,
' texte à droite, mini-barre d'avancement si pct >= 0) ; renvoie sa hauteur
Private Function AjouterLigneListe(fra As MSForms.Frame, ByVal liste As String, ByVal i As Long, ByVal y As Single, _
                                   ByVal picto As String, ByVal coulPicto As Long, ByVal id As String, ByVal texte As String, _
                                   ByVal droite As String, ByVal maxLignes As Long, ByVal grisee As Boolean, _
                                   ByVal bulle As String, ByVal pct As Long) As Single
    Dim nom As String, w As Single, xTxt As Single, wTxt As Single, wDr As Single, nb As Long, hl As Single
    Dim c As Object, ligne As cClicLigne
    nom = liste & i
    w = fra.Width - 18                                  ' place pour la barre de défilement
    wDr = IIf(droite = "", 0, 40)
    xTxt = 26 + IIf(id = "", 0, 50)
    wTxt = w - xTxt - wDr - 4
    nb = LignesTexte(texte, wTxt, 0)
    If nb > maxLignes Then nb = maxLignes
    hl = nb * UILigneH() + 8 + IIf(liste = "C", 6, 0)

    UIFond fra, nom & "_fond", 0, y, w, hl, fra.BackColor
    UIFond fra, nom & "_barre", 0, y, IIf(liste = "C", 5, 4), hl, uiAction
    UIFond fra, nom & "_sep", 4, y + hl - 1, w - 8, 1, uiLigne
    Set c = UILbl(fra, nom & "_ico", picto, 8, y + 4, 16, UILigneH(), True)
    c.Font.Name = UI_POLICE_SYMB
    c.Tag = CStr(coulPicto)
    If id <> "" Then
        Set c = UILbl(fra, nom & "_id", id, 26, y + 4, 48, UILigneH(), True, -1)
        c.WordWrap = False
    End If
    UILbl fra, nom & "_txt", texte, xTxt, y + 4, wTxt, nb * UILigneH()
    If liste = "C" Then
        ' Mini-barre d'avancement sous le nom du chapitre
        UIFond fra, nom & "_pbf", xTxt, y + 4 + nb * UILigneH() + 2, wTxt * 0.9, 3, RGB(207, 213, 221)
        UIFond fra, nom & "_pb", xTxt, y + 4 + nb * UILigneH() + 2, 1, 3, RGB(76, 175, 80)
        MiniBarre nom, pct
    End If
    If droite <> "" Then
        Set c = UILbl(fra, nom & "_dr", droite, w - wDr - 2, y + 4, wDr, UILigneH(), False, -1, uiGris)
        c.TextAlign = 3
    Else
        UILbl fra, nom & "_dr", "", w - 2, y + 4, 1, UILigneH()
    End If
    ' Zone cliquable transparente au-dessus de la ligne
    Set ligne = New cClicLigne
    Set ligne.Zone = UILbl(fra, nom & "_zone", "", 0, y, w, hl)
    ligne.Zone.ControlTipText = Left$(bulle, 250)
    Set ligne.Formulaire = Me
    ligne.Liste = liste
    ligne.Index = i
    lignes.Add ligne, nom
    CouleursLigne fra, nom, False, grisee
    AjouterLigneListe = hl
End Function

' Couleurs d'une ligne : chapitre ouvert sur fond blanc (il « sort » de la colonne
' grise), question choisie sur fond bleu pâle ; barre bleue à gauche ; pictogramme
' à la couleur de son statut (gardée dans Tag)
Private Sub CouleursLigne(fra As MSForms.Frame, ByVal nom As String, ByVal choisie As Boolean, ByVal grisee As Boolean)
    Dim coul As Long, estChap As Boolean
    On Error Resume Next
    estChap = (fra Is fraChap)
    coul = IIf(grisee, uiGrisClair, uiTexte)
    If choisie Then
        fra.Controls(nom & "_fond").BackColor = IIf(estChap, uiBlanc, uiBleuClair)
    Else
        fra.Controls(nom & "_fond").BackColor = fra.BackColor
    End If
    fra.Controls(nom & "_barre").Visible = choisie
    fra.Controls(nom & "_txt").ForeColor = coul
    fra.Controls(nom & "_txt").Font.Bold = choisie And estChap
    fra.Controls(nom & "_id").ForeColor = IIf(grisee, coul, IIf(choisie, uiAction, uiBandeau))
    fra.Controls(nom & "_ico").ForeColor = IIf(grisee, uiGrisClair, Val(fra.Controls(nom & "_ico").Tag))
End Sub

' Mini-barre d'avancement d'un chapitre (pct < 0 : masquée)
Private Sub MiniBarre(ByVal nom As String, ByVal pct As Long)
    Dim f As Object, b As Object
    On Error Resume Next
    Set f = fraChap.Controls(nom & "_pbf")
    Set b = fraChap.Controls(nom & "_pb")
    f.Visible = (pct >= 0): b.Visible = (pct > 0)
    If pct > 0 Then
        b.Width = UIBorne(f.Width * pct / 100, 1, f.Width)
        b.BackColor = IIf(pct >= 100, uiVert, RGB(76, 175, 80))
    End If
End Sub

Private Sub AjusterDefilement(fra As MSForms.Frame, ByVal hauteur As Single)
    If hauteur > fra.InsideHeight Then
        fra.ScrollBars = 2
        fra.ScrollHeight = hauteur + 2
    Else
        fra.ScrollBars = 0
        fra.ScrollTop = 0
    End If
End Sub

' Supprime les objets cClicLigne d'une liste (avant de vider le Frame)
Private Sub LibererLignes(ByVal liste As String)
    Dim i As Long
    If lignes Is Nothing Then Exit Sub
    For i = lignes.Count To 1 Step -1
        If lignes(i).Liste = liste Then
            Set lignes(i).Formulaire = Nothing
            Set lignes(i).Zone = Nothing
            lignes.Remove i
        End If
    Next i
End Sub

' Place la zone Questions / Détail juste sous les sous-thèmes, ou en haut quand
' le thème n'a qu'un sous-thème (la bande reste sinon vide)
Private Sub PlacerZoneQuestions(ByVal avecSST As Boolean, ByVal avecMasque As Boolean)
    Dim y As Single, yListe As Single
    y = yIn
    If avecSST Then y = y + hSST + 4
    Me.Controls("lblMasque").Top = y
    If avecMasque Then y = y + UILigneH(-2) + 6
    If Abs(y - yZone) < 0.5 Then Exit Sub
    yZone = y
    Me.Controls("secQuestions").Top = y
    Me.Controls("secQuestions_trait").Top = y + UILigneH(-2) + 3
    Me.Controls("lblFiltre").Top = y + 27
    cboFiltre.Top = y + 24
    chkSansObjet.Top = y + 26
    Me.Controls("lblCompteur").Top = y + 24 + UILigneH() + 12
    yListe = y + 24 + UILigneH() + 12 + UILigneH(-2) + 4
    fraQ.Top = yListe
    fraQ.Height = UIBorne(yBas - 36 - yListe, 60, 2000)
    Me.Controls("lblAriane").Top = y
    mp.Top = y + UILigneH(-1) + 4
    mp.Height = UIBorne(yBas - 38 - mp.Top, 120, 2000)
    hDet = mp.Height - mp.TabFixedHeight - 10
End Sub

Private Function PasseFiltre(ByVal q As cQuestionSGO) As Boolean
    Select Case cboFiltre.ListIndex
        Case 1: PasseFiltre = (q.Statut = ST_NON_REPONDUE)
        Case 2: PasseFiltre = (q.Statut = ST_A_JUSTIFIER)
        Case 3: PasseFiltre = q.StandardModifie
        Case 4: PasseFiltre = (q.Conformite = CONF_NON_CONFORME)
        Case Else: PasseFiltre = (q.Statut <> ST_MASQUEE Or chkSansObjet.Value)
    End Select
End Function


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
        If Not PasseFiltre(q) Then
            chargement = True: cboFiltre.ListIndex = 0: chargement = False
            chkSansObjet.Enabled = True
            If q.Statut = ST_MASQUEE Then MontrerSansObjet
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
    Me.Controls("lblAriane").Caption = FilAriane()
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
    Dim masquee As Boolean, i As Long, sp As Boolean, avecStd As Boolean, lesFilles As String, x As Single

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

    ' Identifiant et mode en étiquettes, puis condition et questions conditionnées
    x = Etiquette(pg, "lblPillId", curQ.Id, M, 8)
    x = Etiquette(pg, "lblPillMode", curQ.Mode, x + 4, 8)
    txt = ""
    If RaisonAffichage(curQ.Id) <> "" And Not masquee Then txt = RaisonAffichage(curQ.Id)
    If gFilles.Exists(curQ.Id) Then
        lesFilles = gFilles(curQ.Id)
        txt = txt & IIf(txt = "", "", "  ·  ") & "Conditionne : " & lesFilles
    End If
    hz = LignesTexte(txt, iw - (x - M) - 8, -2) * UILigneH(-2)
    Set c = UILbl(pg, "lblTypeQ", txt, x + 8, 10, iw - (x - M) - 8, hz, False, -2, uiGris)
    If lesFilles <> "" Then c.ControlTipText = "Votre réponse fait apparaître ou masque : " & lesFilles
    y = 8 + UIBorne(hz, UILigneH(-2) + 4, 200) + 6
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
        y = Encadre(pg, y, iw, ChrW(&H25A3) & "  STANDARD APPLICABLE  ·  " & UCase$(LibelleNiveau(curQ.Niveau)), curQ.Standard, _
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
    UISection pg, "secReponse", ChrW(&H270E) & "  RÉPONSE", M, y, iw, True
    Set c = UILbl(pg, "lblModifiee", " MODIFIÉE PAR RAPPORT AU STANDARD ", M + 90, y - 1, 210, UILigneH(-3) + 2, True, -3, RGB(122, 90, 0))
    c.BackStyle = 1: c.BackColor = RGB(255, 243, 205): c.BorderStyle = 1: c.BorderColor = RGB(230, 199, 102)
    c.TextAlign = 2: c.WordWrap = False
    c.Visible = ReponseModifiee()
    y = y + 26
    If curQ.EstAnnexe Then
        UILbl pg, "lblRepAnnexe", curQ.Reponse, M, y, iw, UILigneH(), True
        y = y + UILigneH() + 12
    ElseIf curQ.EstListe And curQ.NbChoix > 0 And InStr(1, curQ.Mode, "texte", vbTextCompare) = 0 Then
        ' Jusqu'à 12 choix affichés en entier ; au-delà, la liste défile
        hz = UIBorne(curQ.NbChoix, 2, 12) * UILigneH() + 8
        Set lstChoix = UIListe(pg, "lstChoix", M, y, iw, hz, "")
        lstChoix.ListStyle = 1                       ' cases à cocher / boutons d'option
        lstChoix.MultiSelect = IIf(curQ.EstChoixMultiple, 1, 0)
        For i = 0 To UBound(curQ.Choix)
            lstChoix.AddItem curQ.Choix(i)
        Next i
        SelectionnerChoix
        lstChoix.Enabled = Not masquee
        y = y + hz + 4
        If curQ.NbChoix > 12 Then
            UILbl pg, "lblDefiler", curQ.NbChoix & " choix : faites défiler la liste pour les voir tous.", _
                  M, y, iw, UILigneH(-2), False, -2, uiOrange
            y = y + UILigneH(-2) + 4
        End If
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
        txt = IIf(sp, "JUSTIFICATION DU CAS DIFFÉRENCIANT", "JUSTIFICATION DE LA DÉROGATION AU STANDARD")
        UISection pg, "secJustif", ChrW(&H270D) & "  " & txt, M, y, iw, True
        Set c = UILbl(pg, "lblOblig", "obligatoire", M + 24 + Len(txt) * (uiTaille - 2) * 0.66, y, 90, UILigneH(-2), True, -2, RGB(176, 107, 0))
        c.WordWrap = False
        c.Visible = (curQ.Statut = ST_A_JUSTIFIER)
        y = y + 26
        hz = UIBorne(LignesTexte(curQ.Justification, iw, 0) + 1, 3, 8) * UILigneH() + 8
        Set txtJustif = UITxt(pg, "txtJustif", curQ.Justification, M, y, iw, hz)
        txtJustif.MultiLine = True: txtJustif.WordWrap = True: txtJustif.ScrollBars = 2
        txtJustif.EnterKeyBehavior = True
        StyleJustification
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
    UIFond pg, "lblEncBarre", M, y, 5, hz + UILigneH(-2) + 22, barre
    Set c = UILbl(pg, "lblEncTitre", titre, M + 14, y + 7, iw - 28, UILigneH(-2), True, -2, coulTitre)
    c.Font.Name = UI_POLICE_SYMB
    Set c = UILbl(pg, "lblEncTexte", texte, M + 14, y + 10 + UILigneH(-2), wTexte, hz)
    c.Font.Italic = boutonReprendre             ' la valeur du standard, en italique
    If boutonReprendre Then
        Set btnReprendre = UIBouton(pg, "btnReprendre", ChrW(&H21BA) & " Reprendre le standard", M + iw - 152, y + 10 + UILigneH(-2), 144, UILigneH() + 10, 0)
        btnReprendre.ControlTipText = "Remplace votre réponse par la valeur du standard"
    End If
    Encadre = y + hz + UILigneH(-2) + 22 + 16
End Function

Private Function BandeauConformite(ByVal pg As Object, ByVal y As Single, ByVal iw As Single) As Single
    Dim hz As Single, c As Object
    hz = UILigneH(0.5) + 3 * UILigneH(-1) + 12
    UIFond pg, "lblConfFond", M, y, iw, hz, uiCoulFond
    pg.Controls("lblConfFond").BorderStyle = 1
    UIFond pg, "lblConfBarre", M, y, 6, hz, uiGris
    Set c = UILbl(pg, "lblConfIco", "", M + 12, y + 6, 26, UILigneH(8), True, 8)
    c.Font.Name = UI_POLICE_SYMB
    Set c = UILbl(pg, "lblConformite", "", M + 44, y + 6, iw - 52, UILigneH(0.5), True, 0.5)
    Set c = UILbl(pg, "lblConfInfo", "", M + 44, y + 8 + UILigneH(0.5), iw - 52, 3 * UILigneH(-1), False, -1)
    MettreAJourBandeau
    BandeauConformite = y + hz + 14
End Function

' Couleurs et textes du bandeau, selon le statut (mise à jour sur place)
Private Sub MettreAJourBandeau()
    Dim ico As String, titre As String, info As String, fond As Long, bord As Long, coul As Long, txtCoul As Long
    Dim pg As Object, compare As String
    Set pg = mp.Pages(0)
    compare = IIf(curQ.Operateur = "=", "n'est pas égale au", "ne contient pas le")
    Select Case curQ.Statut
        Case ST_CONFORME
            ico = ChrW(&H2713): titre = "Conforme au standard"
            info = "Votre réponse " & IIf(curQ.Operateur = "=", "est égale au", "contient le") & " standard."
            coul = uiVert: fond = uiVertClair: bord = RGB(167, 214, 180)
        Case ST_A_JUSTIFIER, ST_JUSTIFIEE
            If curQ.Conformite = CONF_SPECIFIQUE Then
                ico = ChrW(&H25C6): coul = uiBleuStatut: fond = uiBleuClair: bord = RGB(156, 195, 230)
                titre = "Spécifique à l'offre · différente de l'offre précédente"
                info = IIf(curQ.Statut = ST_A_JUSTIFIER, "Justification obligatoire : expliquez la différence ci-dessous.", _
                           "Différence justifiée.")
            Else
                ico = ChrW(&H2717): coul = uiRouge: fond = uiRougeClair: bord = RGB(241, 184, 177)
                titre = "Non conforme au standard"
                info = "Votre réponse " & compare & " standard « " & Abrege(Ligne1(curQ.Standard), 110) & " »." & _
                       IIf(curQ.Statut = ST_A_JUSTIFIER, " Expliquez la dérogation ci-dessous.", " Dérogation justifiée.")
            End If
            If curQ.Statut = ST_A_JUSTIFIER Then titre = titre & " · justification obligatoire"
        Case ST_SPECIFIQUE
            ico = ChrW(&H25C6): coul = uiBleuStatut: fond = uiBleuClair: bord = RGB(156, 195, 230)
            titre = "Spécifique à l'offre"
            info = "Aucun standard pour cette question : réponse propre à l'offre, identique à l'offre précédente (à vérifier)."
        Case ST_NON_CONCERNEE
            ico = ChrW(&H2013): coul = uiGris: fond = uiCoulFond: bord = uiLigne
            titre = "Non concerné"
            info = "Cette question est sans objet pour l'offre."
        Case ST_INCOMPLET
            ico = "?": coul = uiOrange: fond = uiStdFond: bord = RGB(230, 199, 102)
            titre = "Standards incomplets"
            info = "Aucun standard n'est défini pour votre direction et votre univers : contactez votre administrateur."
        Case Else
            ico = ChrW(&H25CB): coul = uiGris: fond = uiCoulFond: bord = uiLigne
            titre = "Réponse à saisir"
            info = IIf(curQ.Niveau = "Sp", "Aucun standard pour cette question : saisissez la réponse propre à l'offre.", _
                       "Saisissez une réponse.")
    End Select
    txtCoul = IIf(coul = uiRouge, RGB(107, 42, 36), uiTexte)
    On Error Resume Next
    pg.Controls("lblConfFond").BackColor = fond
    pg.Controls("lblConfFond").BorderColor = bord
    pg.Controls("lblConfBarre").BackColor = coul
    pg.Controls("lblConfIco").Caption = ico
    pg.Controls("lblConfIco").ForeColor = coul
    pg.Controls("lblConformite").Caption = titre
    pg.Controls("lblConformite").ForeColor = coul
    pg.Controls("lblConfInfo").Caption = info
    pg.Controls("lblConfInfo").ForeColor = txtCoul
End Sub

Private Function Abrege(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then Abrege = Left$(s, n - 1) & ChrW(&H2026) Else Abrege = s
End Function

' Réponse saisie différente du standard (pastille « modifiée »)
Private Function ReponseModifiee() As Boolean
    If curQ Is Nothing Then Exit Function
    If curQ.Niveau = "E" Or curQ.Niveau = "D" Or curQ.Niveau = "U" Then
        ReponseModifiee = curQ.Saisie And StrComp(Trim$(curQ.Reponse), Trim$(curQ.Standard), vbTextCompare) <> 0
    End If
End Function

' Zone de justification : fond jaune pâle et bordure orange quand elle est obligatoire
Private Sub StyleJustification()
    If txtJustif Is Nothing Then Exit Sub
    If curQ.Statut = ST_A_JUSTIFIER Then
        txtJustif.BackColor = uiJauneClair
        txtJustif.BorderColor = uiOrangeVif
    Else
        txtJustif.BackColor = uiBlanc
        txtJustif.BorderColor = RGB(170, 178, 189)
    End If
End Sub

' Petite étiquette grise (identifiant, mode de réponse) ; renvoie sa position droite
Private Function Etiquette(ByVal pg As Object, nom As String, texte As String, ByVal x As Single, ByVal y As Single) As Single
    Dim c As Object, w As Single
    w = Len(texte) * (uiTaille - 2) * 0.55 + 12
    Set c = UILbl(pg, nom, texte, x, y, w, UILigneH(-2) + 4, False, -2, RGB(68, 68, 68))
    c.BackStyle = 1: c.BackColor = RGB(238, 241, 245)
    c.BorderStyle = 1: c.BorderColor = uiLigne
    c.TextAlign = 2: c.WordWrap = False
    Etiquette = x + w
End Function

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
    Dim pg As Object
    Set pg = mp.Pages(0)
    On Error Resume Next
    pg.Controls("lblRepComplete").Caption = "Réponse retenue : " & IIf(curQ.Reponse = "", "(aucune)", curQ.Reponse)
    If Not curQ.StandardModifie Then pg.Controls("lblStdModifie").Caption = ""
    pg.Controls("lblModifiee").Visible = ReponseModifiee()
    pg.Controls("lblOblig").Visible = (curQ.Statut = ST_A_JUSTIFIER)
    On Error GoTo 0
    MettreAJourBandeau
    StyleJustification
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
    Me.Controls("lblEnregistre").Caption = ChrW(&H2714) & " " & curQ.Id & " enregistrée à " & Format$(Now, "hh:nn")
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

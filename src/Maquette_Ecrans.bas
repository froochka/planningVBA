Option Explicit

' =============================================================================
' Maquette visuelle des écrans (v3, sans interaction, données d'exemple réelles).
' Uniquement des contrôles MSForms standards : Label, Frame, ListBox, ComboBox,
' TextBox, OptionButton, TabStrip, MultiPage. Aucun ActiveX, aucune image :
' les icônes viennent de la police Windows « Segoe MDL2 Assets ».
'
'   Alt+F8 > MaquetteEcranSaisie : fenêtre principale (chapitres > thèmes > sous-thèmes > question)
'   Alt+F8 > MaquetteRecherche   : même fenêtre, panneau de résultats de recherche ouvert
'   Alt+F8 > MaquettePageGarde   : page de garde (offre, standards, lots)
'   Alt+F8 > MaquetteNettoyer    : à lancer avant d'enregistrer le classeur
'
' Un UserForm vide « frmMaquetteTmp » est ajouté au classeur au premier lancement
' et réutilisé ensuite (VBA ne sait pas recréer un formulaire supprimé dans la
' même session). Nécessite « Accès approuvé au modèle d'objet du projet VBA ».
'
' Taille du texte : TAILLE_BASE pilote toutes les polices et les hauteurs
' (dans l'application : boutons A+ / A-, réglage mémorisé par utilisateur).
' =============================================================================

Private Const NOM_FORM As String = "frmMaquetteTmp"
Private Const POLICE As String = "Segoe UI"
Private Const POLICE_SYMB As String = "Segoe UI Symbol"     ' pictogrammes de statut
Private Const POLICE_ICONES As String = "Segoe MDL2 Assets"  ' icônes Windows 10/11
Private Const TAILLE_BASE As Single = 10
Private Const M As Single = 8                               ' marge (grille de 8 pt)

' Icônes Segoe MDL2 Assets
Private Const ICO_RECHERCHE As Long = &HE721&
Private Const ICO_CRAYON As Long = &HE70F&
Private Const ICO_EXPORT As Long = &HEDE1&
Private Const ICO_PRECEDENT As Long = &HE76B&
Private Const ICO_SUIVANT As Long = &HE76C&
Private Const ICO_COCHE As Long = &HE73E&

' Palette sobre (à remplacer par la charte de l'organisme)
Private cBandeau As Long, cAction As Long, cFond As Long, cBlanc As Long, cLigne As Long
Private cTexte As Long, cGris As Long, cVert As Long, cOrange As Long
Private cStdFond As Long, cStdBarre As Long, cOngletFond As Long

Private etape As String          ' pour situer une éventuelle erreur

Private Sub InitCouleurs()
    cBandeau = RGB(22, 54, 92)
    cAction = RGB(0, 99, 177)
    cFond = RGB(243, 245, 248)
    cBlanc = RGB(255, 255, 255)
    cLigne = RGB(214, 219, 226)
    cTexte = RGB(33, 37, 41)
    cGris = RGB(100, 108, 118)
    cVert = RGB(24, 128, 56)
    cOrange = RGB(160, 100, 0)
    cStdFond = RGB(255, 249, 230)
    cStdBarre = RGB(214, 158, 0)
    cOngletFond = RGB(232, 236, 241)
End Sub

' -----------------------------------------------------------------------------
' Points d'entrée
' -----------------------------------------------------------------------------
Public Sub MaquetteEcranSaisie()
    AfficherEcranSaisie False
End Sub

Public Sub MaquetteRecherche()
    AfficherEcranSaisie True
End Sub

Public Sub MaquettePageGarde()
    Dim f As Object
    InitCouleurs
    Set f = CreerForm("SGO - Page de garde du dossier   (maquette visuelle)")
    If f Is Nothing Then Exit Sub
    On Error GoTo fin
    ConstruirePageGarde f
    f.Show
fin:
    If Err.Number <> 0 Then MsgBox "Erreur maquette (" & etape & ") : " & Err.Description, vbExclamation
    Set f = Nothing
End Sub

Private Sub AfficherEcranSaisie(ByVal avecRecherche As Boolean)
    Dim f As Object
    InitCouleurs
    Set f = CreerForm("SGO - Élaboration du DCO   (maquette visuelle - fermer avec la croix)")
    If f Is Nothing Then Exit Sub
    On Error GoTo fin
    ConstruireEcranSaisie f, avecRecherche
    f.Show
fin:
    If Err.Number <> 0 Then MsgBox "Erreur maquette (" & etape & ") : " & Err.Description, vbExclamation
    Set f = Nothing
End Sub

' -----------------------------------------------------------------------------
' Écran principal : une fenêtre, navigation à trois niveaux + détail
' -----------------------------------------------------------------------------
Private Sub ConstruireEcranSaisie(f As Object, ByVal avecRecherche As Boolean)
    Dim W As Single, H As Single, hEntete As Single, hPied As Single
    Dim yCorps As Single, hCorps As Single, wNav As Single, hTab As Single
    Dim x0 As Single, wZone As Single, xIn As Single, yIn As Single, wIn As Single
    Dim wQ As Single, y2 As Single, yBas As Single, xD As Single, wD As Single, hD As Single
    Dim pw As Single, ph As Single, iw As Single, y As Single, wR As Single, wTiers As Single
    Dim xEnr As Single, wRes As Single
    Dim c As Object, lst As Object, ts As Object, mp As Object, pg As Object, fr As Object
    Dim chap As Variant, etat As Variant, pct As Variant, i As Long
    Dim ok As String, enCours As String, partiel As String, vide As String, fille As String

    ok = ChrW(&H2713): enCours = ChrW(&H25CF): partiel = ChrW(&H25D0)
    vide = ChrW(&H25CB): fille = ChrW(&H21B3)

    etape = "dimensions"
    DimensionnerForm f, 0.95, 760, 420, 1500, 900
    f.BackColor = cFond
    W = f.InsideWidth: H = f.InsideHeight
    hEntete = 56: hPied = 42
    yCorps = hEntete + M
    hCorps = H - yCorps - hPied
    wNav = Borne(W * 0.2, 175, 250)
    x0 = M + wNav
    wZone = W - M - x0

    ' --- Bandeau -------------------------------------------------------------
    etape = "bandeau"
    Set c = Lbl(f, "lblBandeau", "", 0, 0, W, hEntete)
    c.BackStyle = 1: c.BackColor = cBandeau
    Lbl f, "lblTitre", "Élaboration du DCO", M + 4, 6, 300, LigneH(4), True, TAILLE_BASE + 4, cBlanc
    wR = Borne(W * 0.3, 240, 320)
    Lbl f, "lblContexte", "Offre 2026-123 · Logiciels de gestion   |   DAI · Univers Logiciels · Lot : tous", _
        M + 4, 33, W - M - wR - 2 * M - 120 - (M + 4), LigneH(-1), False, TAILLE_BASE - 1, RGB(200, 216, 234)
    BtnPlat f, "btnPageGarde", "Page de garde", W - M - wR - M - 120, 30, 120, 20, 2, ICO_CRAYON
    Set c = Txt(f, "txtRecherche", IIf(avecRecherche, "reconduction", "Rechercher une question (El84, reconduction…)"), _
                W - M - wR, 7, wR - 30, 20)
    c.ForeColor = IIf(avecRecherche, cTexte, cGris)
    BtnPlat f, "btnRechercher", "", W - M - 28, 7, 28, 20, 1, ICO_RECHERCHE
    Set c = Lbl(f, "lblBarreFond", "", W - M - wR, 37, wR - 130, 8)
    c.BackStyle = 1: c.BackColor = RGB(12, 36, 64)
    Set c = Lbl(f, "lblBarre", "", W - M - wR, 37, (wR - 130) * 0.62, 8)
    c.BackStyle = 1: c.BackColor = RGB(76, 175, 80)
    Set c = Lbl(f, "lblAvancement", "62 % · 231 / 372", W - M - 122, 32, 122, LigneH(-1), True, TAILLE_BASE - 1, cBlanc)
    c.TextAlign = 3

    ' --- Niveau 1 : CHAPITRES = onglets verticaux (Labels) -------------------
    etape = "chapitres"
    Lbl f, "lblChapTitre", "CHAPITRES", M + 4, yCorps, wNav - 8, LigneH(-2), True, TAILLE_BASE - 2, cGris
    chap = Array("01 Présentation générale", "02 Environnement de l'offre", "03 Montage contractuel", _
                 "04 Devis et commande", "05 Garantie, SAV, maint.", "06 Exécution financière", _
                 "07 Performance fournisseur", "08 Données personnelles", "09 Propriété intellect.", _
                 "10 Prix", "11 Modalités d'exécution", "12 Analyse des offres", "13 Commercialisation")
    etat = Array(2, 2, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0)       ' 2 terminé, 1 en cours, 0 à faire
    pct = Array(100, 100, 45, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    hTab = Borne((hCorps - 18 - 34) / 13, LigneH() + 4, LigneH() + 14)
    For i = 0 To 12
        y = yCorps + 18 + i * hTab
        Set c = Lbl(f, "lblChapFond" & i, "", M, y, wNav, hTab - 1)
        c.BackStyle = 1
        c.BackColor = IIf(i = 2, cBlanc, cOngletFond)
        Set c = Lbl(f, "lblChap" & i, Choose(etat(i) + 1, vide, partiel, ok) & "  " & chap(i), _
                    M + 10, y + (hTab - 1 - LigneH()) / 2, wNav - 48, LigneH(), (i = 2))
        c.Font.Name = POLICE_SYMB: c.WordWrap = False
        c.ForeColor = Choose(etat(i) + 1, cGris, cTexte, cVert)
        Set c = Lbl(f, "lblChapPct" & i, pct(i) & " %", M + wNav - 40, y + (hTab - 1 - LigneH(-1)) / 2, _
                    36, LigneH(-1), False, TAILLE_BASE - 1, cGris)
        c.TextAlign = 3
    Next i
    Set c = Lbl(f, "lblChapSel", "", M, yCorps + 18 + 2 * hTab, 4, hTab - 1)     ' repère du chapitre ouvert
    c.BackStyle = 1: c.BackColor = cAction
    Set c = Lbl(f, "lblLegende", ok & " terminé   " & partiel & " en cours   " & vide & " à faire   " & _
                "! à justifier   " & fille & " question fille", _
                M + 2, yCorps + hCorps - 30, wNav - 4, 30, False, TAILLE_BASE - 2, cGris)
    c.Font.Name = POLICE_SYMB

    ' --- Niveau 2 : THÈMES = TabStrip ----------------------------------------
    etape = "thèmes"
    Set ts = Ctl(f, "Forms.TabStrip.1", "tsThemes", x0, yCorps, wZone, hCorps)
    ts.Font.Name = POLICE_SYMB
    ts.TabFixedHeight = 22
    ts.Tabs.Clear
    ts.Tabs.Add "th1", ok & " 03.1 Passation"
    ts.Tabs.Add "th2", partiel & " 03.2 Intervention UGAP"
    ts.Tabs.Add "th3", partiel & " 03.3 Durée"
    ts.Tabs.Add "th4", vide & " 03.4 Clauses modif."
    ts.Tabs.Add "th5", vide & " 03.5 Fin de contrat"
    ts.Value = 1
    Set c = Lbl(f, "lblZoneBlanche", "", x0 + 2, yCorps + 26, wZone - 4, hCorps - 28)
    c.BackStyle = 1: c.BackColor = cBlanc
    xIn = x0 + M + 2: yIn = yCorps + 22 + M: wIn = wZone - 2 * M - 4
    yBas = yCorps + hCorps - M

    ' --- Niveau 3 : SOUS-THÈMES = TabStrip style boutons ---------------------
    ' (affiché seulement si le thème a plusieurs sous-thèmes ; un sous-thème
    '  masqué par une question mère n'apparaît pas)
    etape = "sous-thèmes"
    Set ts = Ctl(f, "Forms.TabStrip.1", "tsSousThemes", xIn, yIn, wIn, 26)
    ts.Font.Name = POLICE_SYMB
    ts.Style = 1                          ' fmTabStyleButtons
    ts.TabFixedHeight = 22
    ts.Tabs.Clear
    ts.Tabs.Add "st1", ok & " Modalités d'intervention"
    ts.Tabs.Add "st2", partiel & " 03.2.2 Offre non exécutée"
    ts.Value = 1
    Lbl f, "lblMasque", "1 sous-thème masqué (« Offre exécutée ») par la réponse à El77.", _
        xIn + 2, yIn + 31, wIn - 180, LigneH(-2), False, TAILLE_BASE - 2, cGris
    Set c = Lbl(f, "lblAfficherMasques", "Afficher les éléments masqués", xIn + wIn - 176, yIn + 31, 176, _
                LigneH(-2), False, TAILLE_BASE - 2, cAction)
    c.Font.Underline = True: c.TextAlign = 3

    ' --- Questions du sous-thème ---------------------------------------------
    etape = "questions"
    y2 = yIn + 54
    wQ = Borne(wIn * 0.4, 220, 380)
    SectionTitre f, "secQuestions", "QUESTIONS", xIn, y2, wQ
    Lbl f, "lblFiltre", "Afficher", xIn, y2 + 27, 52, LigneH()
    Combo f, "cboFiltre", xIn + 54, y2 + 24, wQ - 54, 22, _
          Array("Toutes les questions", "Non répondues", "À justifier", "À revoir", "Écarts au standard")
    Lbl f, "lblCompteur", "5 affichées · 1 masquée", xIn, y2 + 51, wQ, LigneH(-2), False, TAILLE_BASE - 2, cGris
    Set lst = Liste(f, "lstQuestions", xIn, y2 + 68, wQ, yBas - 36 - (y2 + 68), "16;44;" & CLng(wQ - 16 - 44 - 6))
    AjouterLigne lst, ok, "El78f", "Dérogation au principe de non exclusivité ?"
    AjouterLigne lst, enCours, "El78b", "Adhésion via formulaire sur ugap.fr ?"
    AjouterLigne lst, vide, "El78c", "Qui effectue les marchés subséquents ?"
    AjouterLigne lst, vide, "El78d", "Volume estimé des marchés par an ?"
    AjouterLigne lst, "!", "El78e", "Modalités de passation des marchés subséquents"
    lst.ListIndex = 1
    BtnPlat f, "btnSousThemeStd", "Appliquer le standard aux non répondues", xIn, yBas - 28, wQ, 28, 0, ICO_COCHE

    ' --- Détail de la question : MultiPage ------------------------------------
    etape = "détail"
    xD = xIn + wQ + 2 * M: wD = wIn - wQ - 2 * M: hD = yBas - y2
    Set mp = Ctl(f, "Forms.MultiPage.1", "mpDetail", xD, y2, wD, hD)
    mp.TabFixedHeight = 22
    Do While mp.Pages.Count < 4
        mp.Pages.Add
    Loop
    mp.Pages(0).Caption = "Réponse"
    mp.Pages(1).Caption = "Standards"
    mp.Pages(2).Caption = "Commentaires (1)"
    mp.Pages(3).Caption = "Historique"
    mp.Value = 0
    pw = wD - 6: ph = hD - 30

    ' Page 1 : Réponse
    etape = "page Réponse"
    Set pg = mp.Pages(0)
    iw = Borne(pw - 2 * M, 200, 640)      ' largeur de lecture limitée sur grand écran
    If ph < 380 Then                      ' petit écran : la page défile au lieu d'être tronquée
        pg.ScrollBars = 2: pg.ScrollHeight = 380
        ph = 380: iw = Borne(pw - 2 * M - 14, 200, 640)
    End If
    Set c = Lbl(pg, "lblFondPage", "", 0, 0, pw, ph)
    c.BackStyle = 1: c.BackColor = cBlanc
    Lbl pg, "lblTypeQ", "El78b  ·  Liste à choix unique  ·  affichée car El77 = « Offre non exécutée »", _
        M, 8, iw, LigneH(-2), False, TAILLE_BASE - 2, cGris
    Lbl pg, "lblLibelle", "L’adhésion ou la renonciation se fera-t-elle via un formulaire d’enquête sur ugap.fr ?", _
        M, 26, iw, 3 * LigneH(2), True, TAILLE_BASE + 2
    y = 26 + 3 * LigneH(2) + 10
    Set c = Lbl(pg, "lblStdFond", "", M, y, iw, 64)
    c.BackStyle = 1: c.BackColor = cStdFond
    Set c = Lbl(pg, "lblStdBarre", "", M, y, 4, 64)
    c.BackStyle = 1: c.BackColor = cStdBarre
    Lbl pg, "lblStdTitre", "STANDARD APPLICABLE  ·  Établissement (E)", M + 14, y + 7, iw - 172, LigneH(-2), _
        True, TAILLE_BASE - 2, cOrange
    Lbl pg, "lblStdValeur", "Non, absence de formulaire sur ugap.fr", M + 14, y + 24, iw - 172, 2 * LigneH(), True
    BtnPlat pg, "btnReprendreStd", "Reprendre le standard", M + iw - 152, y + 19, 144, 26, 0
    y = y + 64 + 18
    SectionTitre pg, "secReponse", "RÉPONSE", M, y, iw
    y = y + 26
    Opt pg, "optRep1", "Oui, adhésion ou renonciation sur formulaire ugap.fr", M, y, iw, 22, False, "grpReponse"
    Opt pg, "optRep2", "Non, absence de formulaire sur ugap.fr", M, y + 26, iw, 22, True, "grpReponse"
    y = y + 52 + 14
    SectionTitre pg, "secConformite", "CONFORMITÉ AU STANDARD", M, y, iw
    y = y + 26
    wTiers = iw / 3
    Opt pg, "optConforme", "Conforme", M, y, wTiers, 22, True, "grpConformite"
    Opt pg, "optSpecificite", "Spécificité (à justifier)", M + wTiers, y, wTiers + 24, 22, False, "grpConformite"
    Opt pg, "optNonConcerne", "Non concerné", M + 2 * wTiers + 24, y, wTiers - 24, 22, False, "grpConformite"
    y = y + 30
    Lbl pg, "lblNoteJustif", "La zone de justification apparaît quand « Spécificité » est cochée.", _
        M, y, iw, LigneH(-2), False, TAILLE_BASE - 2, cGris
    Lbl pg, "lblRaccourci", "Ctrl+Entrée", M, ph - 30, 100, LigneH(-2), False, TAILLE_BASE - 2, cGris
    BtnPlat pg, "btnValiderSuivante", "Valider et suivante", M + iw - 176, ph - 40, 176, 30, 1, ICO_SUIVANT

    ' Page 2 : Standards
    etape = "page Standards"
    Set pg = mp.Pages(1)
    iw = pw - 2 * M
    Set c = Lbl(pg, "lblFondPage2", "", 0, 0, pw, ph)
    c.BackStyle = 1: c.BackColor = cBlanc
    Lbl pg, "lblStdInfo", "Le niveau applicable dépend de la direction et de l'univers choisis en page de garde.", _
        M, 8, iw, 2 * LigneH(-1), False, TAILLE_BASE - 1, cGris
    Set lst = Liste(pg, "lstStandards", M, 40, iw, 6 * LigneH() + 8, "120;" & CLng(iw - 126))
    lst.Font.Name = POLICE
    AjouterLigne lst, "Établissement (E)", "Non, absence de formulaire sur ugap.fr"
    AjouterLigne lst, "Direction DAI", "(pas de standard direction)"
    AjouterLigne lst, "Univers Logiciels", "(pas de standard univers)"
    AjouterLigne lst, "Niveau applicable", "E  ->  Établissement"
    AjouterLigne lst, "Alimente le CCAP", "Non"
    lst.ListIndex = 0

    ' Page 3 : Commentaires
    etape = "page Commentaires"
    Set pg = mp.Pages(2)
    Set c = Lbl(pg, "lblFondPage3", "", 0, 0, pw, ph)
    c.BackStyle = 1: c.BackColor = cBlanc
    Set lst = Liste(pg, "lstCommentaires", M, 8, iw, Borne(ph - 130, 50, 400), "70;74;" & CLng(iw - 150))
    lst.Font.Name = POLICE
    AjouterLigne lst, "05/10 10:40", "J. Martin", "Vérifier avec la DAI si le formulaire est prévu en 2027."
    SectionTitre pg, "secNouveauCom", "NOUVEAU COMMENTAIRE", M, ph - 118, iw
    Set c = Txt(pg, "txtNouveauCom", "", M, ph - 92, iw, 50)
    c.MultiLine = True
    BtnPlat pg, "btnAjouterCom", "Ajouter", M + iw - 100, ph - 36, 100, 26, 1

    ' Page 4 : Historique
    etape = "page Historique"
    Set pg = mp.Pages(3)
    Set c = Lbl(pg, "lblFondPage4", "", 0, 0, pw, ph)
    c.BackStyle = 1: c.BackColor = cBlanc
    Set lst = Liste(pg, "lstHistorique", M, 8, iw, ph - 16, "70;74;" & CLng(iw - 150))
    lst.Font.Name = POLICE
    AjouterLigne lst, "05/10 10:42", "J. Martin", "Réponse : « Non, absence de formulaire sur ugap.fr »"
    AjouterLigne lst, "05/10 10:42", "J. Martin", "Standard repris"
    AjouterLigne lst, "02/10 16:05", "P. Durand", "Question affichée (El77 = « Offre non exécutée »)"

    ' --- Pied de page ---------------------------------------------------------
    etape = "pied de page"
    y = H - hPied + 8
    BtnPlat f, "btnPrecedente", "Précédente", M, y, 112, 28, 0, ICO_PRECEDENT
    BtnPlat f, "btnSuivante", "Suivante", M + 120, y, 104, 28, 0, ICO_SUIVANT
    BtnPlat f, "btnSuivanteNR", "Suivante non répondue", M + 232, y, 190, 28, 0, ICO_SUIVANT
    BtnPlat f, "btnExporter", "Exporter", W - M - 254, y, 120, 28, 0, ICO_EXPORT
    BtnPlat f, "btnFermer", "Fermer", W - M - 126, y, 126, 28, 0
    xEnr = W - M - 254 - 8 - 220
    If xEnr > M + 232 + 190 + 8 Then
        Set c = Lbl(f, "lblEnregistre", ChrW(&H2714) & " Enregistré automatiquement à 10:42", _
                    xEnr, y + 7, 220, LigneH(-2), False, TAILLE_BASE - 2, cVert)
    Else
        Set c = Lbl(f, "lblEnregistre", ChrW(&H2714) & " 10:42", W - M - 254 - 8 - 70, y + 7, 70, _
                    LigneH(-2), False, TAILLE_BASE - 2, cVert)
    End If
    c.Font.Name = POLICE_SYMB: c.TextAlign = 3

    ' --- Panneau de résultats de recherche (sous la barre de recherche) ------
    ' Un Frame, pour passer au-dessus des autres contrôles.
    If avecRecherche Then
        etape = "recherche"
        wRes = Borne(W * 0.45, 380, 540)
        Set fr = Ctl(f, "Forms.Frame.1", "fraResultats", W - M - wRes, 29, wRes, 5 * LigneH() + 44)
        fr.Caption = ""
        fr.BackColor = cBlanc
        fr.SpecialEffect = 0: fr.BorderStyle = 1: fr.BorderColor = RGB(150, 160, 175)
        Lbl fr, "lblResTitre", "5 résultats pour « reconduction »   ·   Entrée pour ouvrir, Échap pour fermer", _
            M, 6, wRes - 2 * M, LigneH(-2), False, TAILLE_BASE - 2, cGris
        Set lst = Liste(fr, "lstResultats", M, 24, wRes - 2 * M - 2, 5 * LigneH() + 10, _
                        "48;110;" & CLng(wRes - 2 * M - 2 - 48 - 110 - 6))
        lst.Font.Name = POLICE
        lst.BorderStyle = 0
        AjouterLigne lst, "El84", "03 › 03.3 Durée", "Le principe de reconduction s'applique-t-il ?"
        AjouterLigne lst, "El85b", "03 › 03.3 Durée", "Si oui, modalités de reconduction"
        AjouterLigne lst, "El84b", "03 › 03.3 Durée", "Si oui, nombre de reconductions possibles"
        AjouterLigne lst, "El84c", "03 › 03.3 Durée", "Si oui, durée de chaque reconduction"
        AjouterLigne lst, "El84e", "03 › 03.3 Durée", "(masquée) Modalités particulières par lot"
        lst.ListIndex = 0
    End If
End Sub

' -----------------------------------------------------------------------------
' Page de garde : boîte de dialogue séparée (création du dossier / modification)
' -----------------------------------------------------------------------------
Private Sub ConstruirePageGarde(f As Object)
    Dim W As Single, H As Single, c As Object, lst As Object
    Dim y As Single, xc As Single, wChamp As Single, wSec As Single

    etape = "page de garde"
    f.Width = 520: f.Height = 550
    CentrerSurExcel f
    f.BackColor = cBlanc
    W = f.InsideWidth: H = f.InsideHeight
    xc = 130: wSec = W - 2 * M - 8: wChamp = wSec - xc

    Set c = Lbl(f, "lblBandeau", "", 0, 0, W, 44)
    c.BackStyle = 1: c.BackColor = cBandeau
    Lbl f, "lblTitre", "Nouveau dossier d'élaboration", M + 4, 10, 360, LigneH(4), True, TAILLE_BASE + 4, cBlanc

    y = 58
    SectionTitre f, "secOffre", "OFFRE", M + 4, y, wSec
    y = y + 26
    Lbl f, "lblNum", "N° d'offre", M + 4, y + 3, xc, LigneH()
    Txt f, "txtNum", "2026-123", M + 4 + xc, y, 120, 22
    y = y + 30
    Lbl f, "lblIntitule", "Intitulé", M + 4, y + 3, xc, LigneH()
    Txt f, "txtIntitule", "Renouvellement des logiciels de gestion", M + 4 + xc, y, wChamp, 22
    y = y + 30
    Lbl f, "lblAcheteur", "Acheteur", M + 4, y + 3, xc, LigneH()
    Set c = Txt(f, "txtAcheteur", "Jeanne Martin (identifiant Windows)", M + 4 + xc, y, wChamp, 22)
    c.Locked = True: c.BackColor = cFond: c.ForeColor = cGris

    y = y + 42
    SectionTitre f, "secStandards", "STANDARDS APPLICABLES", M + 4, y, wSec
    y = y + 26
    Lbl f, "lblDirAchat", "Direction d'achat", M + 4, y + 3, xc, LigneH()
    Combo f, "cboDirAchat", M + 4 + xc, y, 180, 22, Array("DAG", "DAI", "DAV", "DS"), 1
    y = y + 30
    Lbl f, "lblDirection", "Direction", M + 4, y + 3, xc, LigneH()
    Combo f, "cboDirection", M + 4 + xc, y, 180, 22, _
          Array("DAV", "DAI", "DS", "DAG Mobilier scolaire et Equipement général", "DAG SOFI", "Non standardisé"), 1
    y = y + 30
    Lbl f, "lblUnivers", "Univers", M + 4, y + 3, xc, LigneH()
    Combo f, "cboUnivers", M + 4 + xc, y, 180, 22, _
          Array("PII", "Materiel_et_Prestations", "Copieurs", "Telecom", "Bureau WEB", "Logiciels"), 5
    y = y + 28
    Lbl f, "lblInfoStd", "Ces choix déterminent les standards affichés pour chaque question. " & _
        "Les modifier plus tard déclenche une analyse d'impact sur les réponses déjà saisies.", _
        M + 4, y, wSec, 2 * LigneH(-2) + 2, False, TAILLE_BASE - 2, cGris

    y = y + 40
    SectionTitre f, "secLots", "LOTS (MARCHÉS)", M + 4, y, wSec
    y = y + 26
    Set lst = Liste(f, "lstLots", M + 4, y, wSec - 112, 84, "48;" & CLng(wSec - 112 - 54))
    lst.Font.Name = POLICE
    AjouterLigne lst, "Lot 1", "Logiciels de gestion financière"
    AjouterLigne lst, "Lot 2", "Logiciels de gestion des ressources humaines"
    AjouterLigne lst, "Lot 3", "Maintenance et support"
    BtnPlat f, "btnAjouterLot", "Ajouter", M + 4 + wSec - 104, y, 104, 24, 0
    BtnPlat f, "btnModifierLot", "Modifier", M + 4 + wSec - 104, y + 30, 104, 24, 0
    BtnPlat f, "btnSupprimerLot", "Supprimer", M + 4 + wSec - 104, y + 60, 104, 24, 0

    BtnPlat f, "btnAnnuler", "Annuler", W - M - 4 - 260, H - 40, 110, 30, 0
    BtnPlat f, "btnCreer", "Créer le dossier", W - M - 4 - 142, H - 40, 142, 30, 1
End Sub

' -----------------------------------------------------------------------------
' Création du UserForm temporaire
' -----------------------------------------------------------------------------
' Le UserForm vide est créé une seule fois puis réutilisé : VBA ne sait pas
' recréer un formulaire supprimé dans la même session (erreur « Objet
' spécifié introuvable »). Les contrôles sont ajoutés à l'exécution et
' disparaissent à la fermeture ; le formulaire, lui, reste vide.
Private Function CreerForm(titre As String) As Object
    Dim comp As Object, frm As Object
    etape = "création du formulaire"
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(NOM_FORM)
    On Error GoTo erreur
    If comp Is Nothing Then
        Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)   ' vbext_ct_MSForm
        comp.Name = NOM_FORM
        On Error Resume Next
        comp.Properties("StartUpPosition") = 0              ' position manuelle (centrée sur Excel)
        On Error GoTo erreur
    End If
    Set frm = VBA.UserForms.Add(NOM_FORM)
    frm.Caption = titre
    Set CreerForm = frm
    Exit Function
erreur:
    MsgBox "Impossible de créer la maquette : " & Err.Description & vbCrLf & vbCrLf & _
           "Fermez le classeur sans l'enregistrer, rouvrez-le et relancez la macro.", vbExclamation
    Set CreerForm = Nothing
End Function

' Retire le formulaire vide de la maquette du classeur (à lancer avant d'enregistrer).
Public Sub MaquetteNettoyer()
    Dim comp As Object
    On Error Resume Next
    Set comp = ThisWorkbook.VBProject.VBComponents(NOM_FORM)
    If comp Is Nothing Then
        MsgBox "Rien à nettoyer.", vbInformation
    Else
        ThisWorkbook.VBProject.VBComponents.Remove comp
        MsgBox "Formulaire de maquette supprimé.", vbInformation
    End If
End Sub

' Taille proportionnelle à la fenêtre Excel, bornée, et centrée dessus
' (donc sur l'écran où se trouve Excel en cas d'écran supplémentaire).
Private Sub DimensionnerForm(f As Object, ByVal ratio As Single, ByVal wMin As Single, ByVal hMin As Single, _
                             ByVal wMax As Single, ByVal hMax As Single)
    f.Width = Borne(Application.UsableWidth * ratio, wMin, wMax)
    f.Height = Borne(Application.UsableHeight * ratio, hMin, hMax)
    CentrerSurExcel f
End Sub

Private Sub CentrerSurExcel(f As Object)
    f.Left = Application.Left + (Application.Width - f.Width) / 2
    f.Top = Application.Top + (Application.Height - f.Height) / 2
End Sub

Private Function Borne(ByVal v As Single, ByVal mini As Single, ByVal maxi As Single) As Single
    If v < mini Then
        Borne = mini
    ElseIf v > maxi Then
        Borne = maxi
    Else
        Borne = v
    End If
End Function

' Hauteur d'une ligne de texte pour la taille de base + delta
Private Function LigneH(Optional ByVal delta As Single = 0) As Single
    LigneH = (TAILLE_BASE + delta) * 1.45
End Function

' -----------------------------------------------------------------------------
' Fabrique de contrôles MSForms
' -----------------------------------------------------------------------------
Private Function Ctl(parent As Object, progId As String, nom As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add(progId, nom, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Name = POLICE
    c.Font.Size = TAILLE_BASE
    Set Ctl = c
End Function

' Bordure fine et plate (style moderne) pour les zones de saisie et les listes
Private Sub BordurePlate(c As Object)
    On Error Resume Next
    c.SpecialEffect = 0
    c.BorderStyle = 1
    c.BorderColor = RGB(170, 178, 189)
End Sub

Private Function Lbl(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     Optional ByVal gras As Boolean = False, Optional ByVal taille As Single = 0, _
                     Optional ByVal couleur As Long = -1) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.Label.1", nom, l, t, w, h)
    c.Caption = texte
    c.BackStyle = 0
    c.WordWrap = True
    c.Font.Bold = gras
    c.Font.Size = IIf(taille = 0, TAILLE_BASE, taille)
    c.ForeColor = IIf(couleur = -1, cTexte, couleur)
    Set Lbl = c
End Function

' Bouton « plat » : Label de fond + libellé (+ icône Segoe MDL2 Assets).
' genre : 0 secondaire (blanc, bordure), 1 principal (couleur d'action), 2 sur le bandeau.
' Dans l'application, le survol changera la couleur (événement MouseMove).
Private Function BtnPlat(parent As Object, nom As String, texte As String, _
                         ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                         Optional ByVal genre As Long = 0, Optional ByVal icone As Long = 0) As Object
    Dim fond As Object, c As Object, xTexte As Single, coul As Long
    Set fond = Lbl(parent, nom, "", l, t, w, h)
    fond.BackStyle = 1
    fond.BorderStyle = 1
    Select Case genre
        Case 1
            fond.BackColor = cAction: fond.BorderColor = cAction: coul = cBlanc
        Case 2
            fond.BackColor = cBandeau: fond.BorderColor = RGB(120, 150, 190): coul = cBlanc
        Case Else
            fond.BackColor = cBlanc: fond.BorderColor = RGB(170, 178, 189): coul = cTexte
    End Select
    xTexte = l
    If icone <> 0 Then
        Set c = Lbl(parent, nom & "_ico", ChrW(icone), l + 8, t + (h - LigneH()) / 2 + 2, 16, LigneH(), False, TAILLE_BASE, coul)
        c.Font.Name = POLICE_ICONES
        If texte = "" Then
            c.Left = l + (w - 16) / 2
            c.TextAlign = 2
        Else
            xTexte = l + 20
        End If
    End If
    If texte <> "" Then
        Set c = Lbl(parent, nom & "_txt", texte, xTexte, t + (h - LigneH()) / 2, w - (xTexte - l), LigneH(), _
                    (genre = 1), TAILLE_BASE, coul)
        c.TextAlign = 2
        c.WordWrap = False
    End If
    Set BtnPlat = fond
End Function

' Titre de section en petites capitales + trait fin (remplace les cadres)
Private Sub SectionTitre(parent As Object, nom As String, texte As String, _
                         ByVal l As Single, ByVal t As Single, ByVal w As Single)
    Dim c As Object
    Lbl parent, nom, texte, l, t, w, LigneH(-2), True, TAILLE_BASE - 2, cBandeau
    Set c = Lbl(parent, nom & "_trait", "", l, t + LigneH(-2) + 3, w, 1)
    c.BackStyle = 1: c.BackColor = cLigne
End Sub

Private Function Txt(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.TextBox.1", nom, l, t, w, h)
    c.Text = texte
    BordurePlate c
    Set Txt = c
End Function

Private Function Liste(parent As Object, nom As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                       largeursColonnes As String) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.ListBox.1", nom, l, t, w, h)
    c.IntegralHeight = False
    c.Font.Name = POLICE_SYMB
    BordurePlate c
    If largeursColonnes <> "" Then
        c.ColumnCount = UBound(Split(largeursColonnes, ";")) + 1
        c.ColumnWidths = largeursColonnes
    End If
    Set Liste = c
End Function

Private Function Combo(parent As Object, nom As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                       elements As Variant, Optional ByVal selection As Long = 0) As Object
    Dim c As Object, e As Variant
    Set c = Ctl(parent, "Forms.ComboBox.1", nom, l, t, w, h)
    c.Style = 2                 ' fmStyleDropDownList
    BordurePlate c
    For Each e In elements
        c.AddItem e
    Next e
    c.ListIndex = selection
    Set Combo = c
End Function

Private Function Opt(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     ByVal coche As Boolean, Optional ByVal groupe As String = "") As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.OptionButton.1", nom, l, t, w, h)
    c.Caption = texte
    c.BackStyle = 0
    c.GroupName = IIf(groupe = "", parent.Name, groupe)
    c.Value = coche
    Set Opt = c
End Function

Private Sub AjouterLigne(lst As Object, ParamArray valeurs() As Variant)
    Dim i As Long
    lst.AddItem valeurs(0)
    For i = 1 To UBound(valeurs)
        lst.List(lst.ListCount - 1, i) = valeurs(i)
    Next i
End Sub

Option Explicit

' =============================================================================
' Maquette visuelle des écrans (sans interaction, données d'exemple réelles).
' Uniquement des contrôles MSForms standards : Label, Frame, ListBox, ComboBox,
' TextBox, OptionButton, CommandButton, TabStrip, MultiPage. Aucun ActiveX.
'
'   Alt+F8 > MaquetteEcranSaisie : fenêtre principale (chapitres > thèmes > sous-thèmes > question)
'   Alt+F8 > MaquettePageGarde   : page de garde (offre, standards, lots)
'
' Un UserForm vide « frmMaquetteTmp » est ajouté au classeur au premier lancement
' et réutilisé ensuite. Avant d'enregistrer le classeur : Alt+F8 > MaquetteNettoyer.
' Nécessite "Accès approuvé au modèle d'objet du projet VBA" (déjà activé pour VbaSync).
'
' La fenêtre principale se dimensionne sur la fenêtre Excel (95 %) et ses trois
' zones se répartissent la largeur disponible : redimensionnez Excel puis
' relancez la macro pour voir l'adaptation (portable / grand écran).
' =============================================================================

Private Const NOM_FORM As String = "frmMaquetteTmp"
Private Const POLICE As String = "Segoe UI"
Private Const POLICE_SYMB As String = "Segoe UI Symbol"   ' contient les pictogrammes de statut
Private Const M As Single = 6                             ' marge

' -----------------------------------------------------------------------------
' Points d'entrée
' -----------------------------------------------------------------------------
Public Sub MaquetteEcranSaisie()
    Dim f As Object
    Set f = CreerForm("SGO - Élaboration du DCO   (maquette visuelle - fermer avec la croix)")
    If f Is Nothing Then Exit Sub
    On Error GoTo fin
    ConstruireEcranSaisie f
    f.Show
fin:
    If Err.Number <> 0 Then MsgBox "Erreur maquette : " & Err.Description, vbExclamation
    Set f = Nothing
End Sub

Public Sub MaquettePageGarde()
    Dim f As Object
    Set f = CreerForm("SGO - Page de garde du dossier   (maquette visuelle)")
    If f Is Nothing Then Exit Sub
    On Error GoTo fin
    ConstruirePageGarde f
    f.Show
fin:
    If Err.Number <> 0 Then MsgBox "Erreur maquette : " & Err.Description, vbExclamation
    Set f = Nothing
End Sub

' -----------------------------------------------------------------------------
' Écran principal : une seule fenêtre, trois zones
' -----------------------------------------------------------------------------
Private Sub ConstruireEcranSaisie(f As Object)
    Dim W As Single, H As Single, hEntete As Single, hPied As Single
    Dim yCorps As Single, hCorps As Single, wNav As Single, hTab As Single
    Dim x0 As Single, wC As Single, xIn As Single, yIn As Single, wIn As Single
    Dim wQ As Single, y2 As Single, yBas As Single, xD As Single, wD As Single, hD As Single
    Dim pw As Single, ph As Single, iw As Single, y As Single, wOpt As Single, wR As Single
    Dim c As Object, lst As Object, ts As Object, mp As Object, pg As Object, sub_ As Object
    Dim chap As Variant, etat As Variant, pct As Variant, i As Long
    Dim ok As String, enCours As String, partiel As String, vide As String, fille As String

    ok = ChrW(&H2713): enCours = ChrW(&H25CF): partiel = ChrW(&H25D0)
    vide = ChrW(&H25CB): fille = ChrW(&H21B3)

    DimensionnerForm f, 0.95, 760, 420, 1500, 900
    W = f.InsideWidth: H = f.InsideHeight
    hEntete = 52: hPied = 34
    yCorps = hEntete + M
    hCorps = H - yCorps - hPied
    wNav = Borne(W * 0.19, 165, 230)
    x0 = M + wNav                         ' les onglets de chapitres touchent la zone de contenu
    wC = W - M - x0

    ' --- Bandeau -------------------------------------------------------------
    Set c = Lbl(f, "lblBandeau", "", 0, 0, W, hEntete)
    c.BackStyle = 1: c.BackColor = RGB(31, 78, 121)
    Lbl f, "lblTitre", "Élaboration du DCO", M + 4, 5, 260, 20, True, 13, vbWhite
    wR = Borne(W * 0.28, 220, 300)        ' bloc recherche + avancement, à droite
    Lbl f, "lblContexte", "Offre 2026-123 · Logiciels de gestion   |   DAI · Univers Logiciels · Lot : tous", _
        M + 4, 29, W - M - wR - 66 - 6 - (M + 4), 16, False, 9, RGB(215, 228, 240)
    Btn f, "btnPageGarde", "Modifier…", W - M - wR - 66, 26, 58, 20
    Set c = Txt(f, "txtRecherche", "Rechercher (ex : El84, reconduction…)", W - M - wR, 5, wR - 80, 19)
    c.ForeColor = RGB(120, 120, 120)
    Btn f, "btnRechercher", "Chercher", W - M - 76, 5, 76, 19
    Set c = Lbl(f, "lblBarreFond", "", W - M - wR, 30, wR - 140, 14)
    c.BackStyle = 1: c.BackColor = RGB(20, 55, 90)
    Set c = Lbl(f, "lblBarre", "", W - M - wR, 30, (wR - 140) * 0.62, 14)
    c.BackStyle = 1: c.BackColor = RGB(84, 168, 84)
    Lbl f, "lblAvancement", "62 % · 231 / 372", W - M - 134, 30, 134, 14, True, 9, vbWhite

    ' --- Niveau 1 : CHAPITRES = onglets verticaux (Labels) -------------------
    ' Des Labels plutôt qu'un TabStrip vertical : texte horizontal lisible,
    ' couleur selon l'avancement, pourcentage, largeur réglable.
    Lbl f, "lblChapTitre", "CHAPITRES", M + 4, yCorps + 2, wNav - 8, 13, True, 8, RGB(90, 90, 90)
    chap = Array("01 Présentation générale", "02 Environnement de l'offre", "03 Montage contractuel", _
                 "04 Devis et commande", "05 Garantie, SAV, maint.", "06 Exécution financière", _
                 "07 Performance fournisseur", "08 Données personnelles", "09 Propriété intellect.", _
                 "10 Prix", "11 Modalités d'exécution", "12 Analyse des offres", "13 Commercialisation")
    etat = Array(2, 2, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0)       ' 2 terminé, 1 en cours, 0 à faire
    pct = Array(100, 100, 45, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    hTab = Borne((hCorps - 18 - 46) / 13, 17, 26)
    For i = 0 To 12
        y = yCorps + 18 + i * hTab
        Set c = Lbl(f, "lblChapFond" & i, "", M, y, wNav, hTab - 1)
        c.BackStyle = 1
        c.BackColor = IIf(i = 2, RGB(250, 250, 250), RGB(226, 232, 239))
        Set c = Lbl(f, "lblChap" & i, Choose(etat(i) + 1, vide, partiel, ok) & "  " & chap(i), _
                    M + 8, y + (hTab - 14) / 2, wNav - 44, 14, (i = 2), 9)
        c.Font.Name = POLICE_SYMB: c.WordWrap = False
        If etat(i) = 2 Then c.ForeColor = RGB(40, 120, 40)
        If etat(i) = 0 Then c.ForeColor = RGB(95, 95, 95)
        Set c = Lbl(f, "lblChapPct" & i, pct(i) & " %", M + wNav - 38, y + (hTab - 13) / 2, 34, 13, False, 8, RGB(110, 110, 110))
        c.TextAlign = 3
    Next i
    Set c = Lbl(f, "lblChapSel", "", M, yCorps + 18 + 2 * hTab, 3, hTab - 1)   ' repère du chapitre ouvert
    c.BackStyle = 1: c.BackColor = RGB(31, 78, 121)
    Set c = Lbl(f, "lblLegende", ok & " terminé / répondue   " & partiel & " en cours   " & vide & " à faire   " & _
                "! à justifier   " & enCours & " question ouverte   " & fille & " question fille", _
                M + 2, yCorps + hCorps - 44, wNav - 6, 42, False, 8, RGB(90, 90, 90))
    c.Font.Name = POLICE_SYMB

    ' --- Niveau 2 : THÈMES = TabStrip (onglets du haut) ----------------------
    Set ts = Ctl(f, "Forms.TabStrip.1", "tsThemes", x0, yCorps, wC, hCorps)
    ts.Font.Name = POLICE_SYMB
    ts.Tabs.Clear
    ts.Tabs.Add "th1", ok & " 03.1 Passation"
    ts.Tabs.Add "th2", partiel & " 03.2 Intervention UGAP"
    ts.Tabs.Add "th3", partiel & " 03.3 Durée"
    ts.Tabs.Add "th4", vide & " 03.4 Clauses modif."
    ts.Tabs.Add "th5", vide & " 03.5 Fin de contrat"
    ts.Value = 1
    xIn = x0 + M: yIn = yCorps + 26: wIn = wC - 2 * M
    yBas = yCorps + hCorps - M

    ' --- Niveau 3 : SOUS-THÈMES = TabStrip style boutons ---------------------
    ' (affiché seulement si le thème a plusieurs sous-thèmes)
    Set ts = Ctl(f, "Forms.TabStrip.1", "tsSousThemes", xIn, yIn, wIn, 24)
    ts.Font.Name = POLICE_SYMB
    ts.Style = 1                          ' fmTabStyleButtons
    ts.Tabs.Clear
    ts.Tabs.Add "st1", ok & " Modalités d'intervention"
    ts.Tabs.Add "st2", partiel & " 03.2.2 Offre non exécutée"
    ts.Tabs.Add "st3", "03.2.3 Offre exécutée (masqué)"
    On Error Resume Next
    ts.Tabs(2).Enabled = False            ' sous-thème entier masqué par la question mère
    ts.Tabs(2).ControlTipText = "Masqué : la réponse à El77 est « Offre non exécutée »"
    On Error GoTo 0
    ts.Value = 1
    Lbl f, "lblMasque", "Sous-thème « Offre exécutée » masqué : la réponse à El77 est « Offre non exécutée ».", _
        xIn + 2, yIn + 27, wIn, 13, False, 8, RGB(150, 100, 20)

    ' --- Liste des questions du sous-thème -----------------------------------
    y2 = yIn + 44
    wQ = Borne(wIn * 0.4, 210, 360)
    Lbl f, "lblFiltre", "Afficher :", xIn, y2 + 3, 48, 16
    Combo f, "cboFiltre", xIn + 48, y2, wQ - 48, 19, _
          Array("Toutes les questions", "Non répondues", "À justifier", "À revoir", "Écarts au standard")
    Lbl f, "lblCompteur", "5 affichées · 1 masquée (El78a)", xIn, y2 + 23, wQ, 13, False, 8, RGB(90, 90, 90)
    Set lst = Liste(f, "lstQuestions", xIn, y2 + 38, wQ, yBas - 28 - (y2 + 38), "14;38;" & CLng(wQ - 14 - 38 - 4))
    AjouterLigne lst, ok, "El78f", "Dérogation au principe de non exclusivité ?"
    AjouterLigne lst, enCours, "El78b", "Adhésion via formulaire sur ugap.fr ?"
    AjouterLigne lst, vide, "El78c", "Qui effectue les marchés subséquents ?"
    AjouterLigne lst, vide, "El78d", "Volume estimé des marchés par an ?"
    AjouterLigne lst, "!", "El78e", "Modalités de passation des marchés subséquents"
    lst.ListIndex = 1
    Btn f, "btnSousThemeStd", "Appliquer le standard aux questions non répondues", xIn, yBas - 24, wQ, 24

    ' --- Détail de la question : MultiPage ------------------------------------
    xD = xIn + wQ + M: wD = wIn - wQ - M: hD = yBas - y2
    Set mp = Ctl(f, "Forms.MultiPage.1", "mpDetail", xD, y2, wD, hD)
    Do While mp.Pages.Count < 4
        mp.Pages.Add
    Loop
    mp.Pages(0).Caption = "Réponse"
    mp.Pages(1).Caption = "Standards"
    mp.Pages(2).Caption = "Commentaires (1)"
    mp.Pages(3).Caption = "Historique"
    mp.Value = 0
    pw = wD - 8: ph = hD - 28

    ' Page 1 : Réponse
    Set pg = mp.Pages(0)
    iw = pw - 2 * M
    If ph < 340 Then                      ' petit écran : la page défile au lieu d'être tronquée
        pg.ScrollBars = 2: pg.ScrollHeight = 340
        ph = 340: iw = iw - 14
    End If
    Lbl pg, "lblTypeQ", "El78b  ·  Liste à choix unique  ·  affichée car El77 = « Offre non exécutée »", _
        M, 4, iw, 24, False, 8, RGB(110, 110, 110)
    Lbl pg, "lblLibelle", "L’adhésion ou la renonciation se fera-t-elle via un formulaire d’enquête sur ugap.fr ?", _
        M, 28, iw, 44, True, 11
    Set sub_ = Cadre(pg, "fraStandard", "Standard applicable", M, 76, iw, 58)
    sub_.BackColor = RGB(255, 248, 220)
    Lbl sub_, "lblNiveauStd", "Établissement (E)", M, 2, 120, 13, False, 8, RGB(110, 110, 110)
    Lbl sub_, "lblValeurStd", "Non, absence de formulaire sur ugap.fr", M, 15, iw - 144, 28, True, 10
    Btn sub_, "btnReprendreStd", "Reprendre le standard", sub_.InsideWidth - M - 124, 8, 124, 22
    y = 142
    Lbl pg, "lblReponse", "Réponse", M, y, 100, 14, True
    Set lst = Liste(pg, "lstReponse", M, y + 15, iw, 34, "")
    lst.ListStyle = 1                     ' boutons radio (cases à cocher si choix multiple)
    lst.Font.Name = POLICE
    lst.AddItem "Oui, adhésion ou renonciation sur formulaire ugap.fr"
    lst.AddItem "Non, absence de formulaire sur ugap.fr"
    lst.ListIndex = 1
    y = y + 56
    Lbl pg, "lblConformite", "Conformité au standard", M, y, 200, 14, True
    wOpt = iw / 3
    Opt pg, "optConforme", "Conforme", M, y + 15, wOpt, 18, True
    Opt pg, "optSpecificite", "Spécificité (à justifier)", M + wOpt, y + 15, wOpt + 20, 18, False
    Opt pg, "optNonConcerne", "Non concerné", M + 2 * wOpt + 20, y + 15, wOpt - 20, 18, False
    y = y + 40
    Lbl pg, "lblJustif", "Justification (seulement si spécificité)", M, y, 220, 14, True
    Set c = Txt(pg, "txtJustif", "", M, y + 15, iw, Borne(ph - (y + 15) - 34, 30, 200))
    c.MultiLine = True: c.Enabled = False: c.BackColor = RGB(238, 238, 238)
    Lbl pg, "lblRaccourci", "Ctrl+Entrée", M, ph - 22, 100, 14, False, 8, RGB(130, 130, 130)
    Set c = Btn(pg, "btnValiderSuivante", "Valider et suivante  " & ChrW(&H25B6), iw - 150 + M, ph - 26, 150, 22)
    c.Font.Name = POLICE_SYMB: c.Font.Bold = True

    ' Page 2 : Standards (comparaison des niveaux)
    Set pg = mp.Pages(1)
    iw = pw - 2 * M
    Lbl pg, "lblStdInfo", "Le niveau applicable dépend de la direction et de l'univers choisis en page de garde.", _
        M, 4, iw, 26, False, 8, RGB(110, 110, 110)
    Set lst = Liste(pg, "lstStandards", M, 32, iw, 100, "110;" & CLng(iw - 114))
    lst.Font.Name = POLICE
    AjouterLigne lst, "Établissement (E)", "Non, absence de formulaire sur ugap.fr"
    AjouterLigne lst, "Direction DAI", "(pas de standard direction)"
    AjouterLigne lst, "Univers Logiciels", "(pas de standard univers)"
    AjouterLigne lst, "Niveau applicable", "E  ->  Établissement"
    AjouterLigne lst, "Alimente le CCAP", "Non"
    lst.ListIndex = 0

    ' Page 3 : Commentaires
    Set pg = mp.Pages(2)
    Set lst = Liste(pg, "lstCommentaires", M, 4, iw, Borne(ph - 110, 50, 400), "62;70;" & CLng(iw - 136))
    lst.Font.Name = POLICE
    AjouterLigne lst, "05/10 10:40", "J. Martin", "Vérifier avec la DAI si le formulaire est prévu en 2027."
    Lbl pg, "lblNouveauCom", "Nouveau commentaire", M, ph - 100, 200, 14, True
    Set c = Txt(pg, "txtNouveauCom", "", M, ph - 84, iw, 52)
    c.MultiLine = True
    Btn pg, "btnAjouterCom", "Ajouter", iw - 80 + M, ph - 28, 80, 22

    ' Page 4 : Historique
    Set pg = mp.Pages(3)
    Set lst = Liste(pg, "lstHistorique", M, 4, iw, ph - 10, "62;70;" & CLng(iw - 136))
    lst.Font.Name = POLICE
    AjouterLigne lst, "05/10 10:42", "J. Martin", "Réponse : « Non, absence de formulaire sur ugap.fr »"
    AjouterLigne lst, "05/10 10:42", "J. Martin", "Standard repris"
    AjouterLigne lst, "02/10 16:05", "P. Durand", "Question affichée (El77 = « Offre non exécutée »)"

    ' --- Pied de page ---------------------------------------------------------
    y = H - hPied + 6
    Set c = Btn(f, "btnPrecedente", ChrW(&H25C0) & "  Précédente", M, y, 92, 22): c.Font.Name = POLICE_SYMB
    Set c = Btn(f, "btnSuivante", "Suivante  " & ChrW(&H25B6), M + 98, y, 92, 22): c.Font.Name = POLICE_SYMB
    Set c = Btn(f, "btnSuivanteNR", "Suivante non répondue  " & ChrW(&H25B6) & ChrW(&H25B6), M + 196, y, 150, 22)
    c.Font.Name = POLICE_SYMB
    Set c = Lbl(f, "lblEnregistre", ChrW(&H2714) & " Enregistré automatiquement à 10:42", _
                W - M - 176 - 210, y + 4, 205, 16, False, 8, RGB(60, 120, 60))
    c.Font.Name = POLICE_SYMB: c.TextAlign = 3
    Btn f, "btnExporter", "Exporter…", W - M - 176, y, 84, 22
    Btn f, "btnFermer", "Fermer", W - M - 86, y, 86, 22
End Sub

' -----------------------------------------------------------------------------
' Page de garde : boîte de dialogue séparée (création du dossier / modification)
' -----------------------------------------------------------------------------
Private Sub ConstruirePageGarde(f As Object)
    Dim W As Single, H As Single, fr As Object, c As Object, lst As Object
    Dim iw As Single, xc As Single, wc As Single

    f.Width = 500: f.Height = 445
    CentrerSurExcel f
    W = f.InsideWidth: H = f.InsideHeight
    xc = 110

    Set c = Lbl(f, "lblBandeau", "", 0, 0, W, 40)
    c.BackStyle = 1: c.BackColor = RGB(31, 78, 121)
    Lbl f, "lblTitre", "Nouveau dossier d'élaboration", M + 4, 9, 300, 22, True, 13, vbWhite

    ' Offre
    Set fr = Cadre(f, "fraOffre", "Offre", M, 48, W - 2 * M, 96)
    iw = fr.InsideWidth - 2 * M: wc = iw - xc
    Lbl fr, "lblNum", "N° d'offre", M, 6, xc, 16
    Txt fr, "txtNum", "2026-123", M + xc, 4, 120, 19
    Lbl fr, "lblIntitule", "Intitulé", M, 30, xc, 16
    Txt fr, "txtIntitule", "Renouvellement des logiciels de gestion", M + xc, 28, wc, 19
    Lbl fr, "lblAcheteur", "Acheteur", M, 54, xc, 16
    Set c = Txt(fr, "txtAcheteur", "Jeanne Martin (identifiant Windows)", M + xc, 52, wc, 19)
    c.Locked = True: c.BackColor = RGB(235, 235, 235)

    ' Standards
    Set fr = Cadre(f, "fraStandards", "Standards applicables", M, 150, W - 2 * M, 118)
    Lbl fr, "lblDirAchat", "Direction d'achat", M, 6, xc, 16
    Combo fr, "cboDirAchat", M + xc, 4, 160, 19, Array("DAG", "DAI", "DAV", "DS"), 1
    Lbl fr, "lblDirection", "Direction", M, 30, xc, 16
    Combo fr, "cboDirection", M + xc, 28, 160, 19, _
          Array("DAV", "DAI", "DS", "DAG Mobilier scolaire et Equipement général", "DAG SOFI", "Non standardisé"), 1
    Lbl fr, "lblUnivers", "Univers", M, 54, xc, 16
    Combo fr, "cboUnivers", M + xc, 52, 160, 19, _
          Array("PII", "Materiel_et_Prestations", "Copieurs", "Telecom", "Bureau WEB", "Logiciels"), 5
    Lbl fr, "lblInfoStd", "Ces choix déterminent les standards affichés pour chaque question. " & _
        "Les modifier plus tard déclenche une analyse d'impact sur les réponses déjà saisies.", _
        M, 76, fr.InsideWidth - 2 * M, 26, False, 8, RGB(110, 110, 110)

    ' Lots
    Set fr = Cadre(f, "fraLots", "Lots (marchés)", M, 274, W - 2 * M, 118)
    iw = fr.InsideWidth - 2 * M
    Set lst = Liste(fr, "lstLots", M, 4, iw - 92, fr.InsideHeight - 10, "40;" & CLng(iw - 92 - 44))
    lst.Font.Name = POLICE
    AjouterLigne lst, "Lot 1", "Logiciels de gestion financière"
    AjouterLigne lst, "Lot 2", "Logiciels de gestion des ressources humaines"
    AjouterLigne lst, "Lot 3", "Maintenance et support"
    Btn fr, "btnAjouterLot", "Ajouter…", iw - 82 + M, 4, 82, 22
    Btn fr, "btnModifierLot", "Modifier…", iw - 82 + M, 30, 82, 22
    Btn fr, "btnSupprimerLot", "Supprimer", iw - 82 + M, 56, 82, 22

    ' Boutons
    Btn f, "btnAnnuler", "Annuler", W - M - 220, H - 30, 100, 24
    Set c = Btn(f, "btnCreer", "Créer le dossier", W - M - 114, H - 30, 114, 24)
    c.Font.Bold = True
End Sub

' -----------------------------------------------------------------------------
' Création / suppression du UserForm temporaire
' -----------------------------------------------------------------------------
' Le UserForm vide est créé une seule fois puis réutilisé : VBA ne sait pas
' recréer un formulaire supprimé dans la même session (erreur « Objet
' spécifié introuvable »). Les contrôles sont ajoutés à l'exécution et
' disparaissent à la fermeture ; le formulaire, lui, reste vide.
Private Function CreerForm(titre As String) As Object
    Dim comp As Object, frm As Object
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
    frm.BackColor = RGB(243, 243, 243)
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

' -----------------------------------------------------------------------------
' Fabrique de contrôles MSForms
' -----------------------------------------------------------------------------
Private Function Ctl(parent As Object, progId As String, nom As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add(progId, nom, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Name = POLICE
    c.Font.Size = 9
    Set Ctl = c
End Function

Private Function Lbl(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                     Optional gras As Boolean = False, Optional taille As Single = 9, _
                     Optional couleur As Long = -1) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.Label.1", nom, l, t, w, h)
    c.Caption = texte
    c.BackStyle = 0
    c.WordWrap = True
    c.Font.Bold = gras
    c.Font.Size = taille
    If couleur <> -1 Then c.ForeColor = couleur
    Set Lbl = c
End Function

Private Function Btn(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.CommandButton.1", nom, l, t, w, h)
    c.Caption = texte
    Set Btn = c
End Function

Private Function Txt(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.TextBox.1", nom, l, t, w, h)
    c.Text = texte
    Set Txt = c
End Function

Private Function Cadre(parent As Object, nom As String, titre As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.Frame.1", nom, l, t, w, h)
    c.Caption = titre
    c.Font.Bold = True
    c.BackColor = RGB(250, 250, 250)
    Set Cadre = c
End Function

Private Function Liste(parent As Object, nom As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                       largeursColonnes As String) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.ListBox.1", nom, l, t, w, h)
    c.IntegralHeight = False
    c.Font.Name = POLICE_SYMB
    If largeursColonnes <> "" Then
        c.ColumnCount = UBound(Split(largeursColonnes, ";")) + 1
        c.ColumnWidths = largeursColonnes
    End If
    Set Liste = c
End Function

Private Function Combo(parent As Object, nom As String, _
                       ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                       elements As Variant, Optional selection As Long = 0) As Object
    Dim c As Object, e As Variant
    Set c = Ctl(parent, "Forms.ComboBox.1", nom, l, t, w, h)
    c.Style = 2                 ' fmStyleDropDownList
    For Each e In elements
        c.AddItem e
    Next e
    c.ListIndex = selection
    Set Combo = c
End Function

Private Function Opt(parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, coche As Boolean) As Object
    Dim c As Object
    Set c = Ctl(parent, "Forms.OptionButton.1", nom, l, t, w, h)
    c.Caption = texte
    c.GroupName = parent.Name
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

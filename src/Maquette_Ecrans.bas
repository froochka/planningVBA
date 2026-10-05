Option Explicit

' =============================================================================
' Maquette visuelle des écrans (sans interaction, données d'exemple réelles).
' Uniquement des contrôles MSForms standards : Label, Frame, ListBox, ComboBox,
' TextBox, OptionButton, CommandButton. Aucun ActiveX.
'
'   Alt+F8 > MaquetteEcranSaisie : fenêtre principale (navigation + questions + réponse)
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
    Dim yCorps As Single, hCorps As Single
    Dim wNav As Single, wDet As Single, wLst As Single, xLst As Single, xDet As Single
    Dim fr As Object, sub_ As Object, c As Object, lst As Object
    Dim iw As Single, ih As Single, y As Single, wc As Single, wR As Single
    Dim ok As String, enCours As String, partiel As String, vide As String, fille As String

    ok = ChrW(&H2713): enCours = ChrW(&H25CF): partiel = ChrW(&H25D0)
    vide = ChrW(&H25CB): fille = ChrW(&H21B3)

    DimensionnerForm f, 0.95, 760, 420, 1500, 900
    W = f.InsideWidth: H = f.InsideHeight
    hEntete = 52: hPied = 34
    yCorps = hEntete + M
    hCorps = H - yCorps - hPied
    wNav = Borne(W * 0.24, 175, 270)
    wDet = Borne(W * 0.4, 320, 540)
    xLst = M + wNav + M
    wLst = W - xLst - M - wDet - M
    xDet = xLst + wLst + M

    ' --- Bandeau -------------------------------------------------------------
    Set c = Lbl(f, "lblBandeau", "", 0, 0, W, hEntete)
    c.BackStyle = 1: c.BackColor = RGB(31, 78, 121)
    Lbl f, "lblTitre", "Élaboration du DCO", M + 4, 5, 260, 20, True, 13, vbWhite
    wR = Borne(W * 0.28, 220, 300)        ' bloc recherche + avancement, à droite
    Lbl f, "lblContexte", "Offre 2026-123 · Logiciels de gestion   |   DAI · Univers Logiciels · Lot : tous", _
        M + 4, 29, W - M - wR - 66 - 6 - (M + 4), 16, False, 9, RGB(215, 228, 240)
    Btn f, "btnPageGarde", "Modifier…", W - M - wR - 66, 26, 58, 20

    ' recherche + avancement global (à droite)
    Set c = Txt(f, "txtRecherche", "Rechercher (ex : El84, reconduction…)", W - M - wR, 5, wR - 80, 19)
    c.ForeColor = RGB(120, 120, 120)
    Btn f, "btnRechercher", "Chercher", W - M - 76, 5, 76, 19
    Set c = Lbl(f, "lblBarreFond", "", W - M - wR, 30, wR - 140, 14)
    c.BackStyle = 1: c.BackColor = RGB(20, 55, 90)
    Set c = Lbl(f, "lblBarre", "", W - M - wR, 30, (wR - 140) * 0.62, 14)
    c.BackStyle = 1: c.BackColor = RGB(84, 168, 84)
    Lbl f, "lblAvancement", "62 % · 231 / 372", W - M - 134, 30, 134, 14, True, 9, vbWhite

    ' --- Zone 1 : navigation chapitres / thèmes -----------------------------
    Set fr = Cadre(f, "fraNav", "Chapitres", M, yCorps, wNav, hCorps)
    iw = fr.InsideWidth - 2 * M: ih = fr.InsideHeight
    Set lst = Liste(fr, "lstNav", M, 4, iw, ih - 34, "14;" & CLng(iw - 14 - 38 - 8) & ";38")
    AjouterLigne lst, ok, "01 Présentation générale", "100 %"
    AjouterLigne lst, ok, "02 Environnement de l'offre", "100 %"
    AjouterLigne lst, partiel, "03 Montage contractuel", "45 %"
    AjouterLigne lst, ok, "     03.1 Mode de passation", "100 %"
    AjouterLigne lst, partiel, "     03.2 Intervention UGAP", "40 %"
    AjouterLigne lst, partiel, "     03.3 Durée", "50 %"
    AjouterLigne lst, vide, "     03.4 Clauses modificatives", "0 %"
    AjouterLigne lst, vide, "     03.5 Fin de contrat client", "0 %"
    AjouterLigne lst, partiel, "04 Devis et commande", "20 %"
    AjouterLigne lst, vide, "05 Garantie, SAV, maintenance", "0 %"
    AjouterLigne lst, vide, "06 Exécution financière", "0 %"
    AjouterLigne lst, vide, "07 Performance fournisseur", "0 %"
    AjouterLigne lst, vide, "08 Données personnelles", "0 %"
    AjouterLigne lst, vide, "09 Propriété intellectuelle", "0 %"
    AjouterLigne lst, vide, "10 Prix", "0 %"
    AjouterLigne lst, vide, "11 Modalités d'exécution", "0 %"
    AjouterLigne lst, vide, "12 Analyse des offres", "0 %"
    AjouterLigne lst, vide, "13 Commercialisation", "0 %"
    lst.ListIndex = 5
    Set c = Lbl(fr, "lblLegendeNav", ok & " terminé    " & partiel & " en cours    " & vide & " à faire", _
                M, ih - 26, iw, 22, False, 8, RGB(90, 90, 90))
    c.Font.Name = POLICE_SYMB

    ' --- Zone 2 : questions du thème -----------------------------------------
    Set fr = Cadre(f, "fraQuestions", "Questions  ·  9 affichées, 6 masquées", xLst, yCorps, wLst, hCorps)
    iw = fr.InsideWidth - 2 * M: ih = fr.InsideHeight
    Lbl fr, "lblFilAriane", "03 Montage contractuel  ›  03.3 Durée", M, 4, iw, 16, True, 10
    Lbl fr, "lblFiltre", "Afficher :", M, 26, 50, 16
    Combo fr, "cboFiltre", M + 50, 23, Borne(iw - 50, 120, 220), 19, _
          Array("Toutes les questions", "Non répondues", "À justifier", "À revoir", "Écarts au standard")
    Set lst = Liste(fr, "lstQuestions", M, 48, iw, ih - 48 - 30, "14;38;" & CLng(iw - 14 - 38 - 4))
    AjouterLigne lst, ok, "El80", "La durée du marché est-elle de 48 mois ?"
    AjouterLigne lst, ok, "El80a", "Quelle est la durée initiale du marché ?"
    AjouterLigne lst, ok, "El84", "Le principe de reconduction s'applique-t-il ?"
    AjouterLigne lst, ok, "El85b", "   " & fille & " Modalités de reconduction"
    AjouterLigne lst, enCours, "El84b", "   " & fille & " Nombre de reconductions possibles"
    AjouterLigne lst, vide, "El84c", "   " & fille & " Durée de chaque reconduction"
    AjouterLigne lst, vide, "El84d", "Ces modalités s'appliquent-elles à tous les lots ?"
    AjouterLigne lst, "!", "El86", "Durée des BDC au-delà de la fin du marché ?"
    AjouterLigne lst, vide, "El82", "Travaux d'interfaçage ou de déploiement long ?"
    lst.ListIndex = 4
    Set c = Lbl(fr, "lblLegendeQ", enCours & " en cours   " & ok & " répondue   ! à justifier   " & _
                ChrW(&H21BB) & " à revoir   " & vide & " à faire   " & fille & " question fille", _
                M, ih - 26, iw, 22, False, 8, RGB(90, 90, 90))
    c.Font.Name = POLICE_SYMB

    ' --- Zone 3 : la question et sa réponse ----------------------------------
    Set fr = Cadre(f, "fraDetail", "Question El84b", xDet, yCorps, wDet, hCorps)
    iw = fr.InsideWidth - 2 * M: ih = fr.InsideHeight
    If ih < 330 Then            ' petit écran : la zone défile au lieu d'être tronquée
        fr.ScrollBars = 2       ' fmScrollBarsVertical
        fr.ScrollHeight = 330
        ih = 330: iw = iw - 14
    End If
    Lbl fr, "lblTypeQ", "Question fille de El84 (réponse « Oui »)  ·  Liste à choix unique", _
        M, 4, iw, 13, False, 8, RGB(110, 110, 110)
    Lbl fr, "lblLibelle", "Si oui, quel est le nombre de reconductions possibles ?", _
        M, 18, iw, 34, True, 11

    Set sub_ = Cadre(fr, "fraStandard", "Standard applicable", M, 54, iw, 50)
    sub_.BackColor = RGB(255, 248, 220)
    Lbl sub_, "lblNiveauStd", "Établissement (E)", M, 2, 120, 13, False, 8, RGB(110, 110, 110)
    Lbl sub_, "lblValeurStd", "2 reconductions", M, 15, iw - 140, 18, True, 10
    Btn sub_, "btnReprendreStd", "Reprendre le standard", sub_.InsideWidth - M - 124, 6, 124, 22

    y = 112
    Lbl fr, "lblReponse", "Réponse", M, y, 100, 14, True
    Set lst = Liste(fr, "lstReponse", M, y + 15, iw, 62, "")
    lst.ListStyle = 1          ' fmListStyleOption : boutons radio (cases à cocher si choix multiple)
    lst.Font.Name = POLICE
    lst.AddItem "1 reconduction"
    lst.AddItem "2 reconductions"
    lst.AddItem "3 reconductions"
    lst.AddItem "4 reconductions"
    lst.ListIndex = 2

    y = y + 84
    Lbl fr, "lblConformite", "Conformité au standard", M, y, 200, 14, True
    wc = iw / 3
    Opt fr, "optConforme", "Conforme", M, y + 15, wc, 18, False
    Opt fr, "optSpecificite", "Spécificité (à justifier)", M + wc, y + 15, wc + 20, 18, True
    Opt fr, "optNonConcerne", "Non concerné", M + 2 * wc + 20, y + 15, wc - 20, 18, False

    y = y + 40
    Lbl fr, "lblJustif", "Justification de la dérogation", M, y, 200, 14, True
    Set c = Txt(fr, "txtJustif", "Alignement sur le cycle de vie des licences éditeur (3 ans) : " & _
                "une reconduction supplémentaire évite une remise en concurrence en cours de migration.", _
                M, y + 15, iw, Borne(ih - (y + 15) - 52, 40, 200))
    c.MultiLine = True: c.WordWrap = True

    y = ih - 46
    Lbl fr, "lblComTitre", "Commentaires (1)", M, y, 120, 14, True
    Lbl fr, "lblComDernier", "J. Martin · 03/10 : à confirmer avec l'éditeur avant publication.", _
        M, y + 15, iw - 90, 28, False, 8, RGB(90, 90, 90)
    Btn fr, "btnCommentaires", "Commentaires…", iw - 84 + M, y + 4, 84, 22

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

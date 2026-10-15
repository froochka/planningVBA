Option Explicit

' =============================================================================
' Outils d'affichage communs aux écrans SGO (frmGarde, frmSaisie).
' Mêmes principes que la maquette validée : contrôles MSForms standards
' uniquement, icônes de la police « Segoe MDL2 Assets », couleurs sobres,
' disposition calculée selon la taille de l'écran.
' =============================================================================

Public Const UI_POLICE As String = "Segoe UI"
Public Const UI_POLICE_SYMB As String = "Segoe UI Symbol"
Public Const UI_POLICE_ICONES As String = "Segoe MDL2 Assets"
Public Const UI_MARGE As Single = 8

' Icônes Segoe MDL2 Assets
Public Const ICO_RECHERCHE As Long = &HE721&
Public Const ICO_CRAYON As Long = &HE70F&
Public Const ICO_PRECEDENT As Long = &HE76B&
Public Const ICO_SUIVANT As Long = &HE76C&
Public Const ICO_COCHE As Long = &HE73E&
Public Const ICO_FERMER As Long = &HE711&

' Taille du texte (A- / A+), mémorisée pour l'utilisateur
Public uiTaille As Single

' Palette
Public uiBandeau As Long, uiAction As Long, uiCoulFond As Long, uiBlanc As Long, uiLigne As Long
Public uiTexte As Long, uiGris As Long, uiVert As Long, uiOrange As Long, uiRouge As Long
Public uiStdFond As Long, uiStdBarre As Long, uiOngletFond As Long, uiBleuClair As Long, uiRougeClair As Long

Public Sub UIInit()
    uiBandeau = RGB(22, 54, 92)
    uiAction = RGB(0, 99, 177)
    uiCoulFond = RGB(243, 245, 248)
    uiBlanc = RGB(255, 255, 255)
    uiLigne = RGB(214, 219, 226)
    uiTexte = RGB(33, 37, 41)
    uiGris = RGB(100, 108, 118)
    uiVert = RGB(24, 128, 56)
    uiOrange = RGB(160, 100, 0)
    uiRouge = RGB(176, 42, 30)
    uiStdFond = RGB(255, 249, 230)
    uiStdBarre = RGB(214, 158, 0)
    uiOngletFond = RGB(232, 236, 241)
    uiBleuClair = RGB(232, 240, 250)
    uiRougeClair = RGB(253, 236, 234)
    uiTaille = Val(GetSetting("SGO", "Affichage", "Taille", "10"))
    If uiTaille < 8 Or uiTaille > 16 Then uiTaille = 10
End Sub

Public Sub UIChangerTaille(ByVal delta As Single)
    uiTaille = UIBorne(uiTaille + delta, 8, 16)
    SaveSetting "SGO", "Affichage", "Taille", CStr(uiTaille)
End Sub

' Hauteur d'une ligne de texte pour la taille courante + delta
Public Function UILigneH(Optional ByVal delta As Single = 0) As Single
    UILigneH = (uiTaille + delta) * 1.45
End Function

Public Function UIBorne(ByVal v As Single, ByVal mini As Single, ByVal maxi As Single) As Single
    If v < mini Then
        UIBorne = mini
    ElseIf v > maxi Then
        UIBorne = maxi
    Else
        UIBorne = v
    End If
End Function

' Taille proportionnelle à la fenêtre Excel, bornée, centrée dessus
' (donc sur l'écran où se trouve Excel en cas d'écran supplémentaire)
Public Sub UIDimensionner(ByVal f As Object, ByVal ratio As Single, ByVal wMin As Single, ByVal hMin As Single, _
                          ByVal wMax As Single, ByVal hMax As Single)
    f.Width = UIBorne(Application.UsableWidth * ratio, wMin, wMax)
    f.Height = UIBorne(Application.UsableHeight * ratio, hMin, hMax)
    UICentrer f
End Sub

Public Sub UICentrer(ByVal f As Object)
    On Error Resume Next
    f.StartUpPosition = 0
    f.Left = Application.Left + (Application.Width - f.Width) / 2
    f.Top = Application.Top + (Application.Height - f.Height) / 2
End Sub

' -----------------------------------------------------------------------------
' Fabrique de contrôles
' -----------------------------------------------------------------------------
Public Function UICtl(ByVal parent As Object, progId As String, nom As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = parent.Controls.Add(progId, nom, True)
    c.Left = l: c.Top = t: c.Width = IIf(w < 1, 1, w): c.Height = IIf(h < 1, 1, h)
    c.Font.Name = UI_POLICE
    c.Font.Size = uiTaille
    Set UICtl = c
End Function

Public Sub UIBordurePlate(ByVal c As Object)
    On Error Resume Next
    c.SpecialEffect = 0
    c.BorderStyle = 1
    c.BorderColor = RGB(170, 178, 189)
End Sub

Public Function UILbl(ByVal parent As Object, nom As String, texte As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                      Optional ByVal gras As Boolean = False, Optional ByVal delta As Single = 0, _
                      Optional ByVal couleur As Long = -1) As Object
    Dim c As Object
    Set c = UICtl(parent, "Forms.Label.1", nom, l, t, w, h)
    c.Caption = texte
    c.BackStyle = 0
    c.WordWrap = True
    c.Font.Bold = gras
    c.Font.Size = uiTaille + delta
    c.ForeColor = IIf(couleur = -1, uiTexte, couleur)
    Set UILbl = c
End Function

Public Function UIFond(ByVal parent As Object, nom As String, ByVal l As Single, ByVal t As Single, _
                       ByVal w As Single, ByVal h As Single, ByVal couleur As Long) As Object
    Dim c As Object
    Set c = UILbl(parent, nom, "", l, t, w, h)
    c.BackStyle = 1
    c.BackColor = couleur
    Set UIFond = c
End Function

' Bouton plat : fond + icône + libellé, et par-dessus une zone cliquable
' transparente (le Label renvoyé, qui reçoit les clics).
' genre : 0 secondaire, 1 principal, 2 sur le bandeau
Public Function UIBouton(ByVal parent As Object, nom As String, texte As String, _
                         ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                         Optional ByVal genre As Long = 0, Optional ByVal icone As Long = 0) As Object
    Dim fond As Object, c As Object, xTexte As Single, coul As Long
    Set fond = UILbl(parent, nom & "_fond", "", l, t, w, h)
    fond.BackStyle = 1
    fond.BorderStyle = 1
    Select Case genre
        Case 1
            fond.BackColor = uiAction: fond.BorderColor = uiAction: coul = uiBlanc
        Case 2
            fond.BackColor = uiBandeau: fond.BorderColor = RGB(120, 150, 190): coul = uiBlanc
        Case Else
            fond.BackColor = uiBlanc: fond.BorderColor = RGB(170, 178, 189): coul = uiTexte
    End Select
    xTexte = l
    If icone <> 0 Then
        Set c = UILbl(parent, nom & "_ico", ChrW(icone), l + 8, t + (h - UILigneH()) / 2 + 2, 16, UILigneH(), False, 0, coul)
        c.Font.Name = UI_POLICE_ICONES
        If texte = "" Then
            c.Left = l + (w - 16) / 2
            c.TextAlign = 2
        Else
            xTexte = l + 20
        End If
    End If
    If texte <> "" Then
        Set c = UILbl(parent, nom & "_txt", texte, xTexte, t + (h - UILigneH()) / 2, w - (xTexte - l), UILigneH(), _
                      (genre = 1), 0, coul)
        c.TextAlign = 2
        c.WordWrap = False
    End If
    Set UIBouton = UILbl(parent, nom, "", l, t, w, h)   ' zone cliquable transparente, au-dessus
End Function

' Active ou grise un bouton plat créé par UIBouton
Public Sub UIBoutonActif(ByVal parent As Object, nom As String, ByVal actif As Boolean)
    On Error Resume Next
    parent.Controls(nom).Enabled = actif
    parent.Controls(nom & "_txt").ForeColor = IIf(actif, IIf(parent.Controls(nom & "_fond").BackColor = uiBlanc, uiTexte, uiBlanc), RGB(160, 166, 174))
    parent.Controls(nom & "_ico").ForeColor = parent.Controls(nom & "_txt").ForeColor
End Sub

' Titre de section en petites capitales + trait fin
Public Sub UISection(ByVal parent As Object, nom As String, texte As String, _
                     ByVal l As Single, ByVal t As Single, ByVal w As Single)
    UILbl parent, nom, texte, l, t, w, UILigneH(-2), True, -2, uiBandeau
    UIFond parent, nom & "_trait", l, t + UILigneH(-2) + 3, w, 1, uiLigne
End Sub

Public Function UITxt(ByVal parent As Object, nom As String, texte As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Set c = UICtl(parent, "Forms.TextBox.1", nom, l, t, w, h)
    c.Text = texte
    UIBordurePlate c
    Set UITxt = c
End Function

Public Function UIListe(ByVal parent As Object, nom As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                        largeursColonnes As String) As Object
    Dim c As Object
    Set c = UICtl(parent, "Forms.ListBox.1", nom, l, t, w, h)
    c.IntegralHeight = False
    UIBordurePlate c
    If largeursColonnes <> "" Then
        c.ColumnCount = UBound(Split(largeursColonnes, ";")) + 1
        c.ColumnWidths = largeursColonnes
    End If
    Set UIListe = c
End Function

Public Function UICombo(ByVal parent As Object, nom As String, _
                        ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                        elements As Variant, Optional ByVal libre As Boolean = False) As Object
    Dim c As Object, e As Variant
    Set c = UICtl(parent, "Forms.ComboBox.1", nom, l, t, w, h)
    c.Style = IIf(libre, 0, 2)          ' 0 : saisie libre possible ; 2 : liste seule
    UIBordurePlate c
    If IsArray(elements) Then
        For Each e In elements
            c.AddItem e
        Next e
    End If
    Set UICombo = c
End Function

Public Function UIOpt(ByVal parent As Object, nom As String, texte As String, _
                      ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, _
                      ByVal groupe As String) As Object
    Dim c As Object
    Set c = UICtl(parent, "Forms.OptionButton.1", nom, l, t, w, h)
    c.Caption = texte
    c.BackStyle = 0
    c.GroupName = groupe
    Set UIOpt = c
End Function

Public Sub UIAjouterLigne(ByVal lst As Object, ParamArray valeurs() As Variant)
    Dim i As Long
    lst.AddItem valeurs(0)
    For i = 1 To UBound(valeurs)
        lst.List(lst.ListCount - 1, i) = valeurs(i)
    Next i
End Sub

' Pictogramme de statut d'une question
Public Function UIPicto(ByVal statut As Long, Optional ByVal standardModifie As Boolean = False) As String
    If standardModifie And statut <> ST_MASQUEE Then
        UIPicto = ChrW(&H21BB)                       ' standard modifié : à revoir
        Exit Function
    End If
    Select Case statut
        Case ST_MASQUEE: UIPicto = ChrW(&H2298)
        Case ST_NON_REPONDUE: UIPicto = ChrW(&H25CB)
        Case ST_A_JUSTIFIER: UIPicto = "!"
        Case ST_INCOMPLET: UIPicto = "?"
        Case ST_SPECIFIQUE: UIPicto = ChrW(&H25C6)
        Case ST_NON_CONCERNEE: UIPicto = ChrW(&H2013)
        Case Else: UIPicto = ChrW(&H2713)
    End Select
End Function

' Pictogramme d'avancement d'une catégorie
Public Function UIPictoCategorie(ByVal c As cCategorieSGO) As String
    If c.Grisee() Then
        UIPictoCategorie = ChrW(&H2298)
    ElseIf c.NbAJustifier > 0 Then
        UIPictoCategorie = "!"
    ElseIf c.NbTraitees >= c.NbVisibles Then
        UIPictoCategorie = ChrW(&H2713)
    ElseIf c.NbTraitees > 0 Then
        UIPictoCategorie = ChrW(&H25D0)
    Else
        UIPictoCategorie = ChrW(&H25CB)
    End If
End Function

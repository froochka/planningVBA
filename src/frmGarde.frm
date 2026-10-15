Option Explicit

' =============================================================================
' Page de garde du dossier (décisions PG-1 à PG-6 du chef de projet).
' Formulaire vide dans le classeur : tout est construit par ce code.
' Échange avec l'appelant par gGarde / gGardeValidee (module SGO_Lancement).
' =============================================================================

Private WithEvents lstDirAchat As MSForms.ListBox
Private WithEvents lstDirStd As MSForms.ListBox
Private WithEvents lstUnivers As MSForms.ListBox
Private WithEvents btnValider As MSForms.Label
Private WithEvents btnAnnuler As MSForms.Label
Private chkConserver As MSForms.CheckBox
Private codesDirStd As Variant
Private codesUnivers As Collection

Private Const CHAMPS_TEXTE As String = "numero_procedure;libelle_procedure;acheteur;assistant;chef_produit;juriste;pilote;sponsor"

Private Sub UserForm_Initialize()
    On Error GoTo erreur
    UIInit
    Construire
    Exit Sub
erreur:
    MsgBox "Erreur à l'ouverture de la page de garde : " & Err.Description, vbExclamation, "SGO"
End Sub

Private Sub Construire()
    Dim W As Single, H As Single, xG As Single, xD As Single, wCol As Single, wLib As Single
    Dim y As Single, i As Long, c As Object, libelles As Variant, champs As Variant, k As Variant
    Dim M As Single

    M = UI_MARGE
    Me.Caption = "SGO - Page de garde du dossier"
    Me.BackColor = uiBlanc
    Me.Width = UIBorne(780 * uiTaille / 10, 700, Application.UsableWidth * 0.95)
    Me.Height = UIBorne(560 * uiTaille / 10, 480, Application.UsableHeight * 0.95)
    UICentrer Me
    W = Me.InsideWidth: H = Me.InsideHeight
    wCol = (W - 2 * M - 8 - 24) / 2
    xG = M + 4: xD = xG + wCol + 24
    wLib = UIBorne(wCol * 0.42, 130, 170)

    UIFond Me, "lblBandeau", 0, 0, W, 44, uiBandeau
    UILbl Me, "lblTitre", "Dossier de conception de l'offre  ·  Élaboration de l'offre", M + 4, 10, W - 2 * M, _
          UILigneH(4), True, 4, uiBlanc

    ' --- Colonne gauche : procédure et équipe projet -----------------------------
    champs = Split(CHAMPS_TEXTE, ";")
    y = 58
    UISection Me, "secProcedure", "PROCÉDURE", xG, y, wCol
    y = y + 26
    UILbl Me, "lbl_numero_procedure", "N° de procédure", xG, y + 3, wLib, UILigneH()
    UITxt Me, "txt_numero_procedure", Valeur("numero_procedure"), xG + wLib, y, 140, UILigneH() + 8
    y = y + UILigneH() + 14
    UILbl Me, "lbl_libelle_procedure", "Libellé de la procédure", xG, y + 3, wLib, UILigneH()
    UITxt Me, "txt_libelle_procedure", Valeur("libelle_procedure"), xG + wLib, y, wCol - wLib, UILigneH() + 8

    y = y + UILigneH() + 26
    UISection Me, "secEquipe", "ÉQUIPE PROJET", xG, y, wCol
    y = y + 26
    ' Ordre de tabulation validé : acheteur, assistant(e), chef produit, juriste, pilote, sponsor
    libelles = Array("Nom de l'acheteur", "Assistant(e) support achat", "Nom du chef produit", _
                     "Nom du juriste", "Pilote de l'offre (DPO / RPO / CPA / Leader)", "Nom du sponsor")
    For i = 0 To 5
        Set c = UILbl(Me, "lbl_" & champs(i + 2), CStr(libelles(i)), xG, y - 2, wLib - 6, 2 * UILigneH(-1), _
                      False, IIf(Len(libelles(i)) > 26, -1, 0))
        Set c = UITxt(Me, "txt_" & champs(i + 2), Valeur(CStr(champs(i + 2))), xG + wLib, y, wCol - wLib, UILigneH() + 8)
        y = y + UILigneH() + 14
    Next i
    For i = 0 To UBound(champs)
        Me.Controls("txt_" & champs(i)).TabIndex = i
    Next i

    ' --- Colonne droite : standards (listes ouvertes, décision PG-2) ----------------
    y = 58
    UISection Me, "secStandards", "STANDARDS APPLICABLES  (obligatoires)", xD, y, wCol
    y = y + 26
    UILbl Me, "lblDirAchat", "Direction achat", xD, y + 3, wLib, UILigneH()
    Set lstDirAchat = UIListe(Me, "lstDirAchat", xD + wLib, y, wCol - wLib, gDirAchat.Count * UILigneH() + 6, "")
    For Each k In gDirAchat.Keys
        lstDirAchat.AddItem k
        If StrComp(k, Valeur("direction_achat"), vbTextCompare) = 0 Then lstDirAchat.ListIndex = lstDirAchat.ListCount - 1
    Next k
    y = y + lstDirAchat.Height + 10

    UILbl Me, "lblDirStd", "Standard de direction", xD, y + 3, wLib, UILigneH()
    Set lstDirStd = UIListe(Me, "lstDirStd", xD + wLib, y, wCol - wLib, gDirStd.Count * UILigneH() + 6, "")
    codesDirStd = gDirStd.Keys
    For i = 0 To UBound(codesDirStd)
        lstDirStd.AddItem LibelleCode(CStr(codesDirStd(i)))
    Next i
    y = y + lstDirStd.Height + 10

    UILbl Me, "lblUnivers", "Standard d'univers", xD, y + 3, wLib, UILigneH()
    Set lstUnivers = UIListe(Me, "lstUnivers", xD + wLib, y, wCol - wLib, 6 * UILigneH() + 6, "")
    For i = 0 To UBound(codesDirStd)
        If StrComp(codesDirStd(i), Valeur("standard_direction"), vbTextCompare) = 0 Then lstDirStd.ListIndex = i
    Next i
    RemplirUnivers Valeur("standard_univers")
    y = y + lstUnivers.Height + 10

    Set chkConserver = UICtl(Me, "Forms.CheckBox.1", "chkConserver", xD + wLib, y, wCol - wLib, UILigneH() + 6)
    chkConserver.Caption = "Conserver les réponses déjà saisies"
    chkConserver.BackStyle = 0
    chkConserver.Value = True                      ' décision PG-5
    chkConserver.Visible = Not gGardeCreation
    y = y + UILigneH() + 8
    UILbl Me, "lblInfoStd", IIf(gGardeCreation, _
          "Ces choix déterminent les standards affichés pour chaque question. Ils restent modifiables ensuite.", _
          "En cas de changement de standard, une analyse d'impact est présentée avant d'appliquer le changement."), _
          xD, y, wCol, 3 * UILigneH(-2), False, -2, uiGris

    ' --- Boutons --------------------------------------------------------------------
    Set btnAnnuler = UIBouton(Me, "btnAnnuler", "Annuler", W - M - 4 - 238, H - 40, 110, 30, 0)
    Set btnValider = UIBouton(Me, "btnValider", IIf(gGardeCreation, "Créer le dossier", "Valider"), _
                              W - M - 4 - 120, H - 40, 120, 30, 1)
End Sub

Private Function Valeur(ByVal champ As String) As String
    If Not gGarde Is Nothing Then
        If gGarde.Exists(champ) Then Valeur = CStr(gGarde(champ))
    End If
End Function

' Univers du standard de direction choisi
Private Sub RemplirUnivers(ByVal selection As String)
    Dim k As Variant, dirChoisie As String
    lstUnivers.Clear
    Set codesUnivers = New Collection
    If lstDirStd.ListIndex < 0 Then Exit Sub
    dirChoisie = codesDirStd(lstDirStd.ListIndex)
    For Each k In gUnivers.Keys
        If StrComp(gUnivers(k)(0), dirChoisie, vbTextCompare) = 0 Then
            codesUnivers.Add CStr(k)
            lstUnivers.AddItem LibelleCode(CStr(k))
            If StrComp(k, selection, vbTextCompare) = 0 Then lstUnivers.ListIndex = lstUnivers.ListCount - 1
        End If
    Next k
    If lstUnivers.ListCount = 0 Then lstUnivers.AddItem "(aucun univers pour cette direction)"
End Sub

Private Sub lstDirStd_Change()
    RemplirUnivers ""
End Sub

Private Sub btnValider_Click()
    Dim champs As Variant, i As Long, manque As String
    If lstDirAchat.ListIndex < 0 Then manque = manque & vbCrLf & "- Direction achat"
    If lstDirStd.ListIndex < 0 Then manque = manque & vbCrLf & "- Standard de direction"
    If codesUnivers.Count > 0 And lstUnivers.ListIndex < 0 Then manque = manque & vbCrLf & "- Standard d'univers"
    If manque <> "" Then
        MsgBox "Champs obligatoires à renseigner :" & manque, vbExclamation, "SGO"
        Exit Sub
    End If
    If gGarde Is Nothing Then Set gGarde = NouveauDico()
    champs = Split(CHAMPS_TEXTE, ";")
    For i = 0 To UBound(champs)
        gGarde(champs(i)) = Trim$(Me.Controls("txt_" & champs(i)).Text)
    Next i
    gGarde("direction_achat") = lstDirAchat.List(lstDirAchat.ListIndex)
    gGarde("standard_direction") = codesDirStd(lstDirStd.ListIndex)
    If codesUnivers.Count > 0 Then
        gGarde("standard_univers") = codesUnivers(lstUnivers.ListIndex + 1)
    Else
        gGarde("standard_univers") = ""
    End If
    gGarde("conserver") = IIf(chkConserver.Value, "Oui", "Non")
    gGardeValidee = True
    Unload Me
End Sub

Private Sub btnAnnuler_Click()
    gGardeValidee = False
    Unload Me
End Sub



Private Sub CommandButton1_Click()

'permet d'avoir la question suivante

Dim noeud2 As Node
Dim fil_ariane_id2 As String

'fil_ariane_id2 = UserFormNav.TabStrip3.SelectedItem.Tag & "|" & question_en_cours

fil_ariane_id2 = Split(ListBox1.Tag, "-")(ListBox1.ListIndex)

'MsgBox fil_ariane_id2


Set noeud2 = UserFormNav.TreeView1.Nodes(fil_ariane_id2)
'On Error GoTo sst_suite

If ListBox1.ListIndex <> ListBox1.ListCount - 1 Then

ListBox1.ListIndex = ListBox1.ListIndex + 1
'MsgBox fil_ariane_id2 & " -> " & noeud2.Next.key

End If




Exit Sub

question_en_cours = Split(noeud2.Next.key, "|")(4)
'ListBox1.Selected (ListBox1.ListIndex + 1)

Exit Sub



sst_suite:
On Error GoTo theme_suite
'UserFormNav.TabStrip3.SelectedItem = UserFormNav.TabStrip3.Tabs(UserFormNav.TabStrip3.SelectedItem.Index + 1)
MsgBox UserFormNav.TabStrip2.Tabs.Count
UserFormNav.TabStrip3.SelectedItem = UserFormNav.TabStrip3.Tabs(UserFormNav.TabStrip3.SelectedItem.Index + 1)

UserFormNav.TabStrip3.Value = UserFormNav.TabStrip3.Value + 1

MsgBox "pas d'autres questions !" & vbCr & "ss theme actuel : " & UserFormNav.TabStrip3.SelectedItem.Caption & _
        "le suivant :" & UserFormNav.TabStrip3.SelectedItem.Caption
        
Exit Sub

theme_suite:

Exit Sub


'permet d'avoir le sous-theme suivant à partir du sous theme en cours
'Dim noeud As Node
'Dim fil_ariane_id As String

fil_ariane_id = UserFormNav.TabStrip3.SelectedItem.Tag
MsgBox fil_ariane_id


Set noeud = UserFormNav.TreeView1.Nodes(fil_ariane_id)
On Error Resume Next

MsgBox noeud.Next.key

End Sub

Private Sub CommandButton2_Click()

'compteur_question = compteur_question + 1


'permet d'avoir la question suivante

Dim noeud2 As Node
Dim fil_ariane_id2 As String

'fil_ariane_id2 = UserFormNav.TabStrip3.SelectedItem.Tag & "|" & question_en_cours

fil_ariane_id2 = Split(ListBox1.Tag, "-")(ListBox1.ListIndex)

'MsgBox fil_ariane_id2


Set noeud2 = UserFormNav.TreeView1.Nodes(fil_ariane_id2)
'On Error GoTo sst_suite

If ListBox1.ListIndex <> 0 Then

ListBox1.ListIndex = ListBox1.ListIndex - 1
'MsgBox fil_ariane_id2 & " -> " & noeud2.Next.key

End If


End Sub

Private Sub Frame3_Click()
On Error Resume Next

ProgressBar1.Value = ProgressBar1.Value + 10

  Dim pourcent As Integer
    pourcent = ProgressBar1.Value
    
    ' Calculer la largeur (FrameProgress = conteneur de 300px)
    LabelProgress.Width = (pourcent / 100) * UserForm2.Width
    
    
    ' Changer la couleur
    If pourcent < 30 Then
        LabelProgress.BackColor = RGB(255, 0, 0)      ' Rouge
      
            
    ElseIf pourcent < 50 Then
        LabelProgress.BackColor = RGB(255, 165, 0)   ' Orange
    
     ElseIf pourcent < 70 Then
        LabelProgress.BackColor = RGB(0, 0, 255)   ' Orange
        
    Else
        LabelProgress.BackColor = RGB(0, 255, 0)     ' Vert
    End If
    
    LabelProgress.Caption = pourcent & "%"
    Me.Repaint
    

End Sub






Private Sub ListBox1_Click()

'Stop
'On Error Resume Next

' temp_id contient l'id de la question "El1"
' temp_obj correspond a l'objet class de la question en cours par exemple
' mode_reponse est lu dans l'objet classe dans le dico grace à la clé "Mode de Réponse"
' mon_standard est le standard de la question en cours, il est lié au informations de la page de garde (a developper)
' mon_dico permet de lister les standards dans un dico temporaire

Dim temp_obj, mon_dico, c

Set mon_dico = CreateObject("scripting.dictionary")


Dim temp_id, mode_reponse, mon_standard As String




'MsgBox Trim(Split(ListBox1.List(ListBox1.ListIndex), "-")(0))
'MsgBox Trim(Split(ListBox1.Tag, "-")(ListBox1.ListIndex))

temp_id = Trim(Split(ListBox1.Text, "-")(0))

question_en_cours = temp_id

Set temp_obj = dico_standards(question_en_cours)

Label4.Caption = temp_obj.lireValeur("CHAPITRES DU DCO")

'mon_standard = temp_obj.lireValeur("DAV_Standard Etablissement (E) / Direction (D) / Univers (U) ")
'mon_standard = temp_obj.lireValeur(ma_direction_achat)

'mon_standard = mon_standard & temp_obj.lireValeur(mes_standards_direction)

'mon_standard = mon_standard & temp_obj.lireValeur(mes_standards_univers)
mon_standard = Est_Standard_Etablissement


Label_Standard.Caption = mon_standard


mode_reponse = temp_obj.lireValeur("Mode de Réponse")

Set mon_dico = temp_obj.listerdico

TextBox2.Text = ""


For Each c In mon_dico

TextBox2.Text = TextBox2.Text & c & ":" & mon_dico(c) & vbCr

Next c

'Stop

'MsgBox temp_obj.lireValeur("DAV")
'MsgBox temp_obj.lireValeur("E")


Exit Sub


'demo


'Texte
'Cartographie ORECA
'Liste choix multiple
'Liste choix unique
'Texte


Label4.Caption = ListBox1.List(ListBox1.ListIndex)

Label_compt_question.Caption = Str(ListBox1.ListIndex + 1) & "/" & ListBox1.ListCount

Dim hazard, bloc


'hazard = Int(Rnd(Timer) * 3)

Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False

Label10.Caption = choix_reponse

Select Case mode_reponse


Case "Texte"

Frame5.Visible = True
Frame5.Left = 20
Frame5.Width = 760
Frame5.Height = 120

Label_Standard.Caption = mon_standard   '"Champs libre, saisir l'objet de la procédure en quelques mots"

Case "Liste choix unique"

Frame6.Visible = True
Frame6.Left = 20
Frame6.Width = 760
Frame6.Height = 120

Label_Standard.Caption = mon_standard   '"Liste fermée, choisir une des propositions"

Case "Liste choix multiple"

Frame7.Visible = True
Frame7.Left = 20
Frame7.Width = 860
Frame7.Height = 120

Label_Standard.Caption = mon_standard   '"Choix multiple, choisir les elements parmi la liste, plusieurs choix possibles (CTRL + click)"

End Select

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim temp_id, temp_obj
Dim mon_standard_dir, texte_temp, prefixe As String




'MsgBox Trim(Split(ListBox1.List(ListBox1.ListIndex), "-")(0))
'MsgBox Trim(Split(ListBox1.Tag, "-")(ListBox1.ListIndex))

temp_id = Trim(Split(ListBox1.Text, "-")(0))

Set temp_obj = dico_standards(temp_id)

'MsgBox temp_obj.lireValeur("DAV")


'MsgBox temp_obj.lireValeur("E")
'prefixe = "_Standard Etablissement(E) / Direction(D) / Univers(U)"

mon_standard_dir = colonne_Standard_direction("Standard_" & mes_standards_direction)

'Stop

MsgBox mes_standards_direction & " " & mon_standard_dir


prefixe = "_Standard Etablissement(E) / Direction(D) / Univers(U)"
texte_temp = ""
texte_temp = texte_temp & temp_obj.lireValeur(mes_standards_direction) & vbCr
texte_temp = texte_temp & temp_obj.lireValeur("E") & vbCr
texte_temp = texte_temp & temp_obj.lireValeur(CStr(mon_standard_dir))
'temp_obj.lireValeur ("E")


'MsgBox temp_obj.lireValeur(mes_standards_univers)
MsgBox texte_temp


'UserFormNav.TextBox1.Text = UserFormNav.TextBox1.Text & Mid(c, InStr(1, c, "-"), Len(c)) & vbCrLf

'MsgBox ListBox1.ItemData(ListBox1.ListIndex)
'Stop
End Sub

Private Sub ListBox2_Change()
'Label_reponse_multi = ListBox2.Selected

For i = 0 To ListBox2.ListCount - 1
    If ListBox2.Selected(i) Then
        strSelection = strSelection & ListBox2.List(i) & ", "
    End If
Next i

' Supprimer la dernière virgule
If Len(strSelection) > 0 Then
    strSelection = Left(strSelection, Len(strSelection) - 2)
End If

Label_reponse_multi.Caption = strSelection
End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub TabStrip1_Change()
'MsgBox Me.TabStrip1.Name
Label_fleche.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / 12) - 5
Label_fleche2.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / 12) - 5

'Label_fleche.Top = TabStrip1.Tabs(TabIndex).Top
'MsgBox (TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height))



End Sub



Private Sub ProgressBar1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub ToggleButton1_Click()
    If ToggleButton1.Value = True Then
        ToggleButton2.Value = False
    End If
End Sub

Private Sub ToggleButton2_Click()
    If ToggleButton2.Value = True Then
        ToggleButton1.Value = False
    End If
End Sub

Private Sub UserForm_Activate()
Label19.Caption = "Direction HA : " & ma_direction_achat & "  /  Direction : " & mes_standards_direction & "  / Univers : " & mes_standards_univers



End Sub

Private Sub UserForm_Initialize()
Dim mon_compt As Integer   ' compteur tempo
Dim temp As String ' variable tempo

Me.ListBox1.ColumnHeads = False

'Stop

'init des variables globales pour le test

Call init_global
Call init_questions 'module de lecture des questions en memeoire
Call creer_categories ' creation du dico des categories*

'Me.TabStrip1.TabOrientation = fmTabOrientationLeft

'mon_compt = 0

'For Each c In dico_cat
'temp = dico_cat(c).ma_type

'If temp = "CHAP" Then
'Stop
'temp = dico_cat(c).ma_lib

'Me.TabStrip1.Tabs(mon_compt).Caption = dico_cat(c).ma_lib



'mon_compt = mon_compt + 1
'Debug.Print dico_cat(c).ma_lib & "__";
'End If


'Stop

'Next c

Exit Sub

'Me.TabStrip1.Tabs(0).Caption = "Objet de la procédure"

'Me.TabStrip1.Tabs(1).Caption = "Domaine d'activité"

'Me.TabStrip1.Tabs(2).Caption = "Bénéficiaires"

'Me.TabStrip1.Tabs(3).Caption = "UGAP Bénéficiaire"


ListBox1.AddItem "Des actions ou des particularités s'appliquent-elles pour le paiement du client vis-à-vis de l'UGAP et/ou de l'UGAP vis-à-vis du fournisseur ?"
'ListBox1.List(ListBox1.ListCount - 1, 1) = "À traiter"  ' ou "Fait"

ListBox1.AddItem "Si oui, préciser si celles-ci concernent  le client et les particularités qui s'appliquent"
ListBox1.AddItem "Si oui, préciser si celles-ci concernent le fourniseur et les particularités qui s'appliquent"
ListBox1.AddItem "Si oui, préciser si celles-ci concernent l'UGAP et les particularités qui s'appliquent"
ListBox1.AddItem "L 'offre prévoit-elle un DDP ou Delivered Duty Paid (Rendu droits acquittés) en cas de livraison outre-mer"

ListBox1.AddItem "Des actions ou des particularités s'appliquent-elles pour le paiement du client vis-à-vis de l'UGAP et/ou de l'UGAP vis-à-vis du fournisseur ?"
ListBox1.AddItem "Si oui, préciser si celles-ci concernent  le client et les particularités qui s'appliquent"
ListBox1.AddItem "Si oui, préciser si celles-ci concernent le fourniseur et les particularités qui s'appliquent"
ListBox1.AddItem "Si oui, préciser si celles-ci concernent l'UGAP et les particularités qui s'appliquent"
ListBox1.AddItem "L 'offre prévoit-elle un DDP ou Delivered Duty Paid (Rendu droits acquittés) en cas de livraison outre-mer"

ProgressBar1.Value = 1



  ' Configuration initiale
    LabelProgress.Width = 0
    LabelProgress.Height = 20
    LabelProgress.BackColor = vbBlue
    LabelProgress.BorderStyle = fmBorderStyleSingle
    
 '     Me.Width = Application.Width
 '   Me.Height = Application.Height
 '   Me.Top = 0
 '   Me.Left = 0
    
   ' Me.MultiPage1.Pages(1).Caption = "Périmètre de l'offre"
   ' Me.MultiPage1.Pages(2).Caption = "Engagements"
   '  Me.MultiPage1.Pages(3).Caption = "Allotissement"
   '     Me.MultiPage1.Pages(4).Caption = "Etendue géographique de l'offre"
   '  Me.MultiPage1.Pages(5).Caption = "Modalités / Actions spécifiques ou particulières sont identifiées pour le paiement"
   '   Me.MultiPage1.Pages(6).Caption = "Périmètre de l'offre"
   '    Me.MultiPage1.Pages(7).Caption = "Périmètre de l'offre"
   
   ListBox2.AddItem "Choix 1"
   ListBox2.AddItem "Choix 2"
   ListBox2.AddItem "Choix 3"
   ListBox2.AddItem "Choix 4"
   



End Sub

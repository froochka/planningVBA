Private Sub CommandButton1_Click()


ma_direction_achat = ListBox1.Text
mes_standards_direction = Range(ListBox2.Text)
mes_standards_univers = Range(ListBox3.Text)



'For i = 1 To 1000
'ProgressBar1.Value = ProgressBar1.Value + 0.1
'DoEvents

'Next


'MsgBox ("ok")

'UserForm1.Hide

'UserForm2.Show

UserFormNav.Show


End Sub

Private Sub CommandButton2_Click()
compteur_question = compteur_question + 1

End Sub





Private Sub TabStrip1_Change()
'MsgBox Me.TabStrip1.Name


End Sub





Private Sub ProgressBar1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)

End Sub




Private Sub CommandButton3_Click()
UserForm_Admin.Show

End Sub

Private Sub CommandButton4_Click()
UserForm2.Show
End Sub

Private Sub Frame3_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

Dim key As String
Dim valeur As String
Dim c As Variant

key = ListBox2.List(ListBox2.ListIndex)
valeur = mon_dico_dir(key)  ' Récupère la valeur associée à la clé
ListBox3.Clear

For Each c In Split(valeur, ",")
    ListBox3.AddItem Trim(c)  ' Trim enlève les espaces éventuels
Next c

'ListBox3.AddItem mon_dico_dir(ListBox2.List(ListBox2.ListIndex))
End Sub

Private Sub UserForm_Initialize()
'init des variables globales pour le test


Set mon_dico_dir = CreateObject("scripting.dictionary")

mon_dico_dir("DAV") = "VI,VL,ISP"

mon_dico_dir("DAI") = "PII,Materiel_et_Prestations,Copieurs,Telecom,Bureau WEB,Logiciels"

mon_dico_dir("DS") = "Biomedical,EQ Soins et Secours,Laboratoire,EQ Lourds,SAD,Gauss"

mon_dico_dir("DAG Mobilier scolaire et Equipement général") = "Mobilier scolaire,Equipement général"

mon_dico_dir("DAG SOFI") = "Services aux batiments,Services aux occupants,Offres Financières"

mon_dico_dir("Non standardisé") = "Non standardisé"


'Stop
'Call init_global

'ProgressBar1.Value = 0



  ' Configuration initiale

    
 '     Me.Width = Application.Width
 '   Me.Height = Application.Height
 '   Me.Top = 0
 '   Me.Left = 0
 For Each c In mon_dico_dir
    ListBox2.AddItem c
 Next c
 
    ListBox1.AddItem "DAG"
    ListBox1.AddItem "DAI"
    ListBox1.AddItem "DAV"
    ListBox1.AddItem "DS"

    Call OrdreTabulation

End Sub

' Ordre de passage avec la touche Tab entre les zones de saisie.
' Chaque TextBox est seule dans son propre cadre (Frame) : c'est donc l'ordre
' des cadres, et non celui des TextBox, qui determine le passage de l'une a l'autre.
Private Sub OrdreTabulation()
    Dim noms As Variant, i As Long
    noms = Array("TextBox4", "TextBox8", "TextBox6", "TextBox10", "TextBox7", "TextBox5")
    For i = 0 To UBound(noms)
        With Me.Controls(noms(i))
            .TabStop = True
            .TabIndex = 0               ' premiere position dans son cadre
            .Parent.TabIndex = i        ' position du cadre dans le cadre englobant
        End With
    Next i
End Sub

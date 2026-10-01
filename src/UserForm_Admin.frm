

Private Sub CommandButton1_Click()
UserForm1.Show

End Sub

Private Sub CommandButton2_Click()
UserForm2.Show
End Sub

Private Sub UserForm_Initialize()

ListBox1.AddItem "Établissement (E)"
ListBox1.AddItem "Direction (D) "
ListBox1.AddItem "Univers (U)"
ListBox1.AddItem "Spécifique offre (Sp)"
ListBox1.AddItem "Non standardisé"


End Sub

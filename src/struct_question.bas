'permet de creer les colonnes chapitre, theme et sous theme de chque question
' dans l'onglet structure_question
' à refaire si changement de questions ou de theme...


'*************************************
' ecriture sur la feuille "structure_question"
'*************************************

Sub creer_struct_quest()

Stop

Dim mon_theme, mon_chap, mon_sstheme As String


ma_sel = "$A$2:$A$549"

For Each c In Range(ma_sel)

If c.Value Like "*SST*" Then

mon_sstheme = c.Value


ElseIf c.Value Like "*THEME*" Then

mon_theme = c.Value

ElseIf c.Value Like "*CHAP*" Then

mon_chap = c.Value

Else


c.Offset(0, 2).Value = mon_chap
c.Offset(0, 3).Value = mon_theme
c.Offset(0, 4).Value = mon_sstheme

End If


Next c


End Sub


' creer une classe pour les categories
' sachant qu'on a chapitre, theme et sous theme
' on veut pour chacun avoir :
' ID, lib, parent_id, type (chap, thme, ss thme), liste des enfants (theme, ss theme)
' les questions seront donc les enfants des ss themes ( a prevoir)

Sub creer_categories()


Dim compt
Dim mon_num, mon_id, mon_lib, mon_parent, mon_type, mon_theme, mon_chap, mon_sstheme As String


'Stop


ma_sel = "$A$5:$A$137"

For Each c In ws_categories.Range(ma_sel)

'If c.Value Like "*TTT*" Or c.Value Like "*SST*" Then
If c.Value Like "*SST*" Then

mon_id = c.Value

mon_lib = c.Offset(0, 1).Value
mon_parent = mon_theme
mon_type = "SSTHEME"
'mon_sstheme = mon_id

ElseIf c.Value Like "*THEME*" Then

mon_id = c.Value

mon_lib = c.Offset(0, 1).Value
mon_parent = mon_chap
mon_type = "THEME"
mon_theme = mon_id

ElseIf c.Value Like "*CHAP*" Then

mon_num = compt
mon_id = c.Value

mon_lib = c.Offset(0, 1).Value
mon_parent = "ROOT"
mon_type = "CHAP"
mon_chap = mon_id

compt = compt + 1

End If

'Dim ma_cat As Ccategorie
Set ma_cat = New Ccategorie   ' Création de l'objet

ma_cat.ma_id = mon_id
ma_cat.ma_type = mon_type
ma_cat.ma_lib = mon_lib
ma_cat.ma_parent = mon_parent

dico_cat.Add mon_id, ma_cat


'Set ma_cat = Ccategorie

'Stop

'c.Offset(0, 2).Value = mon_chap
'c.Offset(0, 3).Value = mon_theme
'c.Offset(0, 4).Value = mon_sstheme





Next c

'Stop
'Dim compt

'For Each element In dico_cat
'temp = dico_cat(element).ma_type
'If dico_cat(element).ma_type = "CHAP" Then
'If temp = "CHAP" Then
'compt = compt + 1 ' dico_cat(element).ma_lib

'End If

'Next element
'Stop

End Sub


Sub creer_categories_from_treeview()


Dim compt
Dim mon_num, mon_id, mon_lib, mon_parent, mon_type, mon_theme, mon_chap, mon_sstheme As String


ma_sel = "$A$5:$A$137"

For Each c In ws_categories.Range(ma_sel)

'If c.Value Like "*TTT*" Or c.Value Like "*SST*" Then
If c.Value Like "*SST*" Then

mon_id = c.Value

mon_lib = c.Offset(0, 1).Value
mon_parent = mon_theme
mon_type = "SSTHEME"
'mon_sstheme = mon_id

ElseIf c.Value Like "*THEME*" Then

mon_id = c.Value

mon_lib = c.Offset(0, 1).Value
mon_parent = mon_chap
mon_type = "THEME"
mon_theme = mon_id

ElseIf c.Value Like "*CHAP*" Then

mon_num = compt
mon_id = c.Value

mon_lib = c.Offset(0, 1).Value
mon_parent = "ROOT"
mon_type = "CHAP"
mon_chap = mon_id

compt = compt + 1

End If

'Dim ma_cat As Ccategorie
Set ma_cat = New Ccategorie   ' Création de l'objet

ma_cat.ma_id = mon_id
ma_cat.ma_type = mon_type
ma_cat.ma_lib = mon_lib
ma_cat.ma_parent = mon_parent

dico_cat.Add mon_id, ma_cat


'Set ma_cat = Ccategorie

'Stop

'c.Offset(0, 2).Value = mon_chap
'c.Offset(0, 3).Value = mon_theme
'c.Offset(0, 4).Value = mon_sstheme





Next c



End Sub

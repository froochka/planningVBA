Option Explicit

' =============================================================================
' Une ligne cliquable d'une liste « maison » de frmSaisie (chapitres, questions).
' Les listes sont faites de Labels, pour que le texte passe à la ligne et que
' les lignes sans objet soient vraiment grisées (une ListBox ne sait faire ni
' l'un ni l'autre). Le clic est transmis au formulaire : LigneCliquee.
' =============================================================================

Public WithEvents Zone As MSForms.Label
Public Formulaire As Object
Public Liste As String              ' "C" : chapitres ; "Q" : questions
Public Index As Long                ' position dans la liste (0 = première ligne)

Private Sub Zone_Click()
    Formulaire.LigneCliquee Liste, Index
End Sub

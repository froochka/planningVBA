Private Sub Workbook_Open()
    ' Version de test : le classeur s'ouvre sur la page « Démarrer » si elle existe
    On Error Resume Next
    ThisWorkbook.Worksheets("Démarrer").Activate
End Sub

' Classeur de l'outil fermé pendant la saisie : l'écran de saisie est fermé
' proprement (réponse en cours enregistrée, DCO réaffiché)
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim i As Long
    On Error Resume Next
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        If VBA.UserForms(i).Name = "frmSaisie" Then Unload VBA.UserForms(i)
    Next i
End Sub

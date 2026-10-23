Private Sub Workbook_Open()
    ' Version de test : le classeur s'ouvre sur la page « Démarrer » si elle existe
    On Error Resume Next
    ThisWorkbook.Worksheets("Démarrer").Activate
End Sub

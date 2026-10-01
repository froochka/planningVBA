Public ma_num As String 'son numero dans l'ordre
Public ma_id As String
Public ma_lib As String
Public ma_type As String
Public ma_parent As String
Public mes_enfants As Object

Private Sub Class_Initialize()
    
    Set mes_enfants = CreateObject("Scripting.Dictionary")
    
End Sub

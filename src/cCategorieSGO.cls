Option Explicit

' =============================================================================
' Un chapitre, un thème ou un sous-thème du référentiel (R_Categories),
' avec son avancement calculé par le moteur.
' =============================================================================

Public Id As String
Public TypeCat As String            ' CHAPITRE, THEME, SOUS_THEME
Public Libelle As String
Public Parent As String             ' id du parent ("" pour un chapitre)
Public Ordre As Long
Public Enfants As Collection        ' catégories filles, dans l'ordre
Public Questions As Collection      ' questions (sous-thèmes uniquement), dans l'ordre

' --- Calculé par SGO_Moteur.CalculerAvancement -------------------------------------
Public Visible As Boolean           ' faux si une condition masque la catégorie
Public NbVisibles As Long           ' questions visibles (sous-thèmes inclus)
Public NbTraitees As Long           ' questions visibles conformes, justifiées, spécifiques ou non concernées
Public NbAJustifier As Long

Private Sub Class_Initialize()
    Set Enfants = New Collection
    Set Questions = New Collection
End Sub

' Grisée : masquée par une condition, ou sans aucune question visible
Public Function Grisee() As Boolean
    Grisee = (Not Visible) Or NbVisibles = 0
End Function

Public Function Pourcentage() As Long
    If NbVisibles > 0 Then Pourcentage = Int(100 * NbTraitees / NbVisibles)
End Function

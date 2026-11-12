Option Explicit

' =============================================================================
' Surveillance d'Excel pendant la saisie (démarrée par SGO_Lancement.AfficherSaisie,
' arrêtée à la fermeture de l'écran de saisie) :
'   - une autre fenêtre Excel que celle du DCO est activée (autre classeur,
'     feuille Remarques...) : l'écran de saisie passe en mode réduit ;
'   - fermeture du DCO demandée : refusée tant que l'écran de saisie est ouvert.
' =============================================================================

Public WithEvents xl As Application

Private Sub xl_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If EstLeDossier(Wb) Then Exit Sub
    ReduireSaisie
End Sub

Private Sub xl_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If Not EstLeDossier(Wb) Then Exit Sub
    If SaisieOuverte() Is Nothing Then Exit Sub
    Cancel = True
    MsgBox "Ce DCO est en cours de saisie." & vbCrLf & _
           "Fermez d'abord l'écran de saisie (bouton « Fermer ») : le DCO y est enregistré automatiquement.", _
           vbInformation, "SGO"
End Sub

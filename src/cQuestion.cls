'classe contient les objets questions
'avec toutes les propriétés du DCO
' a definir

'********************************************************************
'Stockage:  mes_qestions.Add "Clé", objQuestion

'Récupération:  Set obj = mes_qestions("Clé")

'Utilisation : obj.LireValeur("CléDico") ou obj.AjouterValeur ...
'********************************************************************

'entete de colonne de elaboration
'Public entete_Elaboration
'Public un_Num As String
'Public de_nb As String
'Public tr_standard As String
'Public qu_cas As String
'Public ci_chapitre As String
'Public si_Mode_reponse As String
'Public se_Choix_reponse As String
'Public hu_Reponses As String
'Public ne_Question_Conditionnante As String
'Public di_Non_concerne As String
'Public on_Conforme_specificite As String
'Public do_Diff As String
'Public treize_Justifier_dérogation As String
'Public quatorze_Commentaires As String
'Public quinze_Commentaires_Intra As String
'Public seize_Operateur_STANDARD As String

 ' contient les 16 colonne de la feuille "3-Elaboration.O" avec une clé simple
Private dico_question_champs ' dico de chaque question clé = ID de la question
' Public entete_Elaboration 'entete des 16 colonnes ' clé du dico =
 
 
Private Sub Class_Initialize()

    Set dico_question_champs = CreateObject("Scripting.Dictionary")
   ' dico_question("Numero") = "un pour commencer"
   
    ' Ajoutez ici vos initialisations
End Sub



' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
' Méthodes publiques
' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Public Sub AjouterValeur(cle As String, valeur As String)
    If dico_question_champs.exists(cle) Then
        dico_question_champs(cle) = valeur
    Else
        dico_question_champs.Add cle, valeur
    End If
End Sub


' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
' Méthodes publiques
' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Function lireValeur(cle As String) As String


  If dico_question_champs.exists(cle) Then
       lireValeur = dico_question_champs(cle)
       Else
        lireValeur = ""
        
    End If
End Function


' a revoir comme sur la classe CStandard qui est ok

Function listerdico()
listerdico = dico_question_champs.items
End Function

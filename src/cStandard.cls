Option Explicit

'classe contient les objets questions standards
'avec un dictionnaire qui contient chaque entete de colonne de la feuille standard, par exemple :
'colonne G, entete est "E", cela permet de stocker clé "E" et valeur :"Renouvellement d'offres avec extension"



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

 ' contient les colonnes de la feuille "14-Standards" avec une clé simple qui est l'entete du champs
Private dico_standards_col ' dico de chaque question clé = ID de la question
' Public entete_Elaboration 'entete des 16 colonnes ' clé du dico =
 
 
Private Sub Class_Initialize()

    Set dico_standards_col = CreateObject("Scripting.Dictionary")
   ' dico_question("Numero") = "un pour commencer"
   
    ' Ajoutez ici vos initialisations
End Sub



' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
' Méthodes publiques
' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Public Sub AjouterValeur(cle As String, valeur As String)
    If dico_standards_col.exists(cle) Then
        dico_standards_col(cle) = valeur
    Else
        dico_standards_col.Add cle, valeur
    End If
End Sub


' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
' Méthodes publiques
' ¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Function lireValeur(cle As String) As String


  If dico_standards_col.exists(cle) Then
       lireValeur = dico_standards_col(cle)
       Else
        lireValeur = ""
        
    End If
End Function


Function listerdico() As Object
'Stop
'listerdico = dico_standards_col.items
'Function listerdico() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim cle As Variant
    For Each cle In dico_standards_col.Keys
        dict.Add cle, dico_standards_col(cle)
    Next cle
'Stop

    Set listerdico = dict
End Function

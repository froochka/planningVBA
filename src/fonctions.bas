Function numero_derlig(adresse As Range)
Stop

numero_derlig = Range(adresse).End(xlDown).Row


End Function


'Fonction permet de controler pour une question le standard a prendre en fonction de la direct, univsers ...
' première etape : Je donne en paramètre l'id de la question, j'affcihe le standard Etablissement

Function Est_Standard_Etablissement()

Dim obj_temp As Object

'MsgBox question_en_cours

Set obj_temp = dico_standards(question_en_cours)


Est_Standard_Etablissement = obj_temp.lireValeur("E")



End Function

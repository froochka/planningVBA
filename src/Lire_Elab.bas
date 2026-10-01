
Option Explicit

'****************************************************************
        
'lire les questions de Elaboration et les stocker en memeoire

'****************************************************************

'SHEET_ELABORATION

'Public entete_Elaboration


Sub test()

Dim ws As Sheets
Dim c, i, col, temp, tempo

Dim obj1 As Object

 
'dico des questions de Elaborations accessibles via l'id de la question
Set mes_questions = CreateObject("scripting.dictionary")

'dico des questions de Elaborations accessibles via l'id de la question
Set dico_standards = CreateObject("scripting.dictionary")


   '   Call init_global
          
        '***********************************************************************
       'boucle
       '************************************************************************
                 
         ' on saute les colonnes vides au nbre de 3, on garde un compteur de col de 1 à 16
          ' GARDER LES COLONNES MASQUEES : E, L, P
         
     tempo = Timer: Debug.Print tempo
   For i = 2 To ThisWorkbook.Sheets(SHEET_ELABORATION).Range("A2").End(xlDown).Row
  ' For Each c In ThisWorkbook.Sheets(SHEET_ELABORATION).Range("A2" & ":A" & ThisWorkbook.Sheets(SHEET_ELABORATION).Range("A2").End(xlDown).Row)
'   Stop
   
 '  Set obj1 = Nothing
   
   Set obj1 = New cQuestion
   
   ThisWorkbook.Sheets(SHEET_ELABORATION).Activate

   For Each c In ThisWorkbook.Sheets(SHEET_ELABORATION).Range("A1:S1").Cells
   
    If Not c.EntireColumn.Hidden Then
    col = col + 1
       
   obj1.AjouterValeur Trim(Str(col)), Cells(i, c.Column).Value
  ' Stop
   
   End If
   
   Next c
'   Stop
   
   mes_questions.Add obj1.lireValeur("1"), obj1
   col = 0
   
   Next i
 
 Debug.Print Timer - tempo 'tempo = Timer
 
' Stop
 
         '***********************************************************************
       'boucle pour les STANDARDS
       '************************************************************************
                 
            col = 0
            
     tempo = Timer: Debug.Print tempo
     On Error GoTo erreur
   For i = 6 To ThisWorkbook.Sheets(SHEET_STANDARDS).Range("A1").End(xlDown).Row
   
  If Cells(i, 8).Value = "Titre" Then GoTo suite
  
   Set obj1 = New cStandard
  ' Stop
   
  ' For Each c In ThisWorkbook.Sheets(SHEET_ELABORATION).Range("A2" & ":A" & ThisWorkbook.Sheets(SHEET_ELABORATION).Range("A2").End(xlDown).Row)
'   Stop
   
 'On Error GoTo erreur:
   
   ThisWorkbook.Sheets(SHEET_STANDARDS).Activate

   For Each c In ThisWorkbook.Sheets(SHEET_STANDARDS).Range("A1:AN1").Cells
   
                   '  Set obj1 = Nothing
                                    
                '  If Not c.EntireColumn.Hidden Then
                  col = col + 1
                     
'                    Stop
                    
                 'obj1.AjouterValeur Trim(Str(col)), Cells(i, c.Column).Value
                 
                 'Stop
                ' Debug.Print c.Column & "    " & i
                 
                 obj1.AjouterValeur Cells(1, c.Column), Cells(i, c.Column).Value
             '  Stop
                 
                ' End If
      
                 Next c
               '  Stop
               '  Debug.Print obj1.lireValeur("1")
                 
'   Stop
   
   'If dico_standards.exists(obj1.lireValeur("1")) Then Stop
'   Stop
   
   dico_standards.Add obj1.lireValeur("Référence"), obj1
   
'   Stop
   col = 0
 '  Stop
suite:
 
   Next i
   
erreur:
 
' Dim valeurs As Variant
'valeurs = obj1.listerdico

'MsgBox valeurs("0")

'MsgBox valeurs(0)   ' Premier item

'Stop

'MsgBox obj1.dico_question("un")
 
'obj1.AjouterValeur entete_Elaboration(0), "Alice"
 ' obj1.AjouterValeur entete_Elaboration(0), "Alice"
    'obj1.Add entete_Elaboration(0), "Alice"

'mes_qestions.Add "El1", obj1

'obj1.AjouterValeur "un", "fff"


'teste de recuperation des données


'Set tempo = mes_questions("El394")

'Stop

'MsgBox tempo.lireValeur("12")
'Stop

'MsgBox obj1.lireValeur("toto")
'mes_qestions.Add "El2", obj1

'obj1.AjouterValeur "de", "de"

'Stop
'MsgBox obj1.dico_question(entete_Elaboration(0))
 
 ' ? Déclenche Class_Initialize()
  '  Dim obj2 As New CMaClasse   ' ? Nouvelle instance, nouveau dico !
    
   ' obj1.AjouterValeur "nom", "Alice"
   ' obj2.AjouterValeur "nom", "Bob"
    
  '  obj1.AfficherContenu
  '  obj2.AfficherContenu
    
    ' Les dictionnaires sont indépendants !


End Sub

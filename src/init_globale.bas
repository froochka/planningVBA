Option Explicit

' cet objet permet de stocker lenom de la colonne dans la feuille standard
' cela permet a partir de la direction de retrouver le nom de la colonne contenant le standrd de direction
' obligé de faire cela car DAG est present 2 fois : DAG et DAG Mobilier scolaire et Equipement général
' cet objet permet de le stocker dans un dico
Public colonne_Standard_direction



'Cette donnée permet de choisir sur la page de garde les standards a prendre en compte
' ils sont modifiables par le user mais cela impacte les données déjà saisies

Public ma_direction_achat As String
Public mes_standards_univers As String
Public mes_standards_direction As String




' creation d'une variable question_en_cours, elle permet de passer d'une question a l'autre (suivant, precedent)
' soit par un num soit par son id
Public question_en_cours

'dico des standards qui sticke la classe cStandard
Public dico_standards



'dico stock clé = id question et valeur = instance de classe
Public mes_questions As Object

'Variales Globales
Public ma_cat As Ccategorie
Public compteur_question 'demo
Public entete_Elaboration


'dictionnaires au niveau global

Public dictNodes As Object ' Dictionary

Public mon_dico_dir  ' dictionnaire des directions pour page de garde demo
Public dico_cat As Object   ' objet dico des categories


'declaration des feuilles

Public ws_str_qu As Worksheet   ' SHEET_STRUCT_QUESTION

Public ws_categories As Worksheet   ' SHEET_CATEGORIES


'declarer toutes les feuilles en constante

Public Const SHEET_STRUCT_QUESTION As String = "structure_question"

Public Const SHEET_ELABORATION As String = "3-Elaboration.O"

Public Const SHEET_STANDARDS As String = "140-Standards"

Public Const SHEET_CATEGORIES As String = "VBA_categories"



Sub init_global()


Set colonne_Standard_direction = CreateObject("Scripting.Dictionary")
colonne_Standard_direction.Add "Standard_DAV", "DAV_Standard Etablissement (E) / Direction (D) / Univers (U)"
colonne_Standard_direction.Add "Standard_DAG Mobilier scolaire et Equipement général", "DAG_Mobilier_scolaire_et_Equipement_general_Standard Etablissement (E) / Direction (D) / Univers (U)"
colonne_Standard_direction.Add "Standard_DAG SOFI", "DAG_SOFI_Standard Etablissement (E) / Direction (D) / Univers (U)"
colonne_Standard_direction.Add "Standard_DAI", "DAI_Standard Etablissement (E) / Direction (D) / Univers (U)"
colonne_Standard_direction.Add "Standard_DS", "DS_Standard Etablissement (E) / Direction (D) / Univers (U)"
colonne_Standard_direction.Add "Standard_Non_Standardisé", "Non_standardisé_Standard Etablissement (E) / Direction (D) / Univers (U)"


question_en_cours = "El1"

'Stop
'entete des 16 colonnes sert de clé pour le dico
' Set dico_question = CreateObject("scripting.dictionary")
 
 entete_Elaboration = Array("Numero", "Nb", "Standard", "Cas", "Chapitre", _
                 "ModeReponse", "ChoixReponse", "Reponses", _
                 "QuestionConditionnante", "NonConcerne", "ConformeSpecificite", _
                 "Diff", "JustifierDerogation", "Commentaires", _
                 "CommentairesIntra", "OperateurStandard")
    
'dico des questions de Elaborations
Set mes_questions = CreateObject("scripting.dictionary")
 
 
Set dico_cat = CreateObject("scripting.dictionary")


'decalarer les feuilles en objet globaux

Set ws_str_qu = ThisWorkbook.Worksheets(SHEET_STRUCT_QUESTION)

Set ws_categories = ThisWorkbook.Worksheets(SHEET_CATEGORIES)

compteur_question = 1
'Stop
Call test


End Sub


'lecture des questions, des chap, themes, et ss themes

Sub init_questions()

'Stop
ws_str_qu.Activate


End Sub

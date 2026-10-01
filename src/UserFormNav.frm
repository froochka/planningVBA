' Déclaration API en haut du module
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


'*******************************
'*******************************
'*******************************
'TESTE *************************
'*******************************
'*******************************

Private Sub CommandButton1_Click()

Dim obj1, mon_dicotemp

Set mon_dicotemp = CreateObject("Scripting.Dictionary")
Set obj1 = mes_questions("El2")

mon_dicotemp = obj1.listerdico

For Each c In mon_dicotemp

MsgBox c

Next c

'Label20.Caption = tempo.lireValeur("1")



Dim ctrl As Control

For Each ctrl In Me.Controls
    If TypeName(ctrl) = "TabStrip" Then
        ' C'est un TabStrip
        Debug.Print "TabStrip trouvé : " & ctrl.Name
        
        ' Boucle sur les onglets de ce TabStrip
        Dim i As Integer
        For i = 0 To ctrl.Tabs.Count - 1
            Debug.Print "  Onglet " & i & " : " & ctrl.Tabs(i).Caption
        Next i
    End If
Next ctrl

'Stop




'If Frame2.Width = 0 Then
'Frame2.Width = 408
'Else
'Frame2.Width = 0
'End If

'For Each c In TreeView1.Nodes

'If c.key Like "ROOT|*SST_3|El*" Then
'MsgBox c.key
'End If


'Next c

'Dim i As Integer
Dim topPos As Long
topPos = Me.TabStrip1.Top


For i = 1 To 13
    Dim lbl As Object
    Set lbl = Me.Frame2.Controls.Add("Forms.Label.1", "encoche_Lbl_" & i, True)
    
    With lbl
        .Caption = ChrW(&H2713)  ' Coche ? (ou ChrW(&H2714) pour ?)
        .Font.Name = "Wingdings"
        .Font.Size = 25
        
       ' Alternance aléatoire entre vert et blanc
        If Rnd() < 0.5 Then
            .ForeColor = RGB(0, 0, 0)    ' Vert
        Else
            .ForeColor = RGB(255, 255, 255) ' Blanc
        End If
        
        .Left = Me.TabStrip1.Left + Me.TabStrip1.Width - 15
        .Top = topPos
        .Width = 20
        .Height = Me.TabStrip1.TabFixedHeight
        .TextAlign = 2  ' Centré
        .Visible = True
         ' Couleur de fond
    .BackColor = &H8000000D 'RGB(5, 22, 255)    ' Gris clair
    .BackStyle = 0                     ' 1 = Opaque (affiche la couleur)
    
    End With
    
   ' topPos = (topPos + Me.TabStrip1.Height / 13)
   
     topPos = (topPos + Me.TabStrip1.TabFixedHeight) + 2.45
     
     
    
    
    
Next i


End Sub

Private Sub CommandButton2_Click()
compteur_question = compteur_question + 1

End Sub

Private Sub Frame3_Click()
On Error Resume Next

ProgressBar1.Value = ProgressBar1.Value + 10

  Dim pourcent As Integer
    pourcent = ProgressBar1.Value
    
    ' Calculer la largeur (FrameProgress = conteneur de 300px)
    LabelProgress.Width = (pourcent / 100) * UserForm2.Width
    
    
    ' Changer la couleur
    If pourcent < 30 Then
        LabelProgress.BackColor = RGB(255, 0, 0)      ' Rouge
      
            
    ElseIf pourcent < 50 Then
        LabelProgress.BackColor = RGB(255, 165, 0)   ' Orange
    
     ElseIf pourcent < 70 Then
        LabelProgress.BackColor = RGB(0, 0, 255)   ' Orange
        
    Else
        LabelProgress.BackColor = RGB(0, 255, 0)     ' Vert
    End If
    
    LabelProgress.Caption = pourcent & "%"
    Me.Repaint
    

End Sub







'****************************************************
'****************************************************
'****************************************************
'selection d'un chapitre
' id stocké dans TAG
'****************************************************
'****************************************************
'****************************************************

Private Sub TabStrip1_Change_temp()
'MsgBox Me.TabStrip1.SelectedItem.Tag
Dim nb_item As Integer

nb_item = TabStrip1.Tabs.Count

'Stop
On Error Resume Next

Label_fleche.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / nb_item) - 5
Label_fleche2.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / nb_item) - 5


'Me.TabStrip1.Tabs(TabStrip1.SelectedItem.Index).Tag

'Stop

    Dim i As Integer
    
    ' Supprimer de la fin vers le début
   Do While TabStrip2.Tabs.Count > 0
    TabStrip2.Tabs.Remove 0
Loop

' Ajouter
Dim t As Object


For Each c In TreeView1.Nodes

If c.key Like Me.TabStrip1.SelectedItem.Tag & "|THEME_#" Or c.key Like Me.TabStrip1.SelectedItem.Tag & "|THEME_##" Then

Set t = TabStrip2.Tabs.Add
t.Caption = c
t.Tag = c.key

'TabStrip1.SelectedItem.Index

'Debug.Print "ok" & c.key
ElseIf c.key Like Me.TabStrip2(0).Tag & "|SST_#" Or c.key Like Me.TabStrip2(0).Tag & "|SST_##" Then

Set t = TabStrip3.Tabs.Add
t.Caption = c
t.Tag = c.key

'Debug.Print "ok" & c.key

End If


Next c






End Sub





Private Sub TabStrip1_Click(ByVal Index As Long)
'MsgBox ("clique")

'Call Liste_questions_Textbox
'Call MAJ_tabStrips(1)

Dim nb_item As Integer

nb_item = TabStrip1.Tabs.Count

'Stop
'On Error Resume Next

Label_fleche.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / nb_item) - 5
Label_fleche2.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / nb_item) - 5


Call MAJ_tabStrips2
Call MAJ_tabStrips3


End Sub

Private Sub MAJ_tabStrips2()
 ' Supprimer de la fin vers le début
   Do While TabStrip2.Tabs.Count > 0
    TabStrip2.Tabs.Remove 0
    Loop
    
    Dim t As Object

For Each c In TreeView1.Nodes

If c.key Like Me.TabStrip1.SelectedItem.Tag & "|THEME_#" Or c.key Like Me.TabStrip1.SelectedItem.Tag & "|THEME_##" Then

Set t = TabStrip2.Tabs.Add
t.Caption = c
t.Tag = c.key
End If


Next c

End Sub

Private Sub MAJ_tabStrips3()


  ' Supprimer de la fin vers le début
   Do While TabStrip3.Tabs.Count > 0
    TabStrip3.Tabs.Remove 0
    Loop
    
    Dim t As Object

For Each c In TreeView1.Nodes

If Me.TabStrip2.Tabs.Count > 0 Then
    If c.key Like Me.TabStrip2.SelectedItem.Tag & "|SST_#" Or c.key Like Me.TabStrip2.SelectedItem.Tag & "|SST_##" And c <> "ROOT" Then
'    Stop
    Set t = TabStrip3.Tabs.Add
    t.Caption = c
    t.Tag = c.key
    End If

End If
Next c

Call Liste_questions_Textbox



End Sub



Private Sub MAJ_tabStrips()

'MsgBox Me.TabStrip1.SelectedItem.Tag
Dim nb_item As Integer

nb_item = TabStrip1.Tabs.Count

'Stop
'On Error Resume Next

Label_fleche.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / nb_item) - 5
Label_fleche2.Top = TabStrip1.Top + (TabStrip1.SelectedItem.Index * TabStrip1.Height / nb_item) - 5


'Me.TabStrip1.Tabs(TabStrip1.SelectedItem.Index).Tag

'Stop

    Dim i As Integer
    
'If choix_strip = 1 Then
    ' Supprimer de la fin vers le début
   Do While TabStrip2.Tabs.Count > 0
    TabStrip2.Tabs.Remove 0
    Loop



   ' Supprimer de la fin vers le début
   Do While TabStrip3.Tabs.Count > 0
    TabStrip3.Tabs.Remove 0
    Loop

'End If

' Ajouter
Dim t As Object

For Each c In TreeView1.Nodes

If c.key Like Me.TabStrip1.SelectedItem.Tag & "|THEME_#" Or c.key Like Me.TabStrip1.SelectedItem.Tag & "|THEME_##" Then

Set t = TabStrip2.Tabs.Add
t.Caption = c
t.Tag = c.key

'TabStrip1.SelectedItem.Index

'Debug.Print "ok" & c.key

ElseIf Me.TabStrip2.Tabs.Count > 0 Then
    If c.key Like Me.TabStrip2(0).Tag & "|SST_#" Or c.key Like Me.TabStrip2(0).Tag & "|SST_##" And c <> "ROOT" Then
'    Stop
    Set t = TabStrip3.Tabs.Add
    t.Caption = c
    t.Tag = c.key
    End If

End If

Next c

'Stop
'MAJ des questions
Call Liste_questions_Textbox



End Sub


'****************************************************
'****************************************************
'****************************************************
'selection d'un THEME
' id stocké dans TAG
'****************************************************
'****************************************************
'****************************************************

Private Sub TabStrip2_Change2()
On Error Resume Next

Dim nb_item As Integer


'Stop


   Dim i As Integer
    
    ' Supprimer de la fin vers le début
   Do While TabStrip3.Tabs.Count > 0
    TabStrip3.Tabs.Remove 0
Loop

' Ajouter
Dim t As Object


'Stop
For Each c In TreeView1.Nodes

If c.key Like Me.TabStrip2.SelectedItem.Tag & "|SST_#" Or c.key Like Me.TabStrip2.SelectedItem.Tag & "|SST_##" Then

Set t = TabStrip3.Tabs.Add
t.Caption = c
t.Tag = c.key

'Debug.Print "ok" & c.key
End If


Next c

nb_item = TabStrip2.Tabs.Count
Debug.Print nb_item

'Stop

'TabStrip2.Height = nb_item * TabStrip2.TabFixedHeight
'Stop


'Label_fleche3.Top = TabStrip2.Top + (TabStrip2.SelectedItem.Index * TabStrip2.Height / nb_item) - 5 ' - (TabStrip2.TabFixedHeight / 2)
'Label_fleche4.Top = TabStrip2.Top + (TabStrip2.SelectedItem.Index * TabStrip2.Height / nb_item) - 5 '- (TabStrip2.TabFixedHeight / 2)



On Error GoTo 0



End Sub

Private Sub TabStrip2_Click(ByVal Index As Long)
Call MAJ_tabStrips3
'Call Liste_questions_Textbox
End Sub



'obsolète

Private Sub TabStrip3_Change_temp()
GoTo suite
On Error Resume Next

'Label_fleche5.Top = TabStrip3.Top + (TabStrip3.SelectedItem.Index * TabStrip3.Height / 12) - 5
'Label_fleche6.Top = TabStrip3.Top + (TabStrip3.SelectedItem.Index * TabStrip3.Height / 12) - 5



   Dim i As Integer
    
    ' Supprimer de la fin vers le début
   Do While TabStrip3.Tabs.Count > 0
    TabStrip3.Tabs.Remove 0
Loop

' Ajouter
Dim t As Object


'Stop
For Each c In TreeView1.Nodes

If c.key Like Me.TabStrip2.SelectedItem.Tag & "|SST_#" Or c.key Like Me.TabStrip2.SelectedItem.Tag & "|SST_##" Then

Set t = TabStrip3.Tabs.Add
t.Caption = c
t.Tag = c.key

'Debug.Print "ok" & c.key
End If


Next c

On Error GoTo 0

suite:


End Sub

Private Sub TabStrip3_Click(ByVal Index As Long)
Call Liste_questions_Textbox
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    ' ID = Node.Key
    MsgBox "ID du nœud : " & Node.key
    
    ' Ou pour l'utiliser ailleurs
    Debug.Print Node.key
End Sub





'Userform NAV

Private Sub UserForm_Initialize()
'Stop

UserForm_Question.Show





'*******************************
'*******************************
'*******************************
' ECRAN
'*******************************
'*******************************
'*******************************

'plein ecran

    ' Calculer le facteur d'échelle
 '   Dim screenWidth, screenHeight As Long
 '   screenWidth = GetSystemMetrics(0)
 '   screenHeight = GetSystemMetrics(1)
    
    
'UserFormNav.Width = screenWidth * 0.75 '- 10
'UserFormNav.Height = screenHeight * 0.75  '- 10
   




    ' Appliquer l'échelle à tout
  '  AppliquerEchelle
    
suite:
    

'Stop

'*******************************
'*******************************
'*******************************
'treeview
'*******************************
'*******************************
'*******************************

Set dictNodes = CreateObject("Scripting.Dictionary")
'Set TreeView1 = Me.TreeView1 ' adapter selon ton contrôle
AlimenterTreeview


'Stop
Dim mon_compt As Integer   ' compteur tempo
Dim temp As String ' variable tempo

'Stop

Call init_global
Call init_questions 'module de lecture des questions en memoire
'Call creer_categories ' creation du dico des categories* avec CLASS STRUCT_QUESTION

Call creer_categories_from_treeview
Me.TabStrip1.TabOrientation = fmTabOrientationLeft

mon_compt = 0
'******************************************************
'******************************************************
'******************************************************
'alimente les categories depuis le treview
'******************************************************
'******************************************************
'******************************************************
'Stop
For Each c In TreeView1.Nodes

If c.key Like "ROOT|CHAP_#" Or c.key Like "ROOT|CHAP_##" Then

Me.TabStrip1.Tabs(mon_compt).Caption = c
Me.TabStrip1.Tabs(mon_compt).Tag = c.key 'stockage de l'id du chapitre dans le TAG du TABS (de TABSTRIP)

mon_compt = mon_compt + 1

End If


Next c

Me.TabStrip1.Value = 1
Me.TabStrip1.Value = 0


'For Each c In dico_cat
'temp = dico_cat(c).ma_type

'If temp = "CHAP" Then
'Stop
'temp = dico_cat(c).ma_lib

'Me.TabStrip1.Tabs(mon_compt).Caption = dico_cat(c).ma_lib
'Me.TabStrip1.Tabs(mon_compt).Tag = c 'stockage de l'id du chapitre dans le TAG du TABS (de TABSTRIP)

'mon_compt = mon_compt + 1

'Debug.Print dico_cat(c).ma_lib & "__";
'End If


'Stop

'Next c

'Stop
Call MAJ_tabStrips


End Sub


'******************************************************************
'***** Affichage de la liste des questions dans un textbox
'*****
'******************************************************************
Private Sub Liste_questions_Textbox()

Dim fil_ariane As String
'Stop
'On Error Resume Next
UserFormNav.TextBox1.Text = "" 'UserFormNav.TextBox1.Text & "sdfksldfjsdlkfjsdlfksdjflskdfj" & vbCrLf
UserForm_Question.ListBox1.Clear
UserForm_Question.ListBox1.Tag = ""

'Stop
For Each c In TreeView1.Nodes
'Stop
If c.key Like Me.TabStrip3.SelectedItem.Tag & "|El*" Then
'Stop

'UserFormNav.TextBox1.Text = UserFormNav.TextBox1.Text & Split(c, "-")(1) & vbCrLf
UserFormNav.TextBox1.Text = UserFormNav.TextBox1.Text & Mid(c, InStr(1, c, "-"), Len(c)) & vbCrLf
'UserFormNav.TextBox1.Text = UserFormNav.TextBox1.Text & c & vbCrLf
'UserForm_Question.ListBox1.AddItem Mid(c, InStr(1, c, "-"), Len(c))

'UserForm_Question.ListBox1.AddItem c.key    'Mid(c, InStr(1, c, "-"), Len(c))
'UserForm_Question.ListBox1.AddItem c    'Mid(c, InStr(1, c, "-"), Len(c))
'Stop

'UserForm_Question.ListBox1.AddItem c   'Mid(c, InStr(1, c, "-"), Len(c))
'UserForm_Question.ListBox1.List(ListBox1.ListCount - 1, 0) = c

' Au lieu de AddItem c (une seule colonne)
UserForm_Question.ListBox1.AddItem
UserForm_Question.ListBox1.List(UserForm_Question.ListBox1.ListCount - 1, 1) = ChrW(9675)  ' ? statut
UserForm_Question.ListBox1.List(UserForm_Question.ListBox1.ListCount - 1, 0) = c

UserForm_Question.ListBox1.Tag = UserForm_Question.ListBox1.Tag & c.key & "-"



'Debug.Print "ok" & c.key
End If

Next c

fil_ariane = Me.TabStrip1.SelectedItem.Caption & vbCr & _
            "   " & Me.TabStrip2.SelectedItem.Caption & vbCr & _
             "        " & Me.TabStrip3.SelectedItem.Caption & vbCr

UserForm_Question.Label_filAriane.Caption = fil_ariane

UserForm_Question.ListBox1.ListIndex = 0



End Sub

'******************************************************************
'***** Affichage de la liste des questions dans un bouton

'*****
'******************************************************************

Private Sub Liste_questions_Textbox2()
    'On Error Resume Next
    UserFormNav.TextBox1.Text = ""
    
    Dim lbl As Object
    Dim topPos As Long
    topPos = 10
    
    ' Nettoyer les anciens contrôles si nécessaire
    Dim ctrl As Object
    For Each ctrl In Me.Frame7.Controls
        If TypeName(ctrl) = "Label" Then Me.Frame7.Controls.Remove ctrl.Name
    Next ctrl
    
    For Each c In TreeView1.Nodes
        If c.key Like Me.TabStrip3.SelectedItem.Tag & "|El*" Then
            Set lbl = Me.Frame7.Controls.Add("Forms.Label.1", "Lbl_" & c.key, True)
            
            With lbl
                ' --- TEXTE ---
                .Caption = Mid(c, InStr(1, c, "-"), Len(c))
                .Tag = c.key
                
                ' --- POSITION & TAILLE ---
                .Left = 10
                .Top = topPos
                .Width = Me.Frame7.Width - 20
                .Height = 29
'                Stop
                ' --- COULEUR DU TEXTE ---
                '.ForeColor = RGB(0, 0, 0)          ' Noir (RGB: Rouge, Vert, Bleu)
                 .ForeColor = RGB(255, 255, 255)          ' blanc
                ' .ForeColor = RGB(255, 0, 0)       ' Rouge
                ' .ForeColor = RGB(0, 0, 255)       ' Bleu
                ' .ForeColor = RGB(0, 128, 0)       ' Vert foncé
                ' .ForeColor = &HFF00&              ' Vert clair
                ' .ForeColor = &HFF&                ' Rouge
                ' .ForeColor = &HFF0000             ' Bleu
                ' .ForeColor = &H80000008           ' Couleur système (texte fenêtre)
                
                ' --- COULEUR DE FOND ---
               ' .BackColor = RGB(240, 240, 240)    ' Gris clair
                ' .BackColor = RGB(255, 255, 255)   ' Blanc
                 '.BackColor = RGB(255, 255, 0)     ' Jaune
                 .BackColor = RGB(0, 23, 255)     ' Bleu très clair
                ' .BackColor = &H80000005           ' Couleur système (fond fenêtre)
                .BackStyle = 1                      ' 1 = Opaque (affiche la couleur), 0 = Transparent
                
                ' --- BORDURE ---
                '.BorderStyle = 1                   ' 1 = Bordure simple, 0 = Sans bordure
                '.BorderColor = RGB(0, 0, 0)        ' Couleur de la bordure
                ' .BorderColor = RGB(128, 128, 128) ' Gris
                ' .BorderColor = RGB(0, 0, 255)     ' Bleu
                
                ' --- POLICE (FONT) ---
                .Font.Name = "Arial"               ' Nom de la police
                ' .Font.Name = "Times New Roman"
                ' .Font.Name = "Calibri"
                ' .Font.Name = "Courier New"       ' Police à chasse fixe
                
                .Font.Size = 13                    ' Taille de la police
                ' .Font.Size = 14
                ' .Font.Size = 8
                
                .Font.Bold = False                 ' True = Gras, False = Normal
                .Font.Italic = False               ' True = Italique, False = Normal
                .Font.Underline = False            ' True = Souligné, False = Normal
                .Font.Strikethrough = False        ' True = Barré, False = Normal
                
                ' --- ALIGNEMENT DU TEXTE ---
                '.TextAlign = 0                     ' 0 = Gauche, 1 = Droite, 2 = Centre
                ' .TextAlign = 2
                
                ' --- AUTRES PROPRIÉTÉS ---
                .AutoSize = False                  ' True = Ajuste automatiquement la taille au texte
                .WordWrap = True                   ' True = Retour à la ligne automatique
                .Enabled = True                    ' True = Actif, False = Désactivé (grisé)
                .Visible = True                    ' True = Visible, False = Caché
                .SpecialEffect = 0                ' 0 = Plat, 1 = Enfoncé, 2 = Relief, 3 = Bombé
                ' .SpecialEffect = 1
                
                ' --- CURSEUR ---
                .MousePointer = 0                  ' 0 = Par défaut, 1 = Flèche, 2 = Croix, etc.
                ' .MousePointer = 2                ' Croix
                ' .MousePointer = 11              ' Main (comme un lien)
                
                ' --- ASTUCE (TOOLTIP) ---
                .ControlTipText = "Cliquez pour sélectionner : " & .Caption
            End With
            
            topPos = topPos + 30
        End If
    Next c
End Sub

Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Code à exécuter quand on clique sur le label
    MsgBox "Label cliqué !"
End Sub

Private Sub AlimenterTreeview()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Dim ref As String
    Dim libelle As String

    Dim chapID As String
    Dim themeID As String
    Dim sstID As String

    Dim chapKey As String
    Dim themeKey As String
    Dim sstKey As String
    Dim questionKey As String

    Dim dictNodes As Object
    Dim dictLibelles As Object

    Set ws = ThisWorkbook.Sheets(SHEET_STRUCT_QUESTION)

    Set dictNodes = CreateObject("Scripting.Dictionary")
    Set dictLibelles = CreateObject("Scripting.Dictionary")

    TreeView1.Nodes.Clear

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    '=================================================
    ' PASSAGE 1 : Charger les libellés
    '=================================================
    For i = 2 To lastRow

        ref = Trim(ws.Cells(i, 1).Value)
        libelle = Trim(ws.Cells(i, 2).Value)

        If ref <> "" Then

            If Left(ref, 5) = "CHAP_" _
            Or Left(ref, 6) = "THEME_" _
            Or Left(ref, 4) = "SST_" Then
            dictLibelles(ref) = libelle

            End If

        End If

    Next i
'Stop
    '=================================================
    ' ROOT
    '=================================================
    TreeView1.Nodes.Add , , "ROOT", "ROOT"

    dictNodes.Add "ROOT", TreeView1.Nodes("ROOT")

    '=================================================
    ' PASSAGE 2 : Construction de l'arbre
    '=================================================
    For i = 2 To lastRow
'Stop
        ref = Trim(ws.Cells(i, 1).Value)

        If Left(ref, 2) = "El" Then

            libelle = Trim(ws.Cells(i, 2).Value)

            chapID = Trim(ws.Cells(i, 3).Value)
            themeID = Trim(ws.Cells(i, 4).Value)
            sstID = Trim(ws.Cells(i, 5).Value)

            '=========================
            ' Clés hiérarchiques
            '=========================
            chapKey = "ROOT|" & chapID

            themeKey = chapKey & "|" & themeID

            sstKey = themeKey & "|" & sstID

            questionKey = sstKey & "|" & ref

            '=========================
            ' Chapitre
            '=========================
            If Not dictNodes.exists(chapKey) Then

                TreeView1.Nodes.Add _
                    "ROOT", _
                    tvwChild, _
                    chapKey, _
                    dictLibelles(chapID)

                dictNodes.Add chapKey, TreeView1.Nodes(chapKey)

            End If

            '=========================
            ' Thème
            '=========================
            If Not dictNodes.exists(themeKey) Then

                TreeView1.Nodes.Add _
                    chapKey, _
                    tvwChild, _
                    themeKey, _
                    dictLibelles(themeID)

                dictNodes.Add themeKey, TreeView1.Nodes(themeKey)

            End If

            '=========================
            ' Sous-thème
            '=========================
            If Not dictNodes.exists(sstKey) Then

                TreeView1.Nodes.Add _
                    themeKey, _
                    tvwChild, _
                    sstKey, _
                    dictLibelles(sstID)

                dictNodes.Add sstKey, TreeView1.Nodes(sstKey)

            End If

            '=========================
            ' Question
            '=========================
            If Not dictNodes.exists(questionKey) Then

                TreeView1.Nodes.Add _
                    sstKey, _
                    tvwChild, _
                    questionKey, _
                    ref & " - " & libelle

                dictNodes.Add questionKey, True

            End If

        End If

    Next i
'Stop
    '=================================================
    ' Déplier ROOT
    '=================================================
    TreeView1.Nodes("ROOT").Expanded = True

    Dim nd As Node

    For Each nd In TreeView1.Nodes

        If Not nd.Parent Is Nothing Then

            If nd.Parent.key = "ROOT" Then
                nd.Expanded = True
            End If

        End If
    Next nd
'Stop
End Sub

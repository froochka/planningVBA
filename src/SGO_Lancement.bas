Option Explicit

' =============================================================================
' Points d'entrée de l'application SGO (itération 6 : vraies données)
'
'   Alt+F8 > SGO                   : créer ou ouvrir un dossier d'offre, puis saisir
'   Alt+F8 > ImporterReferentiel   : (re)charger referentiel\*.csv dans le classeur
'   Alt+F8 > TesterMoteur          : contrôle du moteur par rapport à l'outil d'origine
'   Alt+F8 > CreerFormulairesSGO   : crée les UserForms vides frmGarde et frmSaisie
'                                    (une seule fois, puis relancer ImporterVBA)
'
' Le dossier d'une offre est un fichier .xlsx séparé, sans macro (voir SGO_Dossier).
' =============================================================================

' Échange avec la page de garde (frmGarde)
Public gGarde As Object              ' champ -> valeur (voir CHAMPS_OFFRE) + "conserver"
Public gGardeValidee As Boolean
Public gGardeCreation As Boolean     ' True : nouveau dossier (pas de choix « Conserver »)

' Réouverture de l'écran de saisie (taille du texte ou page de garde modifiée)
Public gRouvrirSaisie As Boolean
Public gQuestionDepart As cQuestionSGO

Public Sub SGO()
    If Not PreparerReferentiel() Then Exit Sub
    If DossierOuvert() Then
        If MsgBox("Reprendre le dossier ouvert « " & gWbDossier.Name & " » ?", vbYesNo + vbQuestion, "SGO") = vbYes Then
            If InitialiserSession() Then AfficherSaisie
            Exit Sub
        End If
    End If
    Select Case MsgBox("Oui : créer un nouveau dossier d'offre" & vbCrLf & _
                       "Non : ouvrir un dossier existant (.xlsx)", vbYesNoCancel + vbQuestion, "SGO")
        Case vbYes: NouveauDossier
        Case vbNo: OuvrirDossierExistant
    End Select
End Sub

Private Function PreparerReferentiel() As Boolean
    If Not ReferentielPresent() Then
        MsgBox "Le référentiel n'est pas encore dans le classeur." & vbCrLf & _
               "Lancez d'abord la macro ImporterReferentiel.", vbExclamation, "SGO"
        Exit Function
    End If
    On Error GoTo erreur
    ChargerReferentiel
    PreparerReferentiel = True
    Exit Function
erreur:
    MsgBox "Lecture du référentiel impossible : " & Err.Description & vbCrLf & _
           "Relancez la macro ImporterReferentiel.", vbExclamation, "SGO"
End Function

Private Sub NouveauDossier()
    Dim chemin As Variant, nom As String
    Set gGarde = NouveauDico()
    gGardeCreation = True
    gGardeValidee = False
    If Not AfficherForm("frmGarde") Then Exit Sub
    If Not gGardeValidee Then Exit Sub

    nom = "Offre_" & NomFichier(gGarde("numero_procedure")) & ".xlsx"
    If nom = "Offre_.xlsx" Then nom = "Offre_" & Format$(Date, "yyyy-mm-dd") & ".xlsx"
    chemin = Application.GetSaveAsFilename(InitialFileName:=nom, _
                                           FileFilter:="Dossier SGO (*.xlsx), *.xlsx", Title:="Enregistrer le dossier de l'offre")
    If VarType(chemin) = vbBoolean Then Exit Sub
    If Dir(CStr(chemin)) <> "" Then
        MsgBox "Ce fichier existe déjà. Choisissez un autre nom, ou ouvrez-le avec « Non » au lancement.", _
               vbExclamation, "SGO"
        Exit Sub
    End If
    If Not CreerDossier(CStr(chemin), gGarde) Then Exit Sub
    If InitialiserSession() Then AfficherSaisie
End Sub

Private Sub OuvrirDossierExistant()
    Dim chemin As Variant
    chemin = Application.GetOpenFilename(FileFilter:="Dossier SGO (*.xlsx), *.xlsx", Title:="Ouvrir un dossier d'offre")
    If VarType(chemin) = vbBoolean Then Exit Sub
    If Not OuvrirDossier(CStr(chemin)) Then Exit Sub
    If InitialiserSession() Then AfficherSaisie
End Sub

' Référentiel neuf + réponses du dossier + évaluation
Public Function InitialiserSession() As Boolean
    Dim inconnues As Long, msg As String
    On Error GoTo erreur
    ChargerReferentiel
    gDirection = DossierValeur("standard_direction")
    gUniversChoisi = DossierValeur("standard_univers")
    inconnues = ChargerReponses()
    EvaluerDossier
    If DossierValeur("version_referentiel") <> gVersionRef Then
        msg = "Ce dossier a été rempli avec la version " & DossierValeur("version_referentiel") & _
              " du référentiel ; la version installée est la " & gVersionRef & "."
        DossierEcrire "version_referentiel", gVersionRef
    End If
    If inconnues > 0 Then
        msg = msg & IIf(msg = "", "", vbCrLf) & inconnues & _
              " réponse(s) concernent des questions absentes du référentiel : elles restent dans le fichier."
    End If
    If gRapportOuverture <> "" Then
        msg = msg & IIf(msg = "", "", vbCrLf & vbCrLf) & "Contrôle du dossier à l'ouverture :" & gRapportOuverture
    End If
    gRapportOuverture = ""
    If msg <> "" Then MsgBox msg, vbInformation, "SGO"
    InitialiserSession = True
    Exit Function
erreur:
    MsgBox "Chargement du dossier impossible : " & Err.Description, vbExclamation, "SGO"
End Function

' -----------------------------------------------------------------------------
' Modification de la page de garde depuis l'écran de saisie (cadrage P6, PG-5)
' Renvoie True si le contexte a changé (l'écran doit être rafraîchi).
' -----------------------------------------------------------------------------
Public Function ModifierPageGarde() As Boolean
    Dim k As Variant, ancDir As String, ancUniv As String, nouvDir As String, nouvUniv As String
    Dim q As cQuestionSGO, n As Long, nSaisies As Long, liste As String, conserver As Boolean

    Set gGarde = ValeursOffre()
    gGarde("conserver") = "Oui"
    gGardeCreation = False
    gGardeValidee = False
    If Not AfficherForm("frmGarde") Then Exit Function
    If Not gGardeValidee Then Exit Function

    ancDir = gDirection: ancUniv = gUniversChoisi
    nouvDir = gGarde("standard_direction"): nouvUniv = gGarde("standard_univers")
    conserver = (gGarde("conserver") = "Oui")

    If StrComp(ancDir, nouvDir, vbTextCompare) <> 0 Or StrComp(ancUniv, nouvUniv, vbTextCompare) <> 0 Then
        ' Analyse d'impact avant d'appliquer
        For Each k In gQuestions.Keys
            Set q = gQuestions(k)
            If Not q.Systeme Then
                If ValeurStd(q.Id, nouvDir, nouvUniv, NiveauStd(q.Id, nouvDir, nouvUniv)) <> q.Standard Then
                    n = n + 1
                    If q.Saisie Then
                        nSaisies = nSaisies + 1
                        If nSaisies <= 12 Then liste = liste & IIf(liste = "", "", ", ") & q.Id
                    End If
                End If
            End If
        Next k
        If MsgBox(n & " question(s) ont un standard différent avec la nouvelle direction ou le nouvel univers," & _
                  vbCrLf & "dont " & nSaisies & " déjà saisie(s)" & IIf(liste = "", "", " (" & liste & IIf(nSaisies > 12, "…", "") & ")") & "." & _
                  vbCrLf & vbCrLf & IIf(conserver, _
                    "Les réponses saisies seront conservées ; celles dont le standard change seront marquées « Standard modifié ».", _
                    "Toutes les réponses seront archivées (feuille Archives) et les questions repartiront du standard.") & _
                  vbCrLf & vbCrLf & "Appliquer le changement ?", vbOKCancel + vbQuestion, "SGO - Page de garde") = vbCancel Then
            Exit Function
        End If
        If conserver Then
            For Each k In gQuestions.Keys
                Set q = gQuestions(k)
                If q.Saisie And Not q.Systeme Then
                    If ValeurStd(q.Id, nouvDir, nouvUniv, NiveauStd(q.Id, nouvDir, nouvUniv)) <> q.Standard Then
                        q.StandardModifie = True
                        EnregistrerQuestion q
                    End If
                End If
            Next k
        Else
            ArchiverReponses "Page de garde : " & ancDir & " / " & ancUniv & " -> " & nouvDir & " / " & nouvUniv
        End If
        ModifierPageGarde = True
    End If

    For Each k In Split(CHAMPS_OFFRE, ";")
        If gGarde.Exists(k) And k <> "version_referentiel" And k <> "cree_le" And k <> "modifie_le" And k <> "modifie_par" Then
            DossierEcrire CStr(k), CStr(gGarde(k))
        End If
    Next k
    gDirection = nouvDir: gUniversChoisi = nouvUniv
    EvaluerDossier
    SauverDossier
    ModifierPageGarde = True
End Function

' Écran de saisie ; rouvert tant qu'il le demande (gRouvrirSaisie)
Public Sub AfficherSaisie()
    Do
        gRouvrirSaisie = False
        If Not AfficherForm("frmSaisie") Then Exit Sub
    Loop While gRouvrirSaisie
    Set gQuestionDepart = Nothing
End Sub

' Ouvre un formulaire par son nom : le module se compile même si le formulaire
' n'a pas encore été créé dans le classeur
Public Function AfficherForm(ByVal nom As String) As Boolean
    Dim f As Object
    On Error GoTo erreur
    Set f = VBA.UserForms.Add(nom)
    f.Show
    AfficherForm = True
    Exit Function
erreur:
    MsgBox "Impossible d'ouvrir " & nom & " : " & Err.Description & vbCrLf & vbCrLf & _
           "Si ce formulaire n'existe pas encore, lancez CreerFormulairesSGO puis ImporterVBA.", _
           vbExclamation, "SGO"
End Function

Private Function NomFichier(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "-")
    Next c
    NomFichier = Trim$(s)
End Function

' -----------------------------------------------------------------------------
' Crée les UserForms vides frmGarde et frmSaisie (leur contenu est construit par
' leur code). Nécessite « Accès approuvé au modèle d'objet du projet VBA ».
' -----------------------------------------------------------------------------
Public Sub CreerFormulairesSGO()
    Dim nom As Variant, comp As Object, crees As String
    For Each nom In Array("frmGarde", "frmSaisie")
        Set comp = Nothing
        On Error Resume Next
        Set comp = ThisWorkbook.VBProject.VBComponents(CStr(nom))
        On Error GoTo erreur
        If comp Is Nothing Then
            Set comp = ThisWorkbook.VBProject.VBComponents.Add(3)     ' vbext_ct_MSForm
            comp.Name = CStr(nom)
            On Error Resume Next
            comp.Properties("StartUpPosition") = 0                  ' position calculée par le code
            On Error GoTo erreur
            crees = crees & vbCrLf & "- " & nom
        End If
    Next nom
    If crees = "" Then
        MsgBox "Les formulaires existent déjà.", vbInformation, "SGO"
    Else
        MsgBox "Formulaires créés :" & crees & vbCrLf & vbCrLf & _
               "Relancez maintenant ImporterVBA pour y mettre leur code, puis enregistrez le classeur.", _
               vbInformation, "SGO"
    End If
    Exit Sub
erreur:
    MsgBox "Création impossible : " & Err.Description & vbCrLf & _
           "Vérifiez l'option « Accès approuvé au modèle d'objet du projet VBA ».", vbExclamation, "SGO"
End Sub

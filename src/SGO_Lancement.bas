Option Explicit

' =============================================================================
' Points d'entrée de l'application SGO (itération 6 : vraies données)
'
'   Alt+F8 > SGO                   : créer ou ouvrir un DCO, puis saisir
'   Alt+F8 > ImporterReferentiel   : (re)charger referentiel\*.csv dans le classeur
'   Alt+F8 > TesterMoteur          : contrôle du moteur par rapport à l'outil d'origine
'   Alt+F8 > CreerFormulairesSGO   : crée les UserForms vides frmGarde et frmSaisie
'                                    (une seule fois, puis relancer ImporterVBA)
'
' Le DCO d'une offre est un fichier .xlsx séparé, sans macro (voir SGO_Dossier).
' L'écran de saisie n'est pas bloquant : Excel reste utilisable pendant la
' saisie (feuille Remarques, autres classeurs). Le DCO est alors masqué, et un
' seul écran de saisie peut être ouvert à la fois (SaisieDejaOuverte).
' =============================================================================

' Échange avec la page de garde (frmGarde)
Public gGarde As Object              ' champ -> valeur (voir CHAMPS_OFFRE) + "conserver"
Public gGardeValidee As Boolean
Public gGardeCreation As Boolean     ' True : nouveau dossier (pas de choix « Conserver »)

' Écran de saisie : question à afficher à la (ré)ouverture, mode réduit, et
' numéro du dernier écran ouvert (lors d'une réouverture, le nouvel écran est
' affiché avant la fermeture de l'ancien : seul le dernier rend la main)
Public gQuestionDepart As cQuestionSGO
Public gSaisieReduite As Boolean
Public gNumSaisie As Long

Public Sub SGO()
    If SaisieDejaOuverte() Then Exit Sub
    If Not PreparerReferentiel() Then Exit Sub
    If DossierOuvert() Then
        If MsgBox("Reprendre le DCO ouvert « " & gWbDossier.Name & " » ?", vbYesNo + vbQuestion, "SGO") = vbYes Then
            If InitialiserSession() Then AfficherSaisie
            Exit Sub
        End If
    End If
    Select Case MsgBox("Oui : créer un nouveau DCO" & vbCrLf & _
                       "Non : ouvrir un DCO existant (.xlsx)", vbYesNoCancel + vbQuestion, "SGO")
        Case vbYes: NouveauDossier
        Case vbNo: OuvrirDossierExistant
    End Select
End Sub

' Boutons de la page « Démarrer » (version de test)
Public Sub SGONouveauDossier()
    If SaisieDejaOuverte() Then Exit Sub
    If PreparerReferentiel() Then NouveauDossier
End Sub

Public Sub SGOOuvrirDossier()
    If SaisieDejaOuverte() Then Exit Sub
    If PreparerReferentiel() Then OuvrirDossierExistant
End Sub

' Un seul écran de saisie à la fois (il utilise le référentiel et le DCO
' chargés) : s'il est déjà ouvert, on le ramène au premier plan.
Public Function SaisieDejaOuverte() As Boolean
    Dim f As Object, ouverte As Object
    For Each f In VBA.UserForms
        If f.Name = "frmSaisie" Then Set ouverte = f
    Next f
    If ouverte Is Nothing Then Exit Function
    SaisieDejaOuverte = True
    MsgBox "L'écran de saisie du DCO est déjà ouvert" & IIf(DossierOuvert(), " (" & gWbDossier.Name & ")", "") & "." & vbCrLf & _
           "Pour ouvrir ou créer un autre DCO, fermez d'abord cet écran.", vbInformation, "SGO"
    On Error Resume Next
    ouverte.Show vbModeless
End Function

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

    nom = "DCO_AO_" & NomFichier(gGarde("numero_procedure")) & ".xlsx"
    If nom = "DCO_AO_.xlsx" Then nom = "DCO_AO_" & Format$(Date, "yyyy-mm-dd") & ".xlsx"
    chemin = Application.GetSaveAsFilename(InitialFileName:=nom, _
                                           FileFilter:="DCO (*.xlsx), *.xlsx", Title:="Enregistrer le DCO")
    If VarType(chemin) = vbBoolean Then Exit Sub
    If Dir(CStr(chemin)) <> "" Then
        MsgBox "Ce fichier existe déjà. Choisissez un autre nom, ou ouvrez-le avec « Ouvrir un DCO existant ».", _
               vbExclamation, "SGO"
        Exit Sub
    End If
    If Not CreerDossier(CStr(chemin), gGarde) Then Exit Sub
    If InitialiserSession() Then AfficherSaisie
End Sub

Private Sub OuvrirDossierExistant()
    Dim chemin As Variant
    chemin = Application.GetOpenFilename(FileFilter:="DCO (*.xlsx), *.xlsx", Title:="Ouvrir un DCO")
    If VarType(chemin) = vbBoolean Then Exit Sub
    If Not OuvrirDossier(CStr(chemin)) Then Exit Sub
    If InitialiserSession() Then AfficherSaisie
End Sub

' Référentiel neuf + réponses du DCO + évaluation
Public Function InitialiserSession() As Boolean
    Dim inconnues As Long, msg As String
    On Error GoTo erreur
    ChargerReferentiel
    gDirection = DossierValeur("standard_direction")
    gUniversChoisi = DossierValeur("standard_univers")
    inconnues = ChargerReponses()
    EvaluerDossier
    If DossierValeur("version_referentiel") <> gVersionRef Then
        msg = "Ce DCO a été rempli avec la version " & DossierValeur("version_referentiel") & _
              " du référentiel ; la version installée est la " & gVersionRef & "."
        DossierEcrire "version_referentiel", gVersionRef
    End If
    If inconnues > 0 Then
        msg = msg & IIf(msg = "", "", vbCrLf) & inconnues & _
              " réponse(s) concernent des questions absentes du référentiel : elles restent dans le fichier."
    End If
    If gRapportOuverture <> "" Then
        msg = msg & IIf(msg = "", "", vbCrLf & vbCrLf) & "Contrôle du DCO à l'ouverture :" & gRapportOuverture
    End If
    gRapportOuverture = ""
    If msg <> "" Then MsgBox msg, vbInformation, "SGO"
    InitialiserSession = True
    Exit Function
erreur:
    MsgBox "Chargement du DCO impossible : " & Err.Description, vbExclamation, "SGO"
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
                    "Les réponses saisies seront conservées : vérifiez celles dont le standard change (elles peuvent devenir des dérogations au standard).", _
                    "Toutes les réponses seront archivées (feuille Archives) et les questions repartiront du standard.") & _
                  vbCrLf & vbCrLf & "Appliquer le changement ?", vbOKCancel + vbQuestion, "SGO - Page de garde") = vbCancel Then
            Exit Function
        End If
        ' Réponses conservées : la conformité est recalculée avec le nouveau standard
        ' (le marquage « Standard modifié » a été supprimé le 09/11/2026)
        If Not conserver Then
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

' Écran de saisie, non bloquant : Excel reste accessible pendant la saisie et
' le DCO est masqué (SGO_Dossier.MasquerDossier). Aussi appelé par l'écran
' lui-même pour se rouvrir (taille du texte, page de garde, réduire / agrandir) :
' le nouvel écran est affiché avant la fermeture de l'ancien (pas d'écran vide).
Public Function AfficherSaisie() As Boolean
    Dim f As Object
    On Error GoTo erreur
    Set f = VBA.UserForms.Add("frmSaisie")
    f.Show vbModeless
    MasquerDossier True
    AfficherSaisie = True
    Exit Function
erreur:
    MsgBox "Impossible d'ouvrir l'écran de saisie : " & Err.Description & vbCrLf & vbCrLf & _
           "Si ce formulaire n'existe pas encore, lancez CreerFormulairesSGO puis ImporterVBA.", _
           vbExclamation, "SGO"
End Function

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

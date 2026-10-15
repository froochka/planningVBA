Option Explicit

' =============================================================================
' Tests du moteur, à lancer dans Excel : Alt+F8 > TesterMoteur
'
' 1. Standards et conformité : rejoue referentiel\tests\attendus_standards.csv,
'    calculé avec les formules de l'outil d'origine (7 combinaisons direction /
'    univers x 413 questions), et compare au moteur VBA.
' 2. Masquage : compare la visibilité de chaque question à celle calculée par le
'    moteur Python (colonne « masquage » du même fichier).
' 3. Conditions : scénarios mère / fille (El77 / El78a, chapitre 05 / El445).
' Les écarts sont listés dans la feuille T_Resultats.
' =============================================================================

Private nbTests As Long, nbEcarts As Long, wsRes As Worksheet

Public Sub TesterMoteur()
    Dim t As Variant, col As Object, i As Long, j As Long, niv As String, std As String
    Dim rep As String, conf As String, q As cQuestionSGO, d As String, u As String
    Dim chemin As String, debut As Single, dernier As String, k As Variant

    debut = Timer
    If Not ReferentielPresent() Then
        MsgBox "Lancez d'abord ImporterReferentiel.", vbExclamation, "SGO"
        Exit Sub
    End If
    chemin = ThisWorkbook.Path & "\referentiel\tests\attendus_standards.csv"
    If Dir(chemin) = "" Then
        MsgBox "Fichier introuvable : " & chemin, vbExclamation, "SGO"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set wsRes = FeuilleTechnique(ThisWorkbook, "T_Resultats")
    wsRes.Visible = xlSheetVisible
    wsRes.Cells.Clear
    wsRes.Range("A1:F1").Value = Array("test", "direction / univers", "question", "attendu", "obtenu", "détail")
    nbTests = 0: nbEcarts = 0

    ChargerReferentiel
    t = LireCsv(chemin)
    Set col = NouveauDico()
    For j = 1 To UBound(t, 2)
        col(CStr(t(1, j))) = j
    Next j

    ' --- 1 et 2 : standards, conformité, masquage -----------------------------------
    For i = 2 To UBound(t, 1)
        d = t(i, col("direction")): u = t(i, col("univers"))
        If d & "|" & u <> dernier Then
            ' Nouveau contexte : dossier vierge (réponses pré-remplies avec le standard)
            gDirection = d: gUniversChoisi = u
            For Each k In gQuestions.Keys
                gQuestions(k).Saisie = False: gQuestions(k).Difference = "": gQuestions(k).Justification = ""
            Next k
            EvaluerDossier
            dernier = d & "|" & u
        End If
        Set q = gQuestions(CStr(t(i, col("question"))))
        niv = NiveauStd(q.Id, d, u)
        std = ValeurStd(q.Id, d, u, niv)
        rep = ReponseInitiale(niv, std, q.Mode)
        conf = CalculerConformite(niv, std, rep, q.Mode, q.Operateur)
        Verifier "niveau", d & " / " & u, q.Id, LCase$(t(i, col("niveau"))), LCase$(niv)
        Verifier "standard", d & " / " & u, q.Id, t(i, col("standard")), std
        Verifier "réponse pré-remplie", d & " / " & u, q.Id, t(i, col("reponse_initiale")), rep
        Verifier "conformité", d & " / " & u, q.Id, t(i, col("conformite")), conf
        Verifier "masquage", d & " / " & u, q.Id, t(i, col("masquage")), q.Motif
    Next i

    ' --- 3 : scénarios de conditions ----------------------------------------------------
    gDirection = "DAI": gUniversChoisi = "Logiciels"
    For Each k In gQuestions.Keys
        gQuestions(k).Saisie = False
    Next k
    EvaluerDossier
    Verifier "condition", "El77 = standard", "El78a", NC_OFFRE, gQuestions("El78a").Motif
    gQuestions("El77").Saisie = True: gQuestions("El77").ReponseSaisie = "Offre non exécutée"
    EvaluerDossier
    Verifier "condition", "El77 = Offre non exécutée", "El78a", "", gQuestions("El78a").Motif
    Verifier "condition", "El77 = Offre non exécutée", "ElTTT417 (sous-thème)", "False", IIf(gCategories("ElTTT417").Visible, "True", "False")
    gQuestions("El77").Saisie = False
    EvaluerDossier
    Verifier "condition", "El77 remis au standard", "El78a", NC_OFFRE, gQuestions("El78a").Motif

    ' Mise en sommeil : une réponse saisie dans une fille masquée est conservée
    gQuestions("El78a").Saisie = True: gQuestions("El78a").ReponseSaisie = "Réponse de test"
    EvaluerDossier
    Verifier "mise en sommeil", "El78a masquée", "El78a", "Réponse de test", gQuestions("El78a").ReponseSaisie
    Verifier "mise en sommeil", "El78a masquée", "El78a statut", CStr(ST_MASQUEE), CStr(gQuestions("El78a").Statut)
    gQuestions("El78a").Saisie = False

    ' Conformité : les 4 cas
    Verifier "conformité", "égal, casse différente", "-", CONF_CONFORME, CalculerConformite("E", "Oui", "oui", "Liste choix unique", "=")
    Verifier "conformité", "contient", "-", CONF_CONFORME, CalculerConformite("E", "Oui", "Oui, avec réserve", "Liste choix unique", "")
    Verifier "conformité", "égal strict", "-", CONF_NON_CONFORME, CalculerConformite("E", "Oui", "Oui, avec réserve", "Liste choix unique", "=")
    Verifier "conformité", "Sp renseigné", "-", CONF_SPECIFIQUE, CalculerConformite("Sp", "", "texte", "Texte", "")
    Verifier "conformité", "Non concerné, standard non vide", "-", CONF_NON_CONFORME, CalculerConformite("E", "Oui", NC_OFFRE, "Texte", "")
    Verifier "conformité", "standard non applicable", "-", CONF_NON_CONCERNE, CalculerConformite("D", "Oui", MSG_STANDARD_NON_APPLICABLE, "Texte", "=")
    Set q = New cQuestionSGO
    q.Niveau = "Sp": q.Reponse = "x": q.Conformite = CONF_SPECIFIQUE
    Verifier "statut", "Sp identique (cas 3)", "-", CStr(ST_SPECIFIQUE), CStr(CalculerStatut(q))
    q.Difference = DIFF_DIFFERENTE
    Verifier "statut", "Sp différente sans justification (cas 4)", "-", CStr(ST_A_JUSTIFIER), CStr(CalculerStatut(q))
    q.Justification = "Motif"
    Verifier "statut", "Sp différente justifiée (cas 4)", "-", CStr(ST_JUSTIFIEE), CStr(CalculerStatut(q))

    wsRes.Columns("A:F").ColumnWidth = 24
    Application.ScreenUpdating = True
    If nbEcarts = 0 Then wsRes.Range("A2").Value = "Aucun écart."
    wsRes.Activate
    MsgBox nbTests & " vérifications, " & nbEcarts & " écart(s)." & vbCrLf & _
           "Durée : " & Format$(Timer - debut, "0.0") & " s" & _
           IIf(nbEcarts > 0, vbCrLf & vbCrLf & "Détail dans la feuille T_Resultats.", ""), _
           IIf(nbEcarts = 0, vbInformation, vbExclamation), "SGO - Tests du moteur"
End Sub

Private Sub Verifier(test As String, contexte As String, question As String, ByVal attendu As String, ByVal obtenu As String)
    Dim r As Long
    nbTests = nbTests + 1
    If attendu = obtenu Then Exit Sub
    nbEcarts = nbEcarts + 1
    If nbEcarts > 2000 Then Exit Sub
    r = nbEcarts + 1
    wsRes.Cells(r, 1).Resize(1, 5).NumberFormat = "@"
    wsRes.Cells(r, 1).Resize(1, 5).Value = Array(test, contexte, question, TexteCellule(attendu), TexteCellule(obtenu))
End Sub

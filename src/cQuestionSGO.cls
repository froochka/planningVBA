Option Explicit

' =============================================================================
' Une question du référentiel, et son état dans le dossier ouvert.
' Les champs « référentiel » viennent de R_Questions ; les champs « état » sont
' calculés par le moteur (module SGO_Moteur) ; les champs « saisie » viennent
' du dossier de l'offre (feuille Reponses).
' =============================================================================

' --- Référentiel -------------------------------------------------------------
Public Id As String
Public Libelle As String
Public SousTheme As String          ' id du sous-thème (vide pour les questions système)
Public Ordre As Long
Public Mode As String               ' Liste choix unique, Liste choix multiple, Texte, Annexe 10 - ...
Public Choix As Variant             ' tableau de chaînes ; tableau vide si saisie libre
Public Operateur As String          ' "=" : égalité ; sinon la réponse doit contenir le standard
Public AlimenteCCAP As Boolean
Public Systeme As Boolean           ' El393 / El394 : direction et univers (utilisées par des conditions)

' --- État calculé par le moteur ---------------------------------------------------
Public Niveau As String             ' E, D, U, Sp, Titre, ou "" (standards incomplets)
Public Standard As String           ' valeur du standard applicable
Public Reponse As String            ' réponse en vigueur : saisie, ou pré-remplie avec le standard
Public Motif As String              ' "" si visible ; sinon la raison du masquage
Public Conformite As String         ' libellé de l'outil d'origine
Public Statut As Long               ' constantes ST_... du module SGO_Moteur

' --- Saisie de l'acheteur (dossier) -------------------------------------------------
Public Saisie As Boolean            ' True : l'acheteur a saisi la réponse (sinon : standard)
Public ReponseSaisie As String
Public Difference As String         ' "", "Identique", "Différente", "Non concerné" ("" : valeur par défaut)
Public Justification As String
Public Commentaire As String
Public StandardModifie As Boolean   ' standard changé depuis la saisie (page de garde modifiée)
Public ModifieLe As Date
Public LigneDossier As Long         ' ligne dans la feuille Reponses du dossier (0 : pas encore écrite)

' Vrai si la question a quelque chose à enregistrer dans le dossier
Public Function AQuelqueChose() As Boolean
    AQuelqueChose = Saisie Or Difference <> "" Or Justification <> "" Or Commentaire <> "" Or StandardModifie
End Function

Public Function EstListe() As Boolean
    EstListe = (LCase$(Left$(Mode, 5)) = "liste")
End Function

Public Function EstChoixMultiple() As Boolean
    EstChoixMultiple = (InStr(1, Mode, "multiple", vbTextCompare) > 0)
End Function

Public Function EstAnnexe() As Boolean
    EstAnnexe = (InStr(1, Mode, "annex", vbTextCompare) > 0)
End Function

Public Function NbChoix() As Long
    If IsArray(Choix) Then NbChoix = UBound(Choix) - LBound(Choix) + 1
End Function

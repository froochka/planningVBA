"""Moteur de règles du SGO, version Python.

C'est la copie conforme du module VBA `src/Moteur.bas` : mêmes règles, mêmes
noms. Il sert à vérifier la migration (comparaison avec les formules de
l'outil d'origine) et à préparer les tests que `Tests_Moteur.bas` rejoue
dans Excel. Toute règle modifiée ici doit l'être aussi dans le VBA.
"""
import csv
import os

SEP_CHOIX = "|"
MSG_STANDARD_INCOMPLET = "Standards incomplets pour cette question, contacter votre administrateur!"
MSG_REPONSE_NON_SAISIE = "Renseigner la réponse DCO"
MSG_STANDARD_NON_APPLICABLE = ("Standard non applicable pour cette offre au regard "
                               "des réponses précédentes apportées")
NC_OFFRE = "Non concerné pour l'offre"
NC_DIRECTION = "Non concerné pour la direction/univers"


def lire_csv(chemin):
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


class Referentiel:
    def __init__(self, dossier):
        self.questions = {}
        for r in lire_csv(os.path.join(dossier, "R_Questions.csv")):
            r["choix"] = r["choix"].split(SEP_CHOIX) if r["choix"] else []
            self.questions[r["id"].lower()] = r
        self.categories = {r["id"].lower(): r for r in lire_csv(os.path.join(dossier, "R_Categories.csv"))}
        self.std = {}
        for r in lire_csv(os.path.join(dossier, "R_Standards.csv")):
            self.std[(r["question"].lower(), r["type"], r["cible"].lower())] = (r["niveau"], r["valeur"])
        self.dir_std = {}
        self.univers = {}
        for r in lire_csv(os.path.join(dossier, "R_Directions.csv")):
            if r["type"] == "STANDARD":
                self.dir_std[r["code"].lower()] = r["a_standards"] == "Oui"
            elif r["type"] == "UNIVERS":
                self.univers[r["code"].lower()] = r["a_standards"] == "Oui"
        self.conditions = {}
        for r in lire_csv(os.path.join(dossier, "R_Conditions.csv")):
            self.conditions.setdefault(r["fille"].lower(), []).append((r["mere"].lower(), r["operateur"], r["critere"]))


# --- Standard ---------------------------------------------------------------
def niveau(ref, q, d, u):
    """Niveau de standard : E, D, U, Sp, Titre, ou '' (incomplet)."""
    q, d, u = q.lower(), d.lower(), u.lower()
    if not ref.dir_std.get(d, False):
        return "Sp"                                   # direction sans colonne de standards
    n = ref.std.get((q, "D", d), ("", ""))[0]
    if n in ("D", "U") and ref.univers.get(u, False):
        vu = ref.std.get((q, "U", u), ("", ""))[1]
        if vu.lower() == "sp":
            return "Sp"
    return n


def standard(ref, q, d, u, niv):
    q, d, u = q.lower(), d.lower(), u.lower()
    if niv == "E":
        return ref.std.get((q, "E", ""), ("", ""))[1]
    if niv == "D":
        return ref.std.get((q, "D", d), ("", ""))[1]
    if niv == "U":
        return ref.std.get((q, "U", u), ("", ""))[1]
    return ""


def est_annexe(mode):
    return "annex" in mode.lower()


def reponse_initiale(niv, std, mode):
    """Réponse pré-remplie au chargement (ES-6)."""
    if est_annexe(mode):
        return mode[len("Annexe "):] if mode.lower().startswith("annexe ") else mode
    return std if niv in ("E", "D", "U") else ""


def conformite(niv, std, rep, mode, op):
    """Formule de l'outil d'origine (colonne Conformité)."""
    if (niv == "Titre" or rep[:11].lower() == "non concern" or est_annexe(mode)) and niv != "Sp":
        if std == "" or std[:11].lower() == "non concern":
            return "Non concerné"
        return "Non conforme au standard"
    if niv == "Sp" and rep != "":
        return "Spécifique Offre"
    if (niv == "" or std == "") and niv != "Sp":
        return MSG_STANDARD_INCOMPLET
    if rep == "":
        return MSG_REPONSE_NON_SAISIE
    if op == "=":
        if rep.lower() == std.lower():
            return "Conforme au standard"
    elif std.lower() in rep.lower():
        return "Conforme au standard"
    return "Non concerné" if rep == MSG_STANDARD_NON_APPLICABLE else "Non conforme au standard"


def masquee_par_standard(std):
    """Standard « non concerné pour la direction/univers » ou « non applicable »."""
    s = std.lower()
    if "direction/" in s:
        return NC_DIRECTION
    if "non applicable" in s:
        return MSG_STANDARD_NON_APPLICABLE
    return ""


# --- Conditions ---------------------------------------------------------------
def condition_vraie(op, critere, reponse):
    if op == "=":
        return reponse.lower() == critere.lower()
    return critere.lower() in reponse.lower()


class Dossier:
    """Réponses d'une offre et calcul de l'état de chaque question."""

    def __init__(self, ref, direction, univers):
        self.ref, self.d, self.u = ref, direction, univers
        self.saisies = {}          # id -> réponse saisie par l'acheteur

    def evaluer(self):
        ref = self.ref
        self.etat = {}
        self.visible = {}
        ordre = sorted(ref.questions.values(), key=lambda r: int(r["ordre"]))
        for r in ordre:
            q = r["id"].lower()
            if r["systeme"] == "Oui":
                rep = self.d if q == "el393" else self.u
                self.etat[q] = {"niveau": "", "standard": "", "reponse": rep, "motif": ""}
                self.visible[q] = True
                continue
            niv = niveau(ref, q, self.d, self.u)
            std = standard(ref, q, self.d, self.u, niv)
            rep = self.saisies.get(q, reponse_initiale(niv, std, r["mode"]))
            self.etat[q] = {"niveau": niv, "standard": std, "reponse": rep, "motif": ""}
        for r in ordre:
            q = r["id"].lower()
            if r["systeme"] == "Oui":
                continue
            motif = masquee_par_standard(self.etat[q]["standard"])
            if not motif and not self.element_visible(q):
                motif = NC_OFFRE
            self.etat[q]["motif"] = motif
            self.visible[q] = motif == ""
            e = self.etat[q]
            e["conformite"] = "" if motif else conformite(e["niveau"], e["standard"], e["reponse"],
                                                          r["mode"], r["operateur"])

    def element_visible(self, x, pile=()):
        """Visible si toutes ses conditions sont vraies, avec mères visibles, et si
        sa catégorie parente est visible (P5)."""
        if x in pile:
            return False                              # boucle : considéré masqué
        pile = pile + (x,)
        for mere, op, crit in self.ref.conditions.get(x, []):
            if mere in self.visible:
                vis_mere = self.visible[mere]
            else:
                vis_mere = masquee_par_standard(self.etat[mere]["standard"]) == "" and \
                    self.element_visible(mere, pile)
            if not vis_mere or not condition_vraie(op, crit, self.etat[mere]["reponse"]):
                return False
        parent = self.ref.questions[x]["sous_theme"] if x in self.ref.questions else \
            self.ref.categories[x]["parent"]
        return self.element_visible(parent.lower(), pile) if parent else True

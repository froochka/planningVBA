"""Vérificateur statique approximatif du VBA de src/ (pour Claude : le VBA n'est pas compilable hors d'Excel).

Usage : python outils/verifier_vba.py src [fichiers à contrôler pour les identifiants...]

- encodage : caractères hors cp1252 (Excel lit les fichiers en ANSI)
- blocs : If/End If, For/Next, Do/Loop, Select/End Select, With/End With, Sub/Function
- continuations de ligne (max 24)
- déclarations en double dans une procédure (sans casse)
- identifiants non déclarés (heuristique)
"""
import re, sys, os, glob

SRC = sys.argv[1]
FICHIERS = sorted(glob.glob(os.path.join(SRC, "*.bas")) + glob.glob(os.path.join(SRC, "*.cls")) + glob.glob(os.path.join(SRC, "*.frm")))
CIBLES = set(sys.argv[2:])  # fichiers à vérifier pour les identifiants

KEYWORDS = set("""
and as boolean byref byval call case const date dim do double each else elseif end enum
error exit explicit false for function get global gosub goto if in integer is let like long loop
me mod new next not nothing null object on option optional or paramarray preserve private property
public redim resume select set single static step string sub then to true type until variant wend
while with withevents xor currency byte decimal empty friend implements raiseevent attribute
lbound ubound collection
""".split())

BUILTINS = set(x.lower() for x in """
Len Mid Mid$ InStr InStrRev Replace Split Trim Trim$ LCase LCase$ UCase UCase$ Format Format$ CStr CLng CInt CDbl CSng CBool
Val IsArray Array ChrW Chr Chr$ IIf Now Timer MsgBox Dir CreateObject GetSetting SaveSetting IsDate CDate IsEmpty
IsError Int StrComp Left Left$ Right Right$ RGB Err Application ThisWorkbook Workbooks ActiveSheet ActiveWorkbook
Worksheets Worksheet Workbook Range VBA UserForms Debug DoEvents IsNumeric IsObject TypeName VarType Abs Round
vbCrLf vbCr vbLf vbTab vbTextCompare vbBinaryCompare vbYes vbNo vbYesNo vbYesNoCancel vbOKCancel vbCancel vbQuestion
vbInformation vbExclamation vbCritical vbBoolean vbObjectError vbNullString vbString
xlSheetHidden xlSheetVisible xlSheetVeryHidden xlUp xlWhole xlWBATWorksheet xlOpenXMLWorkbook
MSForms ListBox TabStrip ComboBox TextBox OptionButton Label Frame MultiPage Page CheckBox ReturnInteger ReturnBoolean
Controls Cells Rows Columns Unload Load Hide Show Space String
UserForm_Initialize UserForm_QueryClose Cancel CloseMode KeyCode Shift
Kill FileLen Environ Choose
""".split())

def nettoyer(ligne):
    """Retire chaînes et commentaire."""
    out, i, dans = [], 0, False
    while i < len(ligne):
        ch = ligne[i]
        if ch == '"':
            if dans and i + 1 < len(ligne) and ligne[i + 1] == '"':
                i += 2; continue
            dans = not dans; out.append('""' if not dans else ''); i += 1; continue
        if not dans and ch == "'":
            break
        if not dans:
            out.append(ch)
        i += 1
    return "".join(out)

def lignes_logiques(texte):
    """Joint les continuations ; renvoie (n° ligne, ligne nettoyée, nb continuations)."""
    res, acc, debut, nb = [], "", None, 0
    for n, l in enumerate(texte.splitlines(), 1):
        c = nettoyer(l).rstrip()
        if debut is None:
            debut = n
        if c.endswith(" _"):
            acc += c[:-2] + " "; nb += 1
            continue
        acc += c
        res.append((debut, acc, nb))
        acc, debut, nb = "", None, 0
    return res

def instructions(ligne):
    """Découpe une ligne logique sur les ':' hors chaîne (chaînes déjà vidées)."""
    # attention aux étiquettes "erreur:" : traitées comme instruction vide
    return [p.strip() for p in re.split(r":(?!=)", ligne)]

erreurs = []
decl_globales = {}   # nom -> fichier
membres_classe = {}

def module_nom(f):
    return os.path.splitext(os.path.basename(f))[0]

# --- 1er passage : déclarations publiques et noms de modules
for f in FICHIERS:
    texte = open(f, encoding="utf-8").read()
    decl_globales[module_nom(f).lower()] = f
    for (n, l, nb) in lignes_logiques(texte):
        s = l.strip()
        m = re.match(r"(?i)^(public|global)?\s*(private\s+)?(sub|function|property\s+(get|let|set))\s+(\w+)", s)
        if m and not re.match(r"(?i)^private", s):
            decl_globales[m.group(5).lower()] = f
        m = re.match(r"(?i)^(public|global)\s+(const\s+)?(withevents\s+)?(.*)$", s)
        if m and not re.match(r"(?i)^public\s+(sub|function|property)", s):
            for part in m.group(4).split(","):
                mm = re.match(r"\s*(\w+)", part)
                if mm:
                    decl_globales[mm.group(1).lower()] = f

for f in FICHIERS:
    texte = open(f, encoding="utf-8").read()
    nomf = os.path.basename(f)
    # encodage
    for n, l in enumerate(texte.splitlines(), 1):
        try:
            l.encode("cp1252")
        except UnicodeEncodeError as e:
            erreurs.append(f"{nomf}:{n}: caractère hors cp1252 : {l[e.start:e.end]!r}")
    ll = lignes_logiques(texte)
    for (n, l, nb) in ll:
        if nb > 24:
            erreurs.append(f"{nomf}:{n}: {nb} continuations (max 24)")
        if len(l) > 1023:
            erreurs.append(f"{nomf}:{n}: ligne logique trop longue ({len(l)})")
    # blocs
    pile = []
    module_vars = set()
    proc = None
    locaux = {}
    for (n, l, nb) in ll:
        for ins in instructions(l):
            s = ins.strip()
            if not s:
                continue
            low = s.lower()
            low = re.sub(r"^(public|private|friend|static)\s+", "", low)
            if re.match(r"^(sub|function|property\s+(get|let|set))\s+\w+", low):
                if proc:
                    erreurs.append(f"{nomf}:{n}: procédure ouverte dans {proc}")
                proc = re.match(r"^(?:sub|function|property\s+\w+)\s+(\w+)", low).group(1)
                pile = []
                locaux = {}
                # paramètres
                mpar = re.search(r"\((.*)\)", s)
                if mpar:
                    for p in mpar.group(1).split(","):
                        mm = re.match(r"(?i)\s*(optional\s+)?(byval\s+|byref\s+)?(paramarray\s+)?(\w+)", p)
                        if mm and mm.group(4):
                            k = mm.group(4).lower()
                            if k in locaux:
                                erreurs.append(f"{nomf}:{n}: paramètre en double {k}")
                            locaux[k] = n
                continue
            if re.match(r"^end\s+(sub|function|property)\b", low):
                if pile:
                    erreurs.append(f"{nomf}:{n}: bloc(s) non fermé(s) dans {proc} : {pile}")
                proc = None
                continue
            if re.match(r"^exit\s+(sub|function|property|for|do)", low):
                continue
            # déclarations
            m = re.match(r"^(dim|static|const|redim(\s+preserve)?)\s+(.*)$", low) if proc else \
                re.match(r"^(dim|const|withevents|private|public|global)?\s*(const\s+|withevents\s+)?(.*)$", low)
            if proc and m and m.group(1).startswith(("dim", "static", "const")):
                for part in re.split(r",(?![^(]*\))", m.group(3)):
                    mm = re.match(r"\s*(withevents\s+)?(\w+)", part)
                    if mm:
                        k = mm.group(2)
                        if k in locaux:
                            erreurs.append(f"{nomf}:{n}: déclaration en double « {k} » dans {proc}")
                        locaux[k] = n
            low0 = s.lower()
            if re.match(r"^(private|public|friend)?\s*(sub|function|property)\b", low0):
                pass
            if not proc and re.match(r"^(dim|private|public|global|const)\b", low0) and not re.search(r"\b(sub|function|property|type|enum)\b", low0.split("(")[0]):
                body = re.sub(r"^(dim|private|public|global)\s+", "", low0)
                body = re.sub(r"^(const|withevents)\s+", "", body)
                for part in re.split(r",(?![^(]*\))", body):
                    mm = re.match(r"\s*(withevents\s+)?(\w+)", part)
                    if mm:
                        module_vars.add(mm.group(2))
            # blocs
            if re.match(r"^if\b.*\bthen\s*$", low):
                pile.append("If")
            elif re.match(r"^elseif\b", low) or low == "else":
                if not pile or pile[-1] != "If":
                    erreurs.append(f"{nomf}:{n}: {s.split()[0]} hors If")
            elif re.match(r"^end\s+if\b", low):
                if not pile or pile.pop() != "If":
                    erreurs.append(f"{nomf}:{n}: End If sans If")
            elif re.match(r"^for\b", low):
                pile.append("For")
            elif re.match(r"^next\b", low):
                if not pile or pile.pop() != "For":
                    erreurs.append(f"{nomf}:{n}: Next sans For")
            elif re.match(r"^do\b", low):
                pile.append("Do")
            elif re.match(r"^loop\b", low):
                if not pile or pile.pop() != "Do":
                    erreurs.append(f"{nomf}:{n}: Loop sans Do")
            elif re.match(r"^while\b", low):
                pile.append("While")
            elif re.match(r"^wend\b", low):
                if not pile or pile.pop() != "While":
                    erreurs.append(f"{nomf}:{n}: Wend sans While")
            elif re.match(r"^select\s+case\b", low):
                pile.append("Select")
            elif re.match(r"^end\s+select\b", low):
                if not pile or pile.pop() != "Select":
                    erreurs.append(f"{nomf}:{n}: End Select sans Select")
            elif re.match(r"^with\b", low):
                pile.append("With")
            elif re.match(r"^end\s+with\b", low):
                if not pile or pile.pop() != "With":
                    erreurs.append(f"{nomf}:{n}: End With sans With")
            elif re.match(r"^if\b.*\bthen\b", low):
                # If sur une ligne : vérifier les sous-instructions (rien à empiler)
                pass
    if proc:
        erreurs.append(f"{nomf}: procédure {proc} non fermée")

    for (n, l, nb) in ll:
        mm = re.match(r"(?i)^\s*(?:private\s+|public\s+|friend\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(\w+)", l)
        if mm:
            module_vars.add(mm.group(1).lower())
        mm = re.match(r"^(\w+):\s*$", l.strip())
        if mm:
            module_vars.add(mm.group(1).lower())
    # identifiants non déclarés (fichiers ciblés)
    if CIBLES and os.path.basename(f) not in CIBLES:
        continue
    proc = None
    locaux = set()
    for (n, l, nb) in ll:
        s = l.strip()
        low = re.sub(r"^(public|private|friend|static)\s+", "", s.lower())
        if re.match(r"^(sub|function|property\s+(get|let|set))\s+\w+", low):
            proc = True
            locaux = set()
            mpar = re.search(r"\((.*)\)", s)
            if mpar:
                for p in mpar.group(1).split(","):
                    mm = re.match(r"(?i)\s*(optional\s+)?(byval\s+|byref\s+)?(paramarray\s+)?(\w+)", p)
                    if mm and mm.group(4):
                        locaux.add(mm.group(4).lower())
            nom = re.match(r"^(?:sub|function|property\s+\w+)\s+(\w+)", low).group(1)
            locaux.add(nom)
            continue
        if re.match(r"^end\s+(sub|function|property)\b", low):
            proc = None
            continue
        if not proc:
            continue
        for ins in instructions(l):
            li = ins.strip().lower()
            m = re.match(r"^(dim|static|const)\s+(.*)$", li)
            if m:
                for part in re.split(r",(?![^(]*\))", m.group(2)):
                    mm = re.match(r"\s*(\w+)", part)
                    if mm:
                        locaux.add(mm.group(1))
                # type names after As are fine; skip checking
                continue
            if re.match(r"^\w+:$", ins.strip() + ":") and re.match(r"^\w+$", ins.strip()):
                # étiquette ou appel sans argument
                pass
            # retirer les noms d'arguments nommés (Filename:=)
            texte_ins = re.sub(r"\w+\s*:=", "", ins)
            texte_ins = re.sub(r"&H[0-9A-Fa-f]+&?", "0", texte_ins)
            for m in re.finditer(r"(?<![\w.$])([A-Za-z_]\w*)\$?", texte_ins):
                nom = m.group(1).lower()
                debut = m.start(1)
                avant = texte_ins[:debut].rstrip()
                if avant.endswith("."):
                    continue
                if nom in KEYWORDS or nom in BUILTINS or nom in locaux or nom in module_vars or nom in decl_globales:
                    continue
                # "As Type"
                if re.search(r"(?i)\bas\s+(new\s+)?$", avant):
                    continue
                # étiquettes GoTo
                if re.search(r"(?i)\b(goto|resume)\s+$", avant):
                    continue
                erreurs.append(f"{nomf}:{n}: identifiant non déclaré ? « {m.group(1)} »")

for e in erreurs:
    print(e)
print(f"{len(erreurs)} remarque(s)")

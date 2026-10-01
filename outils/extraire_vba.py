"""Extrait le code VBA d'un classeur .xlsm vers src/ sans passer par Excel.

Produit le meme format que la macro ExporterVBA (outils/VbaSync.bas) :
un fichier par module, en UTF-8, sans les en-tetes "Attribute VB_...".

Usage : python outils/extraire_vba.py VBA_maquette_SGO.xlsm [dossier_src]
Prerequis : pip install oletools
"""
import pathlib
import sys

from oletools.olevba import VBA_Parser

EXTENSIONS = {".bas", ".cls", ".frm"}


def nettoyer(code: str) -> str:
    lignes = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    if lignes and lignes[0].startswith("VERSION "):
        while i < len(lignes) and not lignes[i].startswith("Attribute VB_Name"):
            i += 1
    while i < len(lignes) and lignes[i].startswith("Attribute VB_"):
        i += 1
    return "\n".join(lignes[i:]).rstrip("\r\n")


def extraire(classeur: pathlib.Path, dossier: pathlib.Path) -> None:
    parser = VBA_Parser(str(classeur))
    if not parser.detect_vba_macros():
        sys.exit(f"Aucun code VBA trouve dans {classeur}")

    dossier.mkdir(parents=True, exist_ok=True)
    ecrits = set()
    for _, _, nom_fichier, code in parser.extract_macros():
        if pathlib.Path(nom_fichier).suffix.lower() not in EXTENSIONS:
            continue
        if isinstance(code, bytes):
            code = code.decode("cp1252", errors="replace")
        code = nettoyer(code)
        if not code.strip() or nom_fichier in ecrits:
            continue
        (dossier / nom_fichier).write_text(code + "\n", encoding="utf-8")
        ecrits.add(nom_fichier)
        print(f"  {nom_fichier}")
    parser.close()

    for f in dossier.iterdir():
        if f.suffix.lower() in EXTENSIONS and f.name not in ecrits:
            f.unlink()
            print(f"  (supprime) {f.name}")
    print(f"{len(ecrits)} module(s) extrait(s) vers {dossier}/")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    extraire(pathlib.Path(sys.argv[1]),
             pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else "src"))

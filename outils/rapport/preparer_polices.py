"""
preparer_polices.py - Prépare les polices Inter (TrueType, sous-ensemble latin) utilisées par le rapport.

Chromium n'incorpore proprement dans un PDF que les polices TrueType : les fichiers Inter publiés en
OpenType/CFF y deviennent des polices Type 3 (rendu inégal selon les lecteurs PDF). Ce script convertit
les contours CFF en contours TrueType et ne garde que les caractères utiles (latin, ponctuation,
symboles), ce qui donne des fichiers légers, livrés dans outils/rapport/polices/.

À n'exécuter qu'une fois (les fichiers produits sont fournis avec la boîte à outils) :
  python outils/rapport/preparer_polices.py /usr/share/fonts/opentype/inter
Licence : SIL Open Font License 1.1 (voir polices/OFL.txt) ; les fichiers dérivés restent sous OFL.
"""
from __future__ import annotations

import sys
from pathlib import Path

from fontTools import subset
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable

ICI = Path(__file__).resolve().parent
GRAISSES = ["Regular", "Italic", "SemiBold", "Bold", "BoldItalic"]
UNICODES = (list(range(0x20, 0x7F)) + list(range(0xA0, 0x180)) + list(range(0x2000, 0x2070))
            + list(range(0x2070, 0x20A0)) + list(range(0x20A0, 0x20C0)) + list(range(0x2100, 0x2150))
            + list(range(0x2190, 0x2200)) + list(range(0x2200, 0x2300)) + list(range(0x25A0, 0x2600))
            + [0x2713, 0x2717, 0x27E6, 0x27E7])


def cff_vers_truetype(police: TTFont, erreur_max=1.0):
    """Remplace la table CFF par des contours quadratiques (glyf/loca)."""
    ordre = police.getGlyphOrder()
    jeu = police.getGlyphSet()
    glyphes = {}
    for nom in ordre:
        stylo = TTGlyphPen(jeu)
        jeu[nom].draw(Cu2QuPen(stylo, erreur_max, reverse_direction=True))
        glyphes[nom] = stylo.glyph()
    police["loca"] = newTable("loca")
    police["glyf"] = glyf = newTable("glyf")
    glyf.glyphOrder = ordre
    glyf.glyphs = glyphes
    del police["CFF "]
    for t in ("VORG",):
        if t in police:
            del police[t]
    glyf.compile(police)
    hmtx = police["hmtx"]
    for nom, g in glyf.glyphs.items():
        if hasattr(g, "xMin"):
            hmtx[nom] = (hmtx[nom][0], g.xMin)
    police["maxp"] = maxp = newTable("maxp")
    maxp.tableVersion = 0x00010000
    for attr in ("maxZones", "maxTwilightPoints", "maxStorage", "maxFunctionDefs", "maxInstructionDefs",
                 "maxStackElements", "maxSizeOfInstructions", "maxComponentElements"):
        setattr(maxp, attr, 1 if attr == "maxZones" else 0)
    maxp.compile(police)
    post = police["post"]
    post.formatType = 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = ordre
    police.sfntVersion = "\000\001\000\000"


def main():
    source = Path(sys.argv[1] if len(sys.argv) > 1 else "/usr/share/fonts/opentype/inter")
    sortie = ICI / "polices"
    sortie.mkdir(exist_ok=True)
    for g in GRAISSES:
        police = TTFont(source / f"Inter-{g}.otf")
        opts = subset.Options()
        opts.layout_features = ["*"]
        opts.name_IDs = ["*"]
        opts.name_languages = ["*"]
        opts.notdef_outline = True
        opts.glyph_names = True
        sub = subset.Subsetter(opts)
        sub.populate(unicodes=UNICODES)
        sub.subset(police)
        cff_vers_truetype(police)
        cible = sortie / f"Inter-{g}.ttf"
        police.save(cible)
        print(f"{cible.name} : {cible.stat().st_size / 1024:.0f} Ko")


if __name__ == "__main__":
    main()

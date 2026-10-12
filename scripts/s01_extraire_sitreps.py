"""
s01_extraire_sitreps.py - Extraction des SitReps nationaux MVE (INSP/SGI) vers la base Excel.

Entrée  : data/entree/sitreps/*.doc ou *.docx (format du SitRep national MVE-BDBV)
Sortie  : feuilles sitrep_meta, epi_province, epi_zs, alertes_province, contacts_province,
          labo_province, pec_province, poe_province, sitrep_texte de data/base/base_CGPP_MVE.xlsx
          + data/intermediaire/pourcentages_texte.csv (affirmations « x % (a/b) » du texte)
          + data/intermediaire/chiffres_texte.csv (chiffres cités dans le texte, pour contrôle)

Usage   : python scripts/s01_extraire_sitreps.py            (tous les SitReps du dossier)
          python scripts/s01_extraire_sitreps.py fichier.doc (un seul SitRep)

Le SitRep national met ses chiffres clés dans des zones de texte (encadrés) : elles sont lues
directement dans le XML du document, en plus des tableaux et des paragraphes.
"""
from __future__ import annotations

import re
import sys
from datetime import date, datetime
from pathlib import Path

import docx
import pandas as pd
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from commun import (PROVINCES, SCHEMA, charger_config, cle, ecrire_tables, fraction_fr, fusionner,  # noqa: E402
                    journal, lire_table, nombre_fr, norm_province, vers_docx)

LOG = journal("sitreps")

MOIS = {"janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
        "juillet": 7, "août": 8, "aout": 8, "septembre": 9, "octobre": 10, "novembre": 11,
        "décembre": 12, "decembre": 12}
NUM = r"\d{1,3}(?:[    ]\d{3})+|\d+"
PCT = r"\d{1,3}(?:[.,]\d+)?"
RE_PROV = (r"(Ituri|Nord[- ]Kivu|Nord[- ]Kivi|Haut[- ]U[ée]l[ée]|Bas[- ]U[ée]l[ée]|Tshopo|"
           r"Sud[- ]Kivu|Sud[- ]Ubangi)")

SECTIONS = [
    ("surveillance épidémiologique", "Surveillance"),
    ("points d'entrées", "PoE/PoC"), ("points d’entrées", "PoE/PoC"),
    ("laboratoire", "Laboratoire"),
    ("prévention et contrôle", "PCI/EDS"),
    ("prise en charge", "Prise en charge"),
    ("vaccination", "Vaccination"),
    ("continuité des soins", "Continuité des soins"),
    ("communication sur les risques", "CREC"),
    ("santé mentale", "SMSPS"),
    ("logistique", "Logistique"),
    ("sécurité", "Sécurité"),
    ("psea", "PSEA"),
    ("coordination", "Coordination"),
]


# --------------------------------------------------------------------------
# Lecture du document
# --------------------------------------------------------------------------

def texte_propre(s: str) -> str:
    return re.sub(r"[ \t ]+", " ", s.replace(" ", " ")).strip()


def zones_de_texte(doc) -> list[str]:
    """Texte des encadrés (w:txbxContent), dédoublonné (Word stocke chaque encadré deux fois)."""
    vus, sortie = set(), []
    for tx in doc.element.body.iter(qn("w:txbxContent")):
        paras = []
        for p in tx.iter(qn("w:p")):
            t = "".join(n.text or "" for n in p.iter(qn("w:t")))
            if t.strip():
                paras.append(t)
        bloc = "\n".join(paras).strip()
        if bloc and bloc not in vus:
            vus.add(bloc)
            sortie.append(bloc)
    return sortie


def lignes_tableau(t) -> list[list[str]]:
    lignes = []
    for r in t.rows:
        lignes.append([texte_propre(c.text) for c in r.cells])
    return lignes


def date_fr(s: str):
    m = re.search(r"(\d{1,2})\s+([a-zéû]+)\s+(\d{4})", s, flags=re.I)
    if not m or m.group(2).lower() not in MOIS:
        return None
    return date(int(m.group(3)), MOIS[m.group(2).lower()], int(m.group(1)))


def entier(s):
    v = nombre_fr(s)
    return None if v is None else int(round(v))


# --------------------------------------------------------------------------
# Tableaux
# --------------------------------------------------------------------------

def trouver_tableaux(doc) -> dict[str, list[list[str]]]:
    trouves = {}
    for t in doc.tables:
        lignes = lignes_tableau(t)
        if not lignes:
            continue
        entete = " ".join(" ".join(l) for l in lignes[:3]).lower()
        if "province / zone de santé" in entete and "t2" not in trouves:
            trouves["t2"] = lignes
        elif "nouveaux cas confirmés" in entete and "zones de santé touchées" in entete:
            trouves.setdefault("t1", lignes)
        elif "alertes" in entete and lignes[0] and lignes[0][0].lower().startswith("dps"):
            trouves.setdefault("t3", lignes)
        elif lignes[0] and lignes[0][0].lower().startswith("indicateurs") and "ituri" in entete:
            trouves.setdefault("t4", lignes)
    return trouves


def lire_t1(lignes, no, d):
    sortie = []
    for l in lignes[1:]:
        prov = norm_province(l[0])
        if not prov:
            continue
        num, den, _ = fraction_fr(l[5]) if len(l) > 5 else (None, None, None)
        sortie.append(dict(sitrep_no=no, date_rapport=d, province=prov,
                           nouveaux_cas_24h=entier(l[1]), cas_cumul=entier(l[2]),
                           deces_cumul=entier(l[3]), letalite_publiee=nombre_fr(l[4]),
                           zs_touchees=num, zs_total=den))
    return sortie


def lire_t2(lignes, no, d):
    sortie, prov = [], None
    for l in lignes:
        nom = l[0].replace("*", "").strip()
        if not nom or nom.lower().startswith(("province", "total")) or len(l) < 8:
            continue
        if all(not x for x in l[1:]):
            continue
        p = norm_province(nom)
        if p and p != prov:          # en-tête de province (une ZS peut porter le nom de sa province : Tshopo)
            prov = p
            continue
        sortie.append(dict(
            sitrep_no=no, date_rapport=d, province=prov, zone_sante=nom,
            cas_cumul=entier(l[1]), deces_cumul=entier(l[2]), letalite_publiee=nombre_fr(l[3]),
            nouveaux_cas_24h=entier(l[4]), deces_communautaires_24h=entier(l[5]),
            deces_cte_24h=entier(l[6]), deces_total_24h=entier(l[7]),
            est_a_ventiler=cle(nom) == "aventiler"))
    return sortie


def lire_t2_totaux(lignes) -> dict:
    """Lignes province et Total du tableau 2 (pour le contrôle de cohérence)."""
    tot = {}
    for l in lignes:
        nom = l[0].replace("*", "").strip()
        p = norm_province(nom) or ("Total" if nom.lower() == "total" else None)
        if p and len(l) >= 8 and p not in tot:
            tot[p] = dict(cas_cumul=entier(l[1]), deces_cumul=entier(l[2]),
                          nouveaux_cas_24h=entier(l[4]), deces_communautaires_24h=entier(l[5]),
                          deces_cte_24h=entier(l[6]), deces_total_24h=entier(l[7]))
    return tot


def lire_t3(lignes, no, d):
    sortie = []
    for l in lignes:
        prov = norm_province(l[0])
        if not prov:
            continue
        v = [x for x in l[1:]]
        nd = any(x.strip().upper() == "ND" for x in v)
        vals = [entier(x) for x in v] + [None] * 10
        sortie.append(dict(sitrep_no=no, date_rapport=d, province=prov,
                           statut_donnees="ND" if nd else "OK",
                           recues_vivants=vals[0], recues_deces=vals[1], recues_total=vals[2],
                           validees_vivants=vals[3], validees_deces=vals[4],
                           invalidees_vivants=vals[5], invalidees_deces=vals[6],
                           suspects_investigues=vals[7], suspects_transferes=vals[8]))
    return sortie


def lire_t4(lignes, no, d):
    sortie = []
    entete = lignes[0]
    for l in lignes[1:]:
        if not l or not l[0]:
            continue
        for j in range(1, len(entete)):
            if j >= len(l):
                continue
            num, den, pct = fraction_fr(l[j])
            valeur = pct if pct is not None else nombre_fr(l[j])
            sortie.append(dict(sitrep_no=no, date_rapport=d, indicateur=l[0],
                               province=norm_province(entete[j]) or entete[j],
                               valeur_texte=l[j], valeur=valeur, numerateur=num, denominateur=den))
    return sortie


# --------------------------------------------------------------------------
# Texte : encadrés et paragraphes
# --------------------------------------------------------------------------

def kpi_encadres(boites: list[str]) -> dict:
    r = {}
    for b in boites:
        t = " ".join(b.split())
        for nom, motif in [
            ("cas_cumul", r"^CUMUL DES CAS ([\d ]+)$"),
            ("deces_cumul", r"^CUMUL DES D[ÉE]C[ÈE]S CONFIRM[ÉE]S ([\d ]+)$"),
            ("gueris_cumul", r"^CUMUL DES GU[ÉE]RIS ([\d ]+)$"),
            ("taux_suivi_contacts", r"^TAUX DE SUIVI DES CONTACTS \(Du Jour\) (" + PCT + r") ?%$"),
            ("letalite", r"^LETALITE (" + PCT + r") ?%$"),
            ("patients_isolement", r"^PATIENTS EN ISOLEMENT/CTE ([\d ]+)$"),
        ]:
            m = re.match(motif, t, flags=re.I)
            if m:
                r[nom] = nombre_fr(m.group(1))
        m = re.match(r"^(\d+) Aires de santé touchées Sur ([\d ]+)", t, flags=re.I)
        if m:
            r["aires_sante_touchees"], r["aires_sante_total"] = int(m.group(1)), nombre_fr(m.group(2))
    # Faits saillants : l'encadré qui suit le titre
    for i, b in enumerate(boites):
        if b.upper().startswith("FAITS SAILLANTS") and i + 1 < len(boites):
            fs = " ".join(boites[i + 1].split())
            r["faits_saillants"] = fs
            m = re.search(r"\((\d+)\)\s+nouveaux cas confirmés", fs)
            r["nouveaux_cas_24h"] = int(m.group(1)) if m else None
            m = re.search(r"dont\s+(\d+)\s+décès communautaires", fs)
            r["deces_communautaires_24h"] = int(m.group(1)) if m else None
            m = re.search(r"(\d+)\s+décès ont été enregistrés dans les CTE", fs)
            r["deces_cte_24h"] = int(m.group(1)) if m else None
            m = re.search(r"Un total de\s+(" + NUM + r")\s+nouveaux échantillons", fs)
            r["echantillons_24h"] = nombre_fr(m.group(1)) if m else None
            break
    for b in boites:
        t = " ".join(b.split())
        m = re.search(r"(\d+)\s+patients\s+\(([^)]*)\)\s+ont été déclarés guéris", t)
        if m:
            r["gueris_24h"] = int(m.group(1))
            r["gueris_24h_detail"] = m.group(2)
    return r


def decoupe_sections(doc) -> list[dict]:
    paras, section, sous = [], "En-tête", ""
    for i, p in enumerate(doc.paragraphs):
        t = texte_propre(p.text)
        if not t:
            continue
        runs = [r for r in p.runs if r.text.strip()]
        tout_gras = bool(runs) and all(r.bold or (r.style is not None and r.style.font.bold)
                                       for r in runs)
        court = len(t) < 120 and tout_gras   # les titres du SitRep sont courts et en gras
        tl = t.lower()
        if court:
            nouveau = None
            for motif, nom in SECTIONS:
                if motif in tl and not tl.startswith(("au ", "en ", "à ")):
                    nouveau = nom
                    break
            if nouveau:
                section, sous = nouveau, ""
                continue
            if tl.startswith("principales actions"):
                sous = "Principales actions"
                continue
            if tl.startswith("défis"):
                sous = "Défis"
                continue
        paras.append(dict(ordre=i, section=section, sous_section=sous, texte=t))
    return paras


def province_du_paragraphe(t: str):
    m = re.match(r"^\W*(?:En|Au|À la|A la|Aux|À l['’])?\s*" + RE_PROV, t)
    return norm_province(m.group(1)) if m else None


def lire_labo(paras, no, d):
    sortie = []
    for p in paras:
        if p["section"] != "Laboratoire":
            continue
        t = p["texte"]
        prov = province_du_paragraphe(t)
        if not prov:
            continue
        pos = re.search(r"(" + NUM + r")\s+nouveaux?\s+résultats?\s+positifs?", t)
        ech = re.search(r"sur\s+(" + NUM + r")\s+[ée]ch\w*", t)
        viv = re.search(r"\((\d+)\s+vivants?\s+et\s+(\d+)\s+d[ée]c[èée]s\)", t)
        ppub = re.search(r"positivit[ée]\s*:\s*(" + PCT + r")\s*%", t)
        if not pos:
            neg = re.search(r"(" + NUM + r")\s+(?:nouveaux\s+)?[ée]ch\w*", t)
            if neg and "négatif" in t:
                sortie.append(dict(sitrep_no=no, date_rapport=d, province=prov,
                                   echantillons=entier(neg.group(1)), positifs=0,
                                   positifs_vivants=0, positifs_deces=0, positivite_publiee=None,
                                   texte_source=t))
            continue
        sortie.append(dict(sitrep_no=no, date_rapport=d, province=prov,
                           echantillons=entier(ech.group(1)) if ech else None,
                           positifs=entier(pos.group(1)),
                           positifs_vivants=int(viv.group(1)) if viv else None,
                           positifs_deces=int(viv.group(2)) if viv else None,
                           positivite_publiee=nombre_fr(ppub.group(1)) if ppub else None,
                           texte_source=t))
    return sortie


def lire_pec(paras, no, d):
    sortie = []
    for p in paras:
        if p["section"] != "Prise en charge" or p["sous_section"] == "Défis":
            continue
        t = p["texte"]
        prov = province_du_paragraphe(t)
        if not prov:
            continue
        def premier(motifs):
            for m in motifs:
                r = re.search(m, t)
                if r:
                    return int(r.group(1))
            return None
        adm = premier([r"(\d+)\s+nouvelles?\s+admissions?", r"(\d+)\s+nouveaux\s+cas\s+suspects\s+ont"
                       r"\s+été\s+admis", r"(\d+)\s+nouvelle\s+admission"])
        sor = premier([r"(\d+)\s+sorties?"])
        if sor is None and "sans aucune sortie" in t:
            sor = 0
        gue = premier([r"dont\s+(\d+)\s+guérisons?", r"\((\d+)\s+guérisons?"])
        if gue is None and "toutes des guérisons" in t:
            gue = sor
        if gue is None and "sans aucune guérison" in t:
            gue = 0
        iso = premier([r"(\d+)\s+patients(?:\s+\([^)]*\))?\s+(?:demeuraient|sont)\s+en\s+isolement",
                       r"(\d+)\s+cas\s+suspects\s+sont\s+isolés"])
        sortie.append(dict(sitrep_no=no, date_rapport=d, province=prov, admissions=adm, sorties=sor,
                           guerisons=gue, patients_isolement=iso, texte_source=t))
    return sortie


def lire_contacts(paras, no, d):
    sortie, total = [], {}
    for p in paras:
        if p["section"] != "Surveillance":
            continue
        t = p["texte"]
        if "contacts en cours de suivi" not in t:
            continue
        m = re.search(r"Parmi?s?\s+(" + NUM + r")\s+contacts en cours de suivi.*?(" + NUM +
                      r")\s+ont été vus", t)
        if m:
            total = dict(contacts_suivis=entier(m.group(1)), contacts_vus=entier(m.group(2)))
        m = re.search(r"proportion de suivi de\s+(" + PCT + r")\s*%", t)
        if m:
            total["taux_publie"] = nombre_fr(m.group(1))
        for m in re.finditer(r"(" + PCT + r")\s*%\s*\((" + NUM + r")\s*/\s*(" + NUM + r")\)\s*"
                             r"(?:au|à la|à l['’]|à|en|dans le)\s*" + RE_PROV, t):
            sortie.append(dict(sitrep_no=no, date_rapport=d, province=norm_province(m.group(4)),
                               statut_donnees="OK", contacts_suivis=entier(m.group(3)),
                               contacts_vus=entier(m.group(2)), taux_publie=nombre_fr(m.group(1))))
        nd = re.search(r"([^.;]*)n['’]ont pas transmis[^.;]*", t)
        if nd:
            for m in re.finditer(RE_PROV, nd.group(0)):
                prov = norm_province(m.group(1))
                if prov and prov not in {s["province"] for s in sortie}:
                    sortie.append(dict(sitrep_no=no, date_rapport=d, province=prov,
                                       statut_donnees="ND", contacts_suivis=None, contacts_vus=None,
                                       taux_publie=None))
    return sortie, total


def chiffres_alertes_texte(paras) -> dict:
    r = {}
    for p in paras:
        if p["section"] != "Surveillance":
            continue
        t = p["texte"]
        m = re.search(r"vérifiées\s*\((" + PCT + r")\s*%\)", t)
        if m:
            r["pct_verifiees_texte"] = nombre_fr(m.group(1))
        m = re.search(r"(" + NUM + r")\s*\((" + PCT + r")\s*%\)\s*(?:ont été\s+)?validées", t)
        if m:
            r["validees_texte"], r["pct_validees_texte"] = entier(m.group(1)), nombre_fr(m.group(2))
        m = re.search(r"(" + NUM + r")\s+alertes ont été (?:rapportées|vérifiées)", t)
        if m:
            r["alertes_texte"] = entier(m.group(1))
    return r


# --------------------------------------------------------------------------
# Vérificateur générique des pourcentages écrits dans le texte
# --------------------------------------------------------------------------

MOTIFS_PCT = [
    # 84,2 % (315/374)   |  84,2 % pour les lits confirmés (169/209)
    ("pct_puis_fraction", r"(" + PCT + r")\s?%[^()%\d]{0,45}\(\s*(" + NUM + r")\s*/\s*(" + NUM + r")\s*\)"),
    # (146/165 ; 88,5 %)
    ("fraction_puis_pct", r"\(\s*(" + NUM + r")\s*/\s*(" + NUM + r")\s*;\s*(" + PCT + r")\s?%\s*\)"),
    # (36,1 % ; 13/36)
    ("pct_point_virgule_fraction", r"\(\s*(" + PCT + r")\s?%\s*;\s*(" + NUM + r")\s*/\s*(" + NUM + r")\s*\)"),
    # 123 feedbacks sur 151 (81,5 %)
    ("a_sur_b_pct", r"(?<![\d,])(" + NUM + r")\s+(?:[^\d()%;]{0,40}?\s)?sur\s+(?:[^\d()%;]{0,20}?)(" + NUM +
     r")(?:\s+[^\d()%;]{0,45}?)?\s*\(\s*(" + PCT + r")\s?%\s*\)"),
    # (10/36) (27,8%)
    ("fraction_puis_pct_parentheses", r"\(\s*(" + NUM + r")\s*/\s*(" + NUM + r")\s*\)\s*\(\s*(" + PCT + r")\s?%\s*\)"),
    # 207/436 AS soit 46%
    ("fraction_soit_pct", r"(" + NUM + r")\s*(?:[A-Za-zÀ-ÿ ]{0,25})/\s*(" + NUM + r")\s*(?:[A-Za-zÀ-ÿ' ]{0,40})"
     r"soit\s+(" + PCT + r")\s?%"),
]


def verifier_pourcentages(texte: str, source: str, ref: str, tolerance: float) -> list[dict]:
    sortie = []
    for nom, motif in MOTIFS_PCT:
        for m in re.finditer(motif, texte):
            g = m.groups()
            if nom in ("pct_puis_fraction", "pct_point_virgule_fraction"):
                pct, a, b = nombre_fr(g[0]), nombre_fr(g[1]), nombre_fr(g[2])
            else:
                a, b, pct = nombre_fr(g[0]), nombre_fr(g[1]), nombre_fr(g[2])
            if a is None or not b or pct is None:
                continue
            recalc = round(100 * a / b, 1)
            ecart = round(pct - recalc, 2)
            statut = "OK" if abs(ecart) <= tolerance + 0.05 else "ÉCART"
            inverse = round(100 * b / a, 1) if a else None
            note = ""
            if statut == "ÉCART" and inverse is not None and abs(inverse - pct) <= 1:
                note = "numérateur et dénominateur inversés"
            debut = max(0, m.start() - 60)
            sortie.append(dict(source=source, reference=ref, motif=nom, extrait=texte[debut:m.end() + 5],
                               pct_publie=pct, numerateur=a, denominateur=b, pct_recalcule=recalc,
                               ecart_points=ecart, statut=statut, note=note))
    # dédoublonnage (un même passage peut correspondre à deux motifs)
    vus, uniques = set(), []
    for s in sortie:
        k = (s["numerateur"], s["denominateur"], s["pct_publie"])
        if k not in vus:
            vus.add(k)
            uniques.append(s)
    return uniques


# --------------------------------------------------------------------------
# Programme principal
# --------------------------------------------------------------------------

def extraire(fichier: Path, cfg: dict):
    docx_path = vers_docx(fichier, cfg["chemins"]["intermediaire"], cfg.get("libreoffice", "soffice"))
    doc = docx.Document(docx_path)
    boites = zones_de_texte(doc)
    tout_texte = "\n".join(boites) + "\n" + "\n".join(p.text for p in doc.paragraphs)

    m = re.search(r"SitRep N°\s*(\d+)\s*/MVE-BDBV/(\d{2})/(\d{2})/(\d{4})", tout_texte)
    if not m:
        raise ValueError(f"Numéro de SitRep introuvable dans {fichier.name}")
    no = int(m.group(1))
    d = date(int(m.group(4)), int(m.group(3)), int(m.group(2)))
    mp = re.search(r"Date de publication\s*:\s*(\d{1,2}\s+\w+\s+\d{4})", tout_texte)

    tabs = trouver_tableaux(doc)
    manquants = {"t1", "t2", "t3", "t4"} - set(tabs)
    if manquants:
        LOG.warning("SitRep %s : tableaux non trouvés %s", no, sorted(manquants))
    paras = decoupe_sections(doc)
    kpi = kpi_encadres(boites)
    contacts, contacts_total = lire_contacts(paras, no, d)

    meta = dict(sitrep_no=no, date_rapport=d, date_publication=date_fr(mp.group(1)) if mp else None,
                fichier_source=fichier.name, date_extraction=datetime.now().date())
    meta.update({k: v for k, v in kpi.items() if k in [c[0] for c in SCHEMA["sitrep_meta"]["colonnes"]]})

    controle = dict(sitrep_no=no, totaux_t2=lire_t2_totaux(tabs.get("t2", [])),
                    contacts_total=contacts_total, alertes_texte=chiffres_alertes_texte(paras),
                    gueris_24h_detail=kpi.get("gueris_24h_detail"))

    tol = cfg["seuils"]["tolerance_pourcentage"]
    pcts = verifier_pourcentages(tout_texte, fichier.name, f"SitRep {no}", tol)
    for t in doc.tables:
        for l in lignes_tableau(t):
            pcts += verifier_pourcentages(" | ".join(l), fichier.name, f"SitRep {no} (tableau)", tol)

    return dict(
        meta=meta,
        epi_province=lire_t1(tabs.get("t1", []), no, d),
        epi_zs=lire_t2(tabs.get("t2", []), no, d),
        alertes_province=lire_t3(tabs.get("t3", []), no, d),
        poe_province=lire_t4(tabs.get("t4", []), no, d),
        contacts_province=contacts,
        labo_province=lire_labo(paras, no, d),
        pec_province=lire_pec(paras, no, d),
        sitrep_texte=[dict(sitrep_no=no, **p) for p in paras],
        pourcentages=pcts,
        controle=controle,
    )


def main(cfg: dict | None = None, fichiers: list[Path] | None = None):
    cfg = cfg or charger_config()
    dossier = cfg["chemins"]["entree_sitreps"]
    fichiers = fichiers or sorted([p for p in dossier.iterdir() if p.suffix.lower() in (".doc", ".docx")
                                   and not p.name.startswith("~$")])
    base = cfg["chemins"]["base"]
    cles = {f: SCHEMA[f]["cle"] for f in SCHEMA}
    acc = {f: [] for f in ["sitrep_meta", "epi_province", "epi_zs", "alertes_province", "poe_province",
                           "contacts_province", "labo_province", "pec_province", "sitrep_texte"]}
    pcts, controles = [], []
    if not fichiers:
        LOG.warning("Aucun SitRep national trouvé dans %s : étape ignorée, la base reste inchangée.", dossier)
        return
    for f in fichiers:
        LOG.info("Extraction de %s", f.name)
        try:
            r = extraire(f, cfg)
        except Exception as e:  # un fichier illisible ne bloque pas les autres
            LOG.error("SitRep ignoré (%s) : %s", f.name, e)
            continue
        acc["sitrep_meta"].append(r["meta"])
        for k in acc:
            if k != "sitrep_meta":
                acc[k].extend(r[k])
        pcts.extend(r["pourcentages"])
        controles.append(r["controle"])
        LOG.info("  SitRep %s du %s : %d ZS, %d provinces, %d pourcentages vérifiés",
                 r["meta"]["sitrep_no"], r["meta"]["date_rapport"], len(r["epi_zs"]),
                 len(r["epi_province"]), len(r["pourcentages"]))

    tables = {}
    for feuille, lignes in acc.items():
        nouveau = pd.DataFrame(lignes).reindex(columns=[c[0] for c in SCHEMA[feuille]["colonnes"]])
        existant = lire_table(base, feuille) if base.exists() else None
        tables[feuille] = fusionner(existant, nouveau, cles[feuille]).sort_values(cles[feuille])
    ecrire_tables(base, tables, cfg["chemins"].get("canevas_base"))

    inter = cfg["chemins"]["intermediaire"]
    pd.DataFrame(pcts).to_csv(inter / "pourcentages_texte_sitreps.csv", index=False, encoding="utf-8-sig")
    pd.to_pickle(controles, inter / "controles_sitreps.pkl")
    LOG.info("Base mise à jour : %s", base)
    return tables


if __name__ == "__main__":
    args = [Path(a) for a in sys.argv[1:]]
    main(fichiers=args or None)

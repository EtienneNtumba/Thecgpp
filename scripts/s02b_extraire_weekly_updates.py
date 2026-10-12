"""
s02b_extraire_weekly_updates.py - Weekly Updates CGPP (rapports hebdomadaires au bailleur, en anglais) vers la base.

Chaque Weekly Update contient :
  1. le tableau épidémiologique national par province, repris d'un SitRep national (n°, date) :
     il donne des points de la série nationale entre deux SitReps traités (feuille historique) ;
  2. un tableau consolidé des données opérationnelles des IP (RECO actifs, sensibilisation, alertes,
     décès prélevés, EDS, contacts, résistances) par partenaire et province : il alimente la feuille
     saisie_hebdo avec les identifiants de la matrice maîtresse (numérateur et dénominateur) ;
  3. des puces narratives (réalisations, contraintes, actions, autres épidémies) : elles sont
     conservées pour le contrôle qualité ; les chiffres utiles en sont codés à la main dans
     data/ref/weekly_updates_activites.csv et data/ref/autres_epidemies.csv (relus par l'équipe SI).

Sorties : base (historique, saisie_hebdo, autres_epidemies) et fichiers intermédiaires pour le contrôle
qualité (weekly_national.csv, weekly_operations.csv, weekly_puces.csv).

Usage : python scripts/s02b_extraire_weekly_updates.py
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

import docx
import pandas as pd
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from commun import SCHEMA, charger_config, cle, colonnes, ecrire_tables, fusionner, journal, lire_table, norm_province  # noqa: E402

LOG = journal("weekly")
MOIS_EN = {m: i for i, m in enumerate(["january", "february", "march", "april", "may", "june", "july", "august",
                                         "september", "october", "november", "december"], start=1)}
MOIS_EN.update({"sept": 9, "sep": 9, "oct": 10, "nov": 11, "dec": 12, "aug": 8, "jul": 7, "jun": 6})

# Lignes du tableau consolidé -> (indicateur, désagrégation, rôle). Un libellé peut alimenter deux indicateurs.
LIGNES_OPERATIONS = {
    "numberofreportsreceived": [("SURV-29", "rapports SBC reçus", "num")],
    "expectedrecos": [("SURV-29", "RECO actifs / RECO attendus", "den")],
    "recoswhoworked": [("SURV-29", "RECO actifs / RECO attendus", "num")],
    "peoplesensitizedcrec": [("CREC-03", "personnes sensibilisées", "num")],
    "sickalertsdetected": [("SURV-05", "alertes (malades) détectées", "num"),
                           ("SURV-07", "alertes notifiées / alertes détectées", "den")],
    "sickalertsreported": [("SURV-07", "alertes notifiées / alertes détectées", "num")],
    "communitydeathsswabbed": [("SURV-30", "décès communautaires prélevés", "num")],
    "safeanddignifiedburials": [("CREC-17", "EDS réalisés", "num")],
    "contactsunderfollowupcumulative": [("SURV-01", "contacts suivis (cumul)", "num"),
                                        ("SURV-02", "contacts vus / contacts suivis (cumul)", "den")],
    "contactsactuallyseen": [("SURV-02", "contacts vus / contacts suivis (cumul)", "num")],
    "resistancesidentified": [("CREC-14", "résistances levées / résistances identifiées", "den")],
    "resistancesresolved": [("CREC-14", "résistances levées / résistances identifiées", "num")],
}


def nombre(txt):
    """'6,155' -> 6155 ; '4 731' -> 4731 ; '46.0%' -> 46.0 ; '97,30%' -> 97.3 ; '28/36' -> (28, 36)."""
    t = str(txt).strip().replace(" ", " ").replace(" ", " ")
    if not t:
        return None
    if "/" in t:
        a, b = t.split("/", 1)
        return (nombre(a), nombre(b))
    if t.endswith("%"):
        t = t[:-1].strip().replace(" ", "")
        return float(t.replace(",", ".")) if t else None
    t = t.replace(" ", "")
    if re.fullmatch(r"\d{1,3}(,\d{3})+", t):
        return int(t.replace(",", ""))
    if re.fullmatch(r"\d+", t):
        return int(t)
    if re.fullmatch(r"\d+[.,]\d+", t):
        return float(t.replace(",", "."))
    return None


def date_titre(texte: str) -> date | None:
    m = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+"
                  r"(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})", texte)
    return date(int(m.group(3)), MOIS_EN[m.group(1).lower()], int(m.group(2))) if m else None


def texte_cellule(tc) -> str:
    return " ".join("".join(x.text or "" for x in p.iter(qn("w:t"))) for p in tc.iter(qn("w:p"))).strip()


def lire_weekly(fichier: Path) -> dict:
    doc = docx.Document(fichier)
    corps = doc.element.body
    elements, section = [], "En-tête"
    for el in corps.iterchildren():
        if el.tag == qn("w:p"):
            txt = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
            if not txt:
                continue
            ppr = el.find(qn("w:pPr"))
            puce = ppr is not None and ppr.find(qn("w:numPr")) is not None
            titre = (not puce and len(txt) < 75 and not txt.startswith(("-", "Source")) and ":" not in txt[:-1])
            if titre:
                section = re.sub(r"^\s*(\d+\.|[A-Z]\.)\s*", "", txt).rstrip(".: ")
            elements.append(dict(type="titre" if titre else ("puce" if puce else "texte"), section=section, texte=txt))
        elif el.tag == qn("w:tbl"):
            lignes = [[texte_cellule(tc) for tc in tr.iter(qn("w:tc"))] for tr in el.iter(qn("w:tr"))]
            elements.append(dict(type="tableau", section=section, lignes=lignes))
    tout = "\n".join(e.get("texte", "") for e in elements)
    d_wu = date_titre(elements[0]["texte"]) if elements else None
    if d_wu is None:
        raise ValueError(f"Date du Weekly Update introuvable dans le titre de {fichier.name}")

    # 1. tableau national et SitRep de référence
    national, ref = [], None
    m = re.search(r"SitRep\s*N°\s*(\d+)\s*/MVE-BDBV/(\d{2})/(\d{2})/(\d{4})", tout)
    if m:
        ref = dict(sitrep_no=int(m.group(1)), date_sitrep=date(int(m.group(4)), int(m.group(3)), int(m.group(2))))
    for e in elements:
        if e["type"] == "tableau" and e["lignes"] and "province" in cle(e["lignes"][0][0]):
            for lg in e["lignes"][1:]:
                if len(lg) < 6:
                    continue
                prov = "Total" if cle(lg[0]) == "total" else norm_province(lg[0])
                zs = nombre(lg[5])
                national.append(dict(date_weekly=d_wu, sitrep_no=ref and ref["sitrep_no"],
                                     date_sitrep=ref and ref["date_sitrep"], province=prov or lg[0],
                                     nouveaux_cas_24h=nombre(lg[1]), cas_cumul=nombre(lg[2]), deces_cumul=nombre(lg[3]),
                                     letalite_publiee=nombre(lg[4]),
                                     zs_touchees=zs[0] if isinstance(zs, tuple) else None,
                                     zs_total=zs[1] if isinstance(zs, tuple) else None))
    resume = {}
    mr = re.search(r"reported\s+([\d,]+)\s+confirmed cases and\s+([\d,]+)\s+confirmed deaths.*?Fatality Rate of\s+"
                   r"([\d.]+)%", tout, flags=re.S)
    if mr:
        resume = dict(cas=nombre(mr.group(1)), deces=nombre(mr.group(2)), letalite=float(mr.group(3)))
    mg = re.search(r"cumulative\s+([\d.,]+)\s+people have recovered", tout)
    resume["gueris_brut"] = mg.group(1) if mg else None
    resume["gueris"] = nombre(mg.group(1)) if mg and re.fullmatch(r"[1-9]\d{0,2}(,\d{3})*|\d+", mg.group(1)) else None

    # 2. tableau consolidé des données opérationnelles des IP
    operations, periode = [], None
    for e in elements:
        mp = re.search(r"\((Sept?|September|Oct|October|Aug|August)\.?\s+(\d{1,2})\s*[-–]\s*(\d{1,2}),\s*(\d{4})\)",
                       e.get("texte", ""))
        if mp:
            mois = MOIS_EN[mp.group(1).lower()]
            periode = (date(int(mp.group(4)), mois, int(mp.group(2))), date(int(mp.group(4)), mois, int(mp.group(3))))
        if e["type"] == "tableau" and e["lignes"] and cle(e["lignes"][0][0]) == "indicator":
            entetes = e["lignes"][0]
            colonnes_ip = []
            for j, h in enumerate(entetes[1:], start=1):
                mm = re.match(r"\s*([A-Za-z]+)\s*\(([^)]+)\)", h)
                if mm:
                    ip = {"WVI": "WV", "WV": "WV"}.get(mm.group(1).upper(), mm.group(1).upper())
                    colonnes_ip.append((j, ip, norm_province(mm.group(2)) or mm.group(2)))
                elif "total" in cle(h):
                    colonnes_ip.append((j, "Total", "Total"))
            for lg in e["lignes"][1:]:
                for j, ip, prov in colonnes_ip:
                    if j < len(lg):
                        operations.append(dict(date_weekly=d_wu, debut=periode[0] if periode else d_wu,
                                               fin=periode[1] if periode else d_wu, ligne=lg[0].strip(),
                                               cle_ligne=re.sub(r"[^a-z0-9]", "", cle(lg[0])), ip=ip, province=prov, valeur=nombre(lg[j]),
                                               brut=lg[j]))
    puces = [dict(date_weekly=d_wu, section=e["section"], type=e["type"], texte=e["texte"])
             for e in elements if e["type"] in ("puce", "texte")]
    return dict(date=d_wu, reference=ref, national=national, resume=resume, operations=operations, puces=puces)


def vers_historique(w: dict) -> list[dict]:
    lignes = []
    ref = w["reference"]
    if not ref:
        return lignes
    src = f"SitRep n° {ref['sitrep_no']} du {ref['date_sitrep']:%d/%m/%Y}, via Weekly Update CGPP du {w['date']:%d/%m/%Y}"
    for r in w["national"]:
        terr = "National" if r["province"] == "Total" else r["province"]
        for ind in ("cas_cumul", "deces_cumul", "nouveaux_cas_24h", "zs_touchees", "letalite_publiee"):
            if r.get(ind) is not None:
                lignes.append(dict(date_reference=pd.Timestamp(ref["date_sitrep"]), source=src, territoire=terr,
                                   indicateur=ind, valeur=float(r[ind]),
                                   commentaire=f"sur {r['zs_total']} ZS" if ind == "zs_touchees" and r.get("zs_total")
                                   else None))
    if w["resume"].get("gueris") is not None:
        lignes.append(dict(date_reference=pd.Timestamp(ref["date_sitrep"]), source=src, territoire="National",
                           indicateur="gueris_cumul", valeur=float(w["resume"]["gueris"]), commentaire=None))
    return lignes


def vers_saisie(w: dict, libelles: dict) -> list[dict]:
    lignes = []
    for r in w["operations"]:
        if r["ip"] == "Total" or r["valeur"] is None or isinstance(r["valeur"], tuple):
            continue
        for id_, desag, role in LIGNES_OPERATIONS.get(r["cle_ligne"], []):
            d0, d1 = r["debut"], r["fin"]
            iso = d0.isocalendar()
            lignes.append(dict(annee=d0.year, periode=f"S{iso.week} ({d0:%d}-{d1:%d/%m})", semaine_iso=iso.week,
                               date_debut=pd.Timestamp(d0), date_fin=pd.Timestamp(d1), ip=r["ip"], province=r["province"],
                               zone_sante=None, id_indicateur=id_, libelle_indicateur=libelles.get(id_),
                               desagregation=desag, role=role, valeur_brute=r["valeur"],
                               source_saisie=f"Weekly Update CGPP du {w['date']:%d/%m/%Y} (tableau consolidé)",
                               saisi_par="s02b (extraction automatique)", date_saisie=pd.Timestamp(w["date"]),
                               commentaire="Province entière (ZS non ventilées)"))
    if not lignes:
        return []
    df = pd.DataFrame(lignes)
    cles = ["annee", "periode", "semaine_iso", "date_debut", "date_fin", "ip", "province", "zone_sante",
            "id_indicateur", "libelle_indicateur", "desagregation", "source_saisie", "saisi_par", "date_saisie",
            "commentaire"]
    df["zone_sante"] = df["zone_sante"].fillna("")
    sortie = (df.pivot_table(index=cles, columns="role", values="valeur_brute", aggfunc="sum")
              .reset_index().rename(columns={"num": "numerateur", "den": "denominateur"}))
    sortie["zone_sante"] = sortie["zone_sante"].replace("", None)
    if "denominateur" not in sortie:
        sortie["denominateur"] = None
    sortie["valeur"] = [round(100 * a / b, 1) if pd.notna(b) and b else a
                        for a, b in zip(sortie.numerateur, sortie.denominateur)]
    return sortie.to_dict(orient="records")


def main(cfg: dict | None = None):
    cfg = cfg or charger_config()
    base = cfg["chemins"]["base"]
    dossier = cfg["chemins"].get("entree_weekly")
    fichiers = sorted(p for p in dossier.iterdir() if p.suffix.lower() == ".docx" and not p.name.startswith("~$")) \
        if dossier and dossier.exists() else []
    if not fichiers:
        LOG.warning("Aucun Weekly Update trouvé dans %s : étape ignorée.", dossier)
        return
    ref_ind = lire_table(base, "ref_indicateurs")
    libelles = dict(zip(ref_ind.id_indicateur, ref_ind.libelle_fr))
    hist, saisie, nat, ops, puces, resumes = [], [], [], [], [], []
    for f in fichiers:
        try:
            w = lire_weekly(f)
        except Exception as e:  # un fichier illisible ne bloque pas les autres
            LOG.error("Weekly Update ignoré (%s) : %s", f.name, e)
            continue
        hist += vers_historique(w)
        saisie += vers_saisie(w, libelles)
        nat += w["national"]
        ops += w["operations"]
        puces += w["puces"]
        resumes.append(dict(date_weekly=w["date"], fichier=f.name, **(w["reference"] or {}), **w["resume"]))
        LOG.info("Weekly Update du %s : SitRep de référence n° %s, %d provinces, %d valeurs opérationnelles, %d puces",
                 f"{w['date']:%d/%m/%Y}", w["reference"] and w["reference"]["sitrep_no"], len(w["national"]) - 1,
                 len(w["operations"]), len(w["puces"]))

    ref = cfg["chemins"]["referentiels"]
    activites = pd.read_csv(ref / "weekly_updates_activites.csv", parse_dates=["date_debut", "date_fin", "date_saisie"]) \
        if (ref / "weekly_updates_activites.csv").exists() else pd.DataFrame()
    if len(activites):
        activites["libelle_indicateur"] = activites.id_indicateur.map(libelles)
        activites["valeur"] = activites.numerateur
    autres = pd.read_csv(ref / "autres_epidemies.csv") if (ref / "autres_epidemies.csv").exists() \
        else pd.DataFrame(columns=colonnes("autres_epidemies"))

    nouvelles = pd.concat([pd.DataFrame(saisie), activites.drop(columns=["extrait_source"], errors="ignore")],
                          ignore_index=True) if len(activites) else pd.DataFrame(saisie)
    tables = {
        "historique": fusionner(lire_table(base, "historique"), pd.DataFrame(hist), SCHEMA["historique"]["cle"]),
        "saisie_hebdo": fusionner(lire_table(base, "saisie_hebdo"), nouvelles.reindex(columns=colonnes("saisie_hebdo")),
                                  SCHEMA["saisie_hebdo"]["cle"]),
        "autres_epidemies": fusionner(lire_table(base, "autres_epidemies"), autres, SCHEMA["autres_epidemies"]["cle"]),
    }
    for k, v in tables.items():
        tables[k] = v.reindex(columns=colonnes(k))
    ecrire_tables(base, tables, cfg["chemins"].get("canevas_base"))

    inter = cfg["chemins"]["intermediaire"]
    pd.DataFrame(nat).to_csv(inter / "weekly_national.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(ops).to_csv(inter / "weekly_operations.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(puces).to_csv(inter / "weekly_puces.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(resumes).to_csv(inter / "weekly_resumes.csv", index=False, encoding="utf-8-sig")
    LOG.info("Weekly Updates : %d points de série nationale, %d lignes de saisie (%d codées à la main), "
             "%d lignes « autres épidémies »", len(hist), len(nouvelles), len(activites), len(autres))
    return tables


if __name__ == "__main__":
    main()

"""
s02_extraire_indicateurs.py - Référentiels et données de programme CGPP vers la base Excel.

1. Référentiels : partenaires, zones de santé, semaines ISO, matrice maîtresse des indicateurs.
2. Classeur hebdomadaire CGPP (une feuille par IP, colonnes semaine x ZS) -> format long
   (feuille saisie_hebdo : une ligne = un indicateur, une ZS, une période).
3. Points de référence (historique), réunions et points d'action initiaux.
4. Fichiers intermédiaires pour le contrôle qualité : complétude du classeur et formules
   de la feuille compilée.

Usage : python scripts/s02_extraire_indicateurs.py
"""
from __future__ import annotations

import re
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from commun import (CODES_PROVINCE, SCHEMA, charger_config, cible_numerique, cle, colonnes,  # noqa: E402
                    ecrire_tables, fusionner, journal, lire_table, norm_zs)
from referentiel_indicateurs import CLASSEUR_VERS_ID, matrice  # noqa: E402

LOG = journal("indicateurs")
DEBUT_REFERENCE = (date(2026, 5, 16), date(2026, 8, 30))   # colonne « Wk 20 to Wk 35 » du classeur


def semaine_iso(annee: int, semaine: int) -> tuple[date, date]:
    lundi = date.fromisocalendar(annee, semaine, 1)
    return lundi, lundi + timedelta(days=6)


def ref_semaines(annee: int = 2026) -> pd.DataFrame:
    lignes = []
    for s in range(1, date(annee, 12, 28).isocalendar()[1] + 1):
        d0, d1 = semaine_iso(annee, s)
        lignes.append(dict(annee=annee, semaine_iso=s, date_debut=d0, date_fin=d1,
                           libelle=f"S{s:02d} ({d0:%d/%m} au {d1:%d/%m/%Y})"))
    return pd.DataFrame(lignes)


def ref_zones(cfg: dict, base: Path) -> pd.DataFrame:
    zs = lire_table(base, "epi_zs")
    zs = zs[~zs["est_a_ventiler"].fillna(False).astype(bool)]
    touchees = zs.groupby(["province", "zone_sante"], as_index=False)["cas_cumul"].max()
    lignes = {}
    for r in touchees.itertuples():
        lignes[(r.province, r.zone_sante)] = dict(province=r.province, zone_sante=r.zone_sante,
                                                   source="SitRep national", cas=r.cas_cumul)
    for nom, info in cfg["zones_cgpp"].items():
        k = (info["province"], nom)
        lignes.setdefault(k, dict(province=info["province"], zone_sante=nom, source="Classeur CGPP", cas=0))
        lignes[k]["ip_cgpp"] = info["ip"]
    sortie = []
    for v in lignes.values():
        ip = v.get("ip_cgpp")
        statut = "hors zone CGPP" if not ip else ("intervention" if (v.get("cas") or 0) > 0 else "préparation")
        sortie.append(dict(province=v["province"], zone_sante=v["zone_sante"], code_zs=None, ip_cgpp=ip,
                           statut_cgpp=statut, population=None, source=v["source"]))
    return pd.DataFrame(sortie).sort_values(["province", "zone_sante"])


def lire_classeur(cfg: dict, referentiel_zs: dict, semaine_reference: int):
    chemin = cfg["chemins"]["classeur_indicateurs"]
    wb = openpyxl.load_workbook(chemin, data_only=True)
    lignes, completude = [], []
    libelles = {cle(k): v for k, v in CLASSEUR_VERS_ID.items()}
    for nom_feuille in wb.sheetnames:
        m = re.match(r"Indicator Framework-(WV|ADRA|IMC|IRC)$", nom_feuille)
        if not m:
            continue
        ip = m.group(1)
        ws = wb[nom_feuille]
        # semaine de chaque colonne : libellé « Wk NN » le plus proche à gauche (ligne 2)
        semaine_col, courante = {}, None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(2, c).value
            if isinstance(v, str):
                mm = re.match(r"\s*Wk\s*(\d+)", v)
                if mm:
                    courante = int(mm.group(1))
            semaine_col[c] = courante
        entetes = {c: ws.cell(3, c).value for c in range(1, ws.max_column + 1)}
        cols_zs = {}
        for c, h in entetes.items():
            if not isinstance(h, str):
                continue
            mz = re.match(r"^(NK|SK|IT)\s+Hz\s+([A-Za-zÀ-ÿ\- ]+?)\d*$", h.strip())
            if mz and c >= 8:
                cols_zs[c] = (CODES_PROVINCE[mz.group(1)], norm_zs(mz.group(2).title(), referentiel_zs))
        for r in range(4, ws.max_row + 1):
            lib = ws.cell(r, 3).value
            if not lib:
                continue
            id_, desag = libelles.get(cle(lib), (None, ""))
            if id_ is None:
                LOG.warning("Indicateur non reconnu (%s, ligne %d) : %s", ip, r, lib)
                continue
            base_val = ws.cell(r, 7).value
            if isinstance(base_val, (int, float)):
                prov_ip = {v[0] for v in cols_zs.values()}
                lignes.append(dict(annee=2026, periode="S20-S35", semaine_iso=None,
                                   date_debut=DEBUT_REFERENCE[0], date_fin=DEBUT_REFERENCE[1], ip=ip,
                                   province=sorted(prov_ip)[0] if len(prov_ip) == 1 else "Plusieurs",
                                   zone_sante=None, id_indicateur=id_, libelle_indicateur=lib.strip(),
                                   desagregation=desag, numerateur=float(base_val), denominateur=None,
                                   valeur=float(base_val), source_saisie="Classeur hebdomadaire 2026",
                                   commentaire="Période agrégée non ventilée par ZS"))
            for c, (prov, zs) in cols_zs.items():
                v = ws.cell(r, c).value
                s = semaine_col.get(c)
                if isinstance(v, (int, float)) and s:
                    d0, d1 = semaine_iso(2026, s)
                    lignes.append(dict(annee=2026, periode=f"S{s}", semaine_iso=s, date_debut=d0, date_fin=d1,
                                       ip=ip, province=prov, zone_sante=zs, id_indicateur=id_,
                                       libelle_indicateur=lib.strip(), desagregation=desag,
                                       numerateur=float(v), denominateur=None, valeur=float(v),
                                       source_saisie="Classeur hebdomadaire 2026", commentaire=None))
        # complétude : une cellule IP x ZS x semaine est « renseignée » si au moins un indicateur a une valeur
        for c, (prov, zs) in cols_zs.items():
            s = semaine_col.get(c)
            if not s or s > semaine_reference:
                continue
            n = sum(isinstance(ws.cell(r, c).value, (int, float)) for r in range(4, ws.max_row + 1))
            completude.append(dict(ip=ip, province=prov, zone_sante=zs, semaine=s, indicateurs_renseignes=n,
                                   renseigne=n > 0))
    return pd.DataFrame(lignes), pd.DataFrame(completude)


def formules_compilees(cfg: dict) -> pd.DataFrame:
    wb = openpyxl.load_workbook(cfg["chemins"]["classeur_indicateurs"], data_only=False)
    ws = next(wb[n] for n in wb.sheetnames if "compil" in n.lower())
    sortie = []
    for r in range(4, ws.max_row + 1):
        lib = ws.cell(r, 3).value
        if not lib:
            continue
        f_zs = str(ws.cell(r, 8).value or "")
        f_tot = str(ws.cell(r, 5).value or "")
        sortie.append(dict(ligne=r, indicateur=lib.strip(), est_pourcentage=lib.strip().startswith("%"),
                           somme_semaines=f_zs.count("+") >= 2, somme_colonnes=f_tot.upper().startswith("=SUM"),
                           formule_zs=f_zs[:160], formule_total=f_tot[:160]))
    return pd.DataFrame(sortie)


def main(cfg: dict | None = None):
    cfg = cfg or charger_config()
    base = cfg["chemins"]["base"]
    ref = cfg["chemins"]["referentiels"]

    zones = ref_zones(cfg, base)
    referentiel_zs = {cle(z): z for z in zones["zone_sante"]}
    meta = lire_table(base, "sitrep_meta")
    derniere = pd.to_datetime(meta["date_rapport"]).max() if not meta.empty else pd.Timestamp.today()
    semaine_reference = int(derniere.isocalendar().week)

    saisie, completude = lire_classeur(cfg, referentiel_zs, semaine_reference)
    LOG.info("Classeur hebdomadaire : %d valeurs extraites ; %d cellules IP x ZS x semaine attendues (S36 à S%d)",
             len(saisie), len(completude), semaine_reference)

    partenaires = pd.DataFrame([dict(code_ip=k, nom=v, role="Partenaire de mise en oeuvre",
                                     provinces=", ".join(sorted({z["province"] for z in cfg["zones_cgpp"].values()
                                                                 if z["ip"] == k})) or "À préciser")
                                for k, v in cfg["partenaires"].items()])
    historique = pd.read_csv(ref / "historique_manuel.csv", parse_dates=["date_reference"])
    reunions = pd.read_csv(ref / "reunions_initiales.csv", parse_dates=["date"])
    actions = pd.read_csv(ref / "points_action_initiaux.csv")
    notes = pd.read_csv(ref / "sitrep_notes_initiales.csv", parse_dates=["date_rapport"]) \
        if (ref / "sitrep_notes_initiales.csv").exists() else pd.DataFrame(columns=colonnes("sitrep_notes"))

    tables = {
        "ref_partenaires": partenaires,
        "ref_zones": zones,
        "ref_semaines": ref_semaines(),
        "ref_indicateurs": matrice().assign(
            cible_dos_num=lambda d: d.cible_dos.map(cible_numerique),
            cible_t5_num=lambda d: d.cible_tranche5.map(cible_numerique)),
        "saisie_hebdo": fusionner(lire_table(base, "saisie_hebdo"), saisie, SCHEMA["saisie_hebdo"]["cle"]),
        "historique": fusionner(lire_table(base, "historique"), historique, SCHEMA["historique"]["cle"]),
        "reunions": fusionner(lire_table(base, "reunions"), reunions, ["id_reunion"]),
        "points_action": fusionner(lire_table(base, "points_action"), actions, ["id_action"]),
        "sitrep_notes": fusionner(lire_table(base, "sitrep_notes"), notes, SCHEMA["sitrep_notes"]["cle"]),
    }
    for k, v in tables.items():
        tables[k] = v.reindex(columns=colonnes(k))
    ecrire_tables(base, tables, cfg["chemins"].get("canevas_base"))

    inter = cfg["chemins"]["intermediaire"]
    completude.to_csv(inter / "completude_classeur.csv", index=False, encoding="utf-8-sig")
    formules_compilees(cfg).to_csv(inter / "formules_classeur.csv", index=False, encoding="utf-8-sig")
    LOG.info("Référentiels et saisies écrits dans %s", base)
    return tables


if __name__ == "__main__":
    main()

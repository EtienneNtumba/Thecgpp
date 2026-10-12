"""
s04_analyses.py - Analyses épidémiologiques et programmatiques à partir de la base Excel.

Sorties : sorties/analyses/analyses.xlsx (tableaux), sorties/figures/*.svg|png (graphiques),
          sorties/analyses/resultats.json (chiffres clés réutilisés par le SitRep et le rapport).

Usage   : python scripts/s04_analyses.py
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import matplotlib
import pandas as pd

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from commun import charger_config, journal, lire_table  # noqa: E402

LOG = journal("analyses")

# Palette (validée : daltonisme et contraste ; libellés directs systématiques)
BLEU, ORANGE, VERT_EAU, JAUNE = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"
GRIS_ACCENT, ENCRE, ENCRE_2, ENCRE_3 = "#c3c2b7", "#0b0b0b", "#52514e", "#898781"
GRILLE, SURFACE = "#e1e0d9", "#ffffff"
BLEU_CLAIR = "#86b6ef"
COULEUR_IP = {"WV": BLEU, "ADRA": ORANGE, "IMC": VERT_EAU, "IRC": JAUNE}


def style():
    police = "Inter" if any("Inter" in f.name for f in font_manager.fontManager.ttflist) else "DejaVu Sans"
    plt.rcParams.update({
        "font.family": police, "font.size": 8.5, "axes.edgecolor": GRIS_ACCENT, "axes.linewidth": 0.8,
        "axes.labelcolor": ENCRE_2, "xtick.color": ENCRE_3, "ytick.color": ENCRE_2, "axes.titlesize": 9.5,
        "axes.titleweight": "semibold", "axes.titlecolor": ENCRE, "axes.titlelocation": "left",
        "axes.titlepad": 10, "svg.fonttype": "none", "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
        "xtick.major.size": 0, "ytick.major.size": 0, "legend.frameon": False, "legend.fontsize": 8,
    })


def fr(x, dec=0):
    if x is None or pd.isna(x):
        return "ND"
    s = f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")
    return s


def axes_propres(ax, grille="x"):
    for c in ("top", "right", "left"):
        ax.spines[c].set_visible(False)
    ax.spines["bottom"].set_color(GRIS_ACCENT)
    if grille:
        ax.grid(axis=grille, color=GRILLE, linewidth=0.7)
        ax.set_axisbelow(True)


def format_fr(ax, axe="x"):
    from matplotlib.ticker import FuncFormatter
    fmt = FuncFormatter(lambda v, _: f"{v:,.0f}".replace(",", "\u202f"))
    (ax.xaxis if axe == "x" else ax.yaxis).set_major_formatter(fmt)


def enregistrer(fig, dossier: Path, nom: str):
    for ext in ("svg", "png"):
        fig.savefig(dossier / f"{nom}.{ext}", dpi=220, bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)


# --------------------------------------------------------------------------
# Calculs
# --------------------------------------------------------------------------

def calculer(cfg):
    base = cfg["chemins"]["base"]
    meta = lire_table(base, "sitrep_meta").sort_values("sitrep_no")
    prov = lire_table(base, "epi_province")
    zs = lire_table(base, "epi_zs")
    contacts = lire_table(base, "contacts_province")
    alertes = lire_table(base, "alertes_province")
    labo = lire_table(base, "labo_province")
    hist = lire_table(base, "historique")
    saisie = lire_table(base, "saisie_hebdo")
    ref_ind = lire_table(base, "ref_indicateurs")
    zones_cgpp = cfg["zones_cgpp"]
    nos = list(meta.sitrep_no.astype(int))
    n1, n0 = nos[-1], nos[-2] if len(nos) > 1 else nos[-1]
    m1 = meta[meta.sitrep_no == n1].iloc[0]
    m0 = meta[meta.sitrep_no == n0].iloc[0]
    jours = (pd.to_datetime(m1.date_rapport) - pd.to_datetime(m0.date_rapport)).days or 1

    h = hist[hist.source.str.contains("SitRep 085", na=False)]
    h_nat = h[h.territoire == "National"].set_index("indicateur")["valeur"]
    h_prov = h[h.territoire != "National"].pivot_table(index="territoire", columns="indicateur", values="valeur")
    d085 = pd.to_datetime(h.date_reference.iloc[0]) if len(h) else None
    jours_085 = (pd.to_datetime(m1.date_rapport) - d085).days if d085 is not None else None

    # --- national
    national = dict(
        sitrep=n1, date=str(pd.to_datetime(m1.date_rapport).date()), sitrep_prec=n0,
        date_prec=str(pd.to_datetime(m0.date_rapport).date()), jours=jours,
        cas=int(m1.cas_cumul), deces=int(m1.deces_cumul), gueris=int(m1.gueris_cumul),
        letalite=round(100 * m1.deces_cumul / m1.cas_cumul, 1), cas_prec=int(m0.cas_cumul),
        deces_prec=int(m0.deces_cumul), nouveaux_periode=int(m1.cas_cumul - m0.cas_cumul),
        deces_periode=int(m1.deces_cumul - m0.deces_cumul),
        moyenne_jour=round((m1.cas_cumul - m0.cas_cumul) / jours, 1), nouveaux_24h=int(m1.nouveaux_cas_24h),
        nouveaux_24h_prec=int(m0.nouveaux_cas_24h), isolement=int(m1.patients_isolement),
        zs_touchees=int(prov[prov.sitrep_no == n1].zs_touchees.sum()),
        zs_total=int(prov[prov.sitrep_no == n1].zs_total.sum()),
        provinces=int((prov[prov.sitrep_no == n1].cas_cumul > 0).sum()),
        cas_085=int(h_nat.get("cas_cumul", 0)), deces_085=int(h_nat.get("deces_cumul", 0)),
        jours_085=jours_085, zs_085=int(h_nat.get("zs_touchees", 0)),
        moyenne_jour_085=round((m1.cas_cumul - h_nat.get("cas_cumul", 0)) / jours_085, 1) if jours_085 else None,
        contacts_attendus_085=int(h_nat.get("contacts_attendus", 0)), contacts_vus_085=int(h_nat.get("contacts_vus", 0)),
        alertes_085=int(h_nat.get("alertes_24h", 0)), alertes_validees_085=int(h_nat.get("alertes_validees_24h", 0)),
        suspects_investigues_085=int(h_nat.get("suspects_investigues_24h", 0)),
    )
    # --- provinces
    p1 = prov[prov.sitrep_no == n1].set_index("province")
    p0 = prov[prov.sitrep_no == n0].set_index("province")
    t_prov = pd.DataFrame({
        "cas_085": h_prov.get("cas_cumul"), "cas_prec": p0.cas_cumul, "cas": p1.cas_cumul,
        "deces": p1.deces_cumul, "nouveaux_24h": p1.nouveaux_cas_24h, "zs_touchees": p1.zs_touchees,
        "zs_total": p1.zs_total}).reindex(p1.index)
    t_prov["hausse_periode"] = t_prov.cas - t_prov.cas_prec
    t_prov["part_cas_pct"] = (100 * t_prov.cas / t_prov.cas.sum()).round(1)
    t_prov["part_085_pct"] = (100 * t_prov.cas_085 / national["cas_085"]).round(1) if national["cas_085"] else None
    t_prov["facteur_depuis_085"] = (t_prov.cas / t_prov.cas_085).round(1)
    t_prov["letalite"] = (100 * t_prov.deces / t_prov.cas).round(1)
    t_prov = t_prov.sort_values("cas", ascending=False)

    # --- ZS : hausse sur la période, ZS CGPP
    z1 = zs[(zs.sitrep_no == n1) & ~zs.est_a_ventiler.fillna(False).astype(bool)].set_index(["province", "zone_sante"])
    z0 = zs[(zs.sitrep_no == n0) & ~zs.est_a_ventiler.fillna(False).astype(bool)].set_index(["province", "zone_sante"])
    t_zs = pd.DataFrame({"cas_prec": z0.cas_cumul, "cas": z1.cas_cumul, "deces": z1.deces_cumul,
                         "nouveaux_24h": z1.nouveaux_cas_24h}).reindex(z1.index)
    t_zs["cas_prec"] = t_zs.cas_prec.fillna(0)
    t_zs["hausse"] = t_zs.cas - t_zs.cas_prec
    t_zs["hausse_pct"] = (100 * t_zs.hausse / t_zs.cas_prec.where(t_zs.cas_prec > 0)).round(1)
    t_zs["letalite"] = (100 * t_zs.deces / t_zs.cas).round(1)
    t_zs = t_zs.reset_index()
    t_zs["ip_cgpp"] = t_zs.zone_sante.map({k: v["ip"] for k, v in zones_cgpp.items()})
    seuil_pct, seuil_n = cfg["seuils"]["foyer_hausse_pct"], cfg["seuils"]["foyer_hausse_min_cas"]
    t_zs["foyer_actif"] = (t_zs.hausse >= seuil_n) & ((t_zs.hausse_pct >= seuil_pct) | (t_zs.hausse >= 10))
    t_zs = t_zs.sort_values(["hausse", "cas"], ascending=False)

    cg = pd.DataFrame([dict(zone_sante=k, province=v["province"], ip=v["ip"]) for k, v in zones_cgpp.items()])
    cg = cg.merge(t_zs[["zone_sante", "cas_prec", "cas", "deces", "hausse", "nouveaux_24h", "letalite"]],
                  on="zone_sante", how="left").fillna({"cas_prec": 0, "cas": 0, "deces": 0, "hausse": 0})
    cg = cg.sort_values("cas", ascending=False)
    cgpp = dict(cas=int(cg.cas.sum()), deces=int(cg.deces.sum()), hausse=int(cg.hausse.sum()),
                part_cas_pct=round(100 * cg.cas.sum() / national["cas"], 1),
                part_hausse_pct=round(100 * cg.hausse.sum() / national["nouveaux_periode"], 1),
                zs_touchees=int((cg.cas > 0).sum()), zs_total=len(cg),
                zs_sans_cas=", ".join(cg[cg.cas == 0].zone_sante))
    hausse_pos = t_zs[t_zs.hausse > 0]
    couverture = dict(zs_en_hausse=len(hausse_pos), zs_en_hausse_cgpp=int(hausse_pos.ip_cgpp.notna().sum()),
                      cas_hausse_total=int(hausse_pos.hausse.sum()),
                      cas_hausse_cgpp=int(hausse_pos[hausse_pos.ip_cgpp.notna()].hausse.sum()),
                      top15_cgpp=int(t_zs.head(15).ip_cgpp.notna().sum()),
                      foyers_hors_cgpp=", ".join(t_zs[(t_zs.foyer_actif) & (t_zs.ip_cgpp.isna())].zone_sante))

    # --- décès communautaires parmi les nouveaux cas
    dc = []
    for no in (n0, n1):
        z = zs[zs.sitrep_no == no]
        for p, g in z.groupby("province"):
            nv = g.nouveaux_cas_24h.fillna(0).sum()
            if nv > 0:
                dc.append(dict(sitrep=no, province=p, nouveaux_cas=int(nv),
                               deces_communautaires=int(g.deces_communautaires_24h.fillna(0).sum()),
                               deces_cte=int(g.deces_cte_24h.fillna(0).sum())))
        dc.append(dict(sitrep=no, province="Total", nouveaux_cas=int(z.nouveaux_cas_24h.fillna(0).sum()),
                       deces_communautaires=int(z.deces_communautaires_24h.fillna(0).sum()),
                       deces_cte=int(z.deces_cte_24h.fillna(0).sum())))
    t_dc = pd.DataFrame(dc)
    t_dc["part_post_mortem_pct"] = (100 * t_dc.deces_communautaires / t_dc.nouveaux_cas).round(1)
    t_dc["part_comm_parmi_deces_pct"] = (100 * t_dc.deces_communautaires /
                                          (t_dc.deces_communautaires + t_dc.deces_cte)).round(1)

    # --- contacts, alertes, labo
    t_ct = contacts[contacts.sitrep_no.isin([n0, n1])].copy()
    t_ct["taux"] = (100 * t_ct.contacts_vus / t_ct.contacts_suivis).round(1)
    t_al = alertes[alertes.sitrep_no.isin([n0, n1])].copy()
    t_al["taux_validation"] = (100 * (t_al.validees_vivants + t_al.validees_deces) / t_al.recues_total).round(1)
    t_al["taux_transfert"] = (100 * t_al.suspects_transferes / t_al.suspects_investigues).round(1)
    t_lb = labo[labo.sitrep_no.isin([n0, n1])].copy()
    t_lb["positivite"] = (100 * t_lb.positifs / t_lb.echantillons).round(1)

    # --- programme CGPP : réalisations et cibles
    def cible(id_, col):
        v = ref_ind.loc[ref_ind.id_indicateur == id_, col]
        return v.iloc[0] if len(v) else None
    from commun import nombre_fr
    real = []
    classeur = saisie[saisie.source_saisie.astype(str).str.contains("Classeur", na=False)]   # réalisations du classeur
    for id_ in ["SURV-05", "SURV-09"]:
        r = pd.to_numeric(classeur[classeur.id_indicateur == id_]["numerateur"], errors="coerce").sum()
        cd, ct = nombre_fr(cible(id_, "cible_dos")), nombre_fr(cible(id_, "cible_tranche5"))
        real.append(dict(id=id_, libelle=cible(id_, "libelle_fr"), realise=r, cible_dos=cd, cible_t5=ct,
                         pct_dos=round(100 * r / cd, 1) if cd else None, pct_t5=round(100 * r / ct, 1) if ct else None))
    t_real = pd.DataFrame(real)
    comp = pd.read_csv(cfg["chemins"]["intermediaire"] / "completude_classeur.csv")
    alertes_wv = saisie[(saisie.id_indicateur == "SURV-05") & saisie.semaine_iso.notna()]
    t_alertes_wv = alertes_wv.pivot_table(index="zone_sante", columns="periode", values="numerateur", aggfunc="sum")

    # --- série nationale et rythme par intervalle ; données opérationnelles des IP
    serie = serie_nationale(base)
    t_iv = intervalles(serie)
    d1 = pd.to_datetime(m1.date_rapport)
    dates_serie = sorted(serie.date.unique())
    d_ref = next((d for d in dates_serie if pd.Timestamp(d) >= d1 - pd.Timedelta(days=15)), dates_serie[0])
    ref_q = serie[serie.date == d_ref].set_index("territoire")
    der = serie[serie.date == d1].set_index("territoire")
    quinzaine = dict(date_ref=str(pd.Timestamp(d_ref).date()), sitrep_ref=int(ref_q.sitrep_no.iloc[0]),
                     jours=int((d1 - pd.Timestamp(d_ref)).days),
                     cas=int(der.loc["National", "cas_cumul"] - ref_q.loc["National", "cas_cumul"]),
                     deces=int(der.loc["National", "deces_cumul"] - ref_q.loc["National", "deces_cumul"]),
                     zs_ref=int(ref_q.loc["National", "zs_touchees"]),
                     provinces={p: int(der.loc[p, "cas_cumul"] - (ref_q.loc[p, "cas_cumul"] if p in ref_q.index else 0))
                                for p in der.index if p != "National"},
                     cas_ref={p: (int(ref_q.loc[p, "cas_cumul"]) if p in ref_q.index else 0) for p in der.index})
    t_ops, ops_meta = operations_ip(base)
    resultats = dict(national=national, cgpp=cgpp, couverture=couverture,
                     intervalles=t_iv.assign(debut=t_iv.debut.astype(str), fin=t_iv.fin.astype(str)).to_dict(orient="records"),
                     quinzaine=quinzaine, operations=dict(meta=ops_meta, lignes=t_ops.to_dict(orient="records")),
                     realisations=t_real.to_dict(orient="records"),
                     completude_classeur=round(100 * comp.renseigne.mean(), 1),
                     cellules_attendues=int(len(comp)), cellules_renseignees=int(comp.renseigne.sum()))
    tables = dict(provinces=t_prov.reset_index().rename(columns={"index": "province"}), zones=t_zs, zs_cgpp=cg,
                  deces_communautaires=t_dc, contacts=t_ct, alertes=t_al, laboratoire=t_lb, realisations=t_real,
                  completude_classeur=comp, alertes_wv=t_alertes_wv.reset_index(), serie_nationale=serie,
                  intervalles=t_iv, operations_ip=t_ops)
    return resultats, tables



# --------------------------------------------------------------------------
# Série nationale (SitReps traités + points repris des Weekly Updates et du plan DOS)
# --------------------------------------------------------------------------

def serie_nationale(base) -> pd.DataFrame:
    """Cumuls par date et par territoire (provinces et « National »), toutes sources confondues."""
    hist = lire_table(base, "historique")
    meta = lire_table(base, "sitrep_meta")
    ep = lire_table(base, "epi_province")
    lignes = []
    h = hist[hist.source.astype(str).str.contains("SitRep", na=False)]
    for r in h.itertuples():
        m = re.search(r"SitRep (?:n° )?0*(\d+)", str(r.source))
        lignes.append(dict(date=pd.to_datetime(r.date_reference), sitrep_no=int(m.group(1)) if m else None,
                           territoire=r.territoire, indicateur=r.indicateur, valeur=float(r.valeur)))
    dates = dict(zip(meta.sitrep_no, pd.to_datetime(meta.date_rapport)))
    for r in ep.itertuples():
        for ind in ("cas_cumul", "deces_cumul", "nouveaux_cas_24h", "zs_touchees"):
            lignes.append(dict(date=dates[r.sitrep_no], sitrep_no=int(r.sitrep_no), territoire=r.province,
                               indicateur=ind, valeur=float(getattr(r, ind))))
    for r in meta.itertuples():
        for ind in ("cas_cumul", "deces_cumul", "nouveaux_cas_24h", "gueris_cumul"):
            lignes.append(dict(date=pd.to_datetime(r.date_rapport), sitrep_no=int(r.sitrep_no), territoire="National",
                               indicateur=ind, valeur=float(getattr(r, ind))))
        zs = ep[ep.sitrep_no == r.sitrep_no].zs_touchees.sum()
        lignes.append(dict(date=pd.to_datetime(r.date_rapport), sitrep_no=int(r.sitrep_no), territoire="National",
                           indicateur="zs_touchees", valeur=float(zs)))
    s = pd.DataFrame(lignes).sort_values("date").drop_duplicates(["date", "territoire", "indicateur"], keep="last")
    s["sitrep_no"] = s.groupby("date").sitrep_no.transform(lambda x: x.dropna().iloc[0] if x.notna().any() else None)
    return (s.pivot_table(index=["date", "sitrep_no", "territoire"], columns="indicateur", values="valeur")
            .reset_index().sort_values(["date", "territoire"]))


def intervalles(serie: pd.DataFrame) -> pd.DataFrame:
    """Nouveaux cas et décès par intervalle entre deux points de la série, et moyennes par jour."""
    cas = serie.pivot_table(index="date", columns="territoire", values="cas_cumul")
    dec = serie.pivot_table(index="date", columns="territoire", values="deces_cumul")
    nos = serie.drop_duplicates("date").set_index("date").sitrep_no
    sortie = []
    for d0, d1 in zip(cas.index, cas.index[1:]):
        j = (d1 - d0).days
        ligne = dict(debut=d0, fin=d1, sitrep_debut=nos[d0], sitrep_fin=nos[d1], jours=j)
        for terr in cas.columns:
            if pd.isna(cas.loc[d1, terr]):
                continue
            avant = 0 if pd.isna(cas.loc[d0, terr]) else cas.loc[d0, terr]
            ligne[f"cas_{terr}"] = cas.loc[d1, terr] - avant
        ligne["deces_National"] = dec.loc[d1, "National"] - dec.loc[d0, "National"]
        ligne["cas_Autres provinces"] = (ligne["cas_National"] - ligne.get("cas_Ituri", 0) - ligne.get("cas_Nord-Kivu", 0))
        for k in [k for k in ligne if k.startswith(("cas_", "deces_"))]:
            ligne[k.replace("cas_", "jour_cas_").replace("deces_", "jour_deces_")] = round(ligne[k] / j, 1)
        sortie.append(ligne)
    return pd.DataFrame(sortie)


def operations_ip(base) -> tuple[pd.DataFrame, dict]:
    """Dernier tableau opérationnel consolidé des IP (Weekly Update) : valeurs par IP, total recalculé, taux."""
    s = lire_table(base, "saisie_hebdo")
    s = s[s.source_saisie.astype(str).str.contains("tableau consolidé", na=False)].copy()
    if s.empty:
        return pd.DataFrame(), {}
    s["date_fin"] = pd.to_datetime(s.date_fin)
    s = s[s.date_fin == s.date_fin.max()]
    for c in ("numerateur", "denominateur"):
        s[c] = pd.to_numeric(s[c], errors="coerce")
    s["colonne"] = s.ip + " " + s.province
    lignes = []
    for (id_, desag), g in s.groupby(["id_indicateur", "desagregation"], sort=False):
        ligne = dict(id_indicateur=id_, desagregation=desag)
        for r in g.itertuples():
            ligne[f"num|{r.colonne}"], ligne[f"den|{r.colonne}"] = r.numerateur, r.denominateur
        ligne["num|Total"], ligne["den|Total"] = g.numerateur.sum(), g.denominateur.sum(min_count=1)
        lignes.append(ligne)
    t = pd.DataFrame(lignes)
    meta = dict(debut=str(pd.to_datetime(s.date_debut).min().date()), fin=str(s.date_fin.max().date()),
                source=s.source_saisie.iloc[0], colonnes=sorted(s.colonne.unique(), key=lambda c: ("WV" in c, c)))
    return t, meta

# --------------------------------------------------------------------------
# Graphiques
# --------------------------------------------------------------------------

def fig_parts_provinces(t_prov, res, dossier):
    """Part des cas cumulés par province : SitRep 085 contre dernier SitRep (barres 100 % empilées)."""
    ordre = ["Ituri", "Nord-Kivu", "Haut-Uélé"]
    couleurs = [BLEU, ORANGE, VERT_EAU, JAUNE]
    lignes = []
    for etiq, col, tot in [(f"7 août\n(SitRep 085)", "cas_085", res["national"]["cas_085"]),
                           (f"10 octobre\n(SitRep {res['national']['sitrep']})", "cas", res["national"]["cas"])]:
        parts = [100 * (t_prov.set_index("province").loc[p, col] or 0) / tot for p in ordre]
        parts.append(100 - sum(parts))
        lignes.append((etiq, parts))
    fig, ax = plt.subplots(figsize=(6.6, 1.9))
    for i, (etiq, parts) in enumerate(lignes):
        gauche = 0
        for j, (v, c) in enumerate(zip(parts, couleurs)):
            ax.barh(i, v, left=gauche, height=0.52, color=c, edgecolor=SURFACE, linewidth=1.5)
            if v >= 9:   # libellé interne seulement s'il tient dans le segment
                ax.text(gauche + v / 2, i, f"{v:.1f} %".replace(".", ","), ha="center", va="center", fontsize=8,
                        color="white" if c in (BLEU, ORANGE) else ENCRE, fontweight="semibold")
            gauche += v
    ax.set_yticks([0, 1], [l[0] for l in lignes])
    ax.invert_yaxis()
    ax.set_xlim(0, 100)
    ax.set_xticks([0, 25, 50, 75, 100], ["0 %", "25 %", "50 %", "75 %", "100 %"])
    axes_propres(ax, grille=None)
    ax.spines["bottom"].set_visible(False)
    ax.tick_params(axis="x", labelsize=7.5)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=c, label=l) for c, l in zip(couleurs, ordre + ["Autres provinces"])],
              ncol=4, loc="upper left", bbox_to_anchor=(0, -0.22), fontsize=8, handlelength=1, columnspacing=1.4)
    ax.set_title("Répartition des cas confirmés cumulés par province")
    enregistrer(fig, dossier, "f1_parts_provinces")


def fig_dumbbell_provinces(t_prov, res, dossier):
    t = t_prov[t_prov.cas_085.notna() & (t_prov.cas_085 > 0)].sort_values("cas")
    fig, ax = plt.subplots(figsize=(6.6, 2.3))
    for i, (p, r) in enumerate(t.iterrows()):
        ax.plot([r.cas_085, r.cas], [i, i], color=GRIS_ACCENT, linewidth=2, solid_capstyle="round", zorder=1)
        ax.scatter([r.cas_085], [i], s=38, color=BLEU_CLAIR, edgecolor=SURFACE, linewidth=1.5, zorder=2)
        ax.scatter([r.cas], [i], s=44, color=BLEU, edgecolor=SURFACE, linewidth=1.5, zorder=3)
        fac = f"x {r.facteur_depuis_085:.1f}".replace(".", ",") if r.cas_085 else ""
        ax.text(r.cas * 1.18, i, f"{fr(r.cas)}  ({fac})", va="center", fontsize=8, color=ENCRE)
    ax.set_xscale("log")
    ax.set_xlim(1.5, 30000)
    ax.set_xticks([3, 10, 30, 100, 300, 1000, 3000, 10000], ["3", "10", "30", "100", "300", "1 000", "3 000", "10 000"])
    ax.set_yticks(range(len(t)), t.province if "province" in t else t.index)
    axes_propres(ax)
    ax.tick_params(axis="x", labelsize=7.5)
    ax.scatter([], [], s=38, color=BLEU_CLAIR, label="7 août (SitRep 085)")
    ax.scatter([], [], s=44, color=BLEU, label=f"10 octobre (SitRep {res['national']['sitrep']})")
    ax.legend(loc="lower right", fontsize=7.5, handletextpad=0.3)
    ax.set_title("Cas confirmés cumulés par province (échelle logarithmique)")
    enregistrer(fig, dossier, "f2_croissance_provinces")


def fig_hausse_zs(t_zs, res, dossier, n=15):
    t = t_zs.head(n).iloc[::-1]
    fig, ax = plt.subplots(figsize=(6.6, 3.6))
    couleurs = [BLEU if pd.notna(ip) else GRIS_ACCENT for ip in t.ip_cgpp]
    ax.barh(range(len(t)), t.hausse, height=0.55, color=couleurs)
    for i, r in enumerate(t.itertuples()):
        pct = f" ({r.hausse_pct:.0f} %)".replace(".", ",") if pd.notna(r.hausse_pct) else ""
        ax.text(r.hausse + 0.4, i, f"+{int(r.hausse)}{pct}", va="center", fontsize=7.8, color=ENCRE)
    etiquettes = [f"{r.zone_sante} ({r.province})" + (f"  [{r.ip_cgpp}]" if pd.notna(r.ip_cgpp) else "")
                  for r in t.itertuples()]
    ax.set_yticks(range(len(t)), etiquettes, fontsize=7.8)
    ax.set_xlim(0, t.hausse.max() * 1.25)
    axes_propres(ax)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=BLEU, label="ZS du classeur CGPP"), Patch(color=GRIS_ACCENT, label="Autres ZS")],
              loc="lower right", fontsize=7.5)
    ax.set_title(f"Hausse des cas confirmés cumulés par ZS, du {pd.to_datetime(res['national']['date_prec']):%d/%m} "
                 f"au {pd.to_datetime(res['national']['date']):%d/%m} (15 premières)")
    enregistrer(fig, dossier, "f3_hausse_zs")


def fig_contacts(t_ct, res, cfg, dossier):
    n0, n1 = res["national"]["sitrep_prec"], res["national"]["sitrep"]
    provs = ["Ituri", "Nord-Kivu", "Haut-Uélé", "Tshopo", "Bas-Uélé", "Sud-Ubangi"]
    fig, ax = plt.subplots(figsize=(6.6, 2.5))
    for i, p in enumerate(provs):
        a = t_ct[(t_ct.sitrep_no == n0) & (t_ct.province == p)]
        b = t_ct[(t_ct.sitrep_no == n1) & (t_ct.province == p)]
        va = a.taux.iloc[0] if len(a) and pd.notna(a.taux.iloc[0]) else None
        vb = b.taux.iloc[0] if len(b) and pd.notna(b.taux.iloc[0]) else None
        if va is not None and vb is not None:
            ax.plot([va, vb], [i, i], color=GRIS_ACCENT, linewidth=2, zorder=1)
        if va is not None:
            ax.scatter([va], [i], s=38, color=BLEU_CLAIR, edgecolor=SURFACE, linewidth=1.5, zorder=2)
        if vb is not None:
            ax.scatter([vb], [i], s=44, color=BLEU, edgecolor=SURFACE, linewidth=1.5, zorder=3)
            ax.text(max(vb, va or 0) + 0.8, i, f"{vb:.1f} %".replace(".", ","), va="center", fontsize=7.8)
        else:
            ax.text((va or 70) + 0.8, i, f"{va:.1f} %  (10/10 : ND)".replace(".", ","), va="center", fontsize=7.8,
                    color=ENCRE_2)
    s_nat, s_cgpp = cfg["seuils"]["suivi_contacts_national"], cfg["seuils"]["suivi_contacts_cible_cgpp"]
    ax.axvline(s_nat, color=ENCRE_3, linewidth=1.0, label=f"Seuil national {s_nat:.0f} %")
    ax.axvline(s_cgpp, color=GRIS_ACCENT, linewidth=1.0, label=f"Cible CGPP {s_cgpp:.0f} %")
    ax.set_yticks(range(len(provs)), provs)
    ax.invert_yaxis()
    ax.set_xlim(70, 101)
    ax.set_xticks(range(70, 101, 5), [f"{x} %" for x in range(70, 101, 5)])
    axes_propres(ax)
    ax.scatter([], [], s=38, color=BLEU_CLAIR, label=f"7 octobre (SitRep {n0})")
    ax.scatter([], [], s=44, color=BLEU, label=f"10 octobre (SitRep {n1})")
    ax.legend(loc="upper left", fontsize=7.5, bbox_to_anchor=(0, -0.16), ncol=4, handletextpad=0.4,
              columnspacing=1.2)
    ax.set_title("Contacts vus dans les 24 h, par province")
    enregistrer(fig, dossier, "f4_suivi_contacts")


def fig_zs_cgpp(cg, res, dossier):
    t = cg.sort_values("cas")
    fig, ax = plt.subplots(figsize=(6.6, 3.0))
    ax.barh(range(len(t)), t.cas, height=0.55, color=[COULEUR_IP[i] for i in t.ip])
    for i, r in enumerate(t.itertuples()):
        lib = f"{fr(r.cas)}" + (f"  (+{int(r.hausse)} en {res['national']['jours']} j)" if r.hausse > 0 else "")
        ax.text(r.cas + 12, i, lib, va="center", fontsize=7.8, color=ENCRE)
    ax.set_yticks(range(len(t)), [f"{r.zone_sante} ({r.province})" for r in t.itertuples()], fontsize=7.8)
    ax.set_xlim(0, t.cas.max() * 1.35)
    axes_propres(ax)
    format_fr(ax)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=COULEUR_IP[k], label=k) for k in ["WV", "ADRA", "IMC"]], loc="lower right",
              fontsize=7.5, title="Partenaire", title_fontsize=7.5)
    ax.set_title(f"Cas confirmés cumulés dans les 12 ZS du classeur CGPP au "
                 f"{pd.to_datetime(res['national']['date']):%d/%m/%Y}")
    enregistrer(fig, dossier, "f5_zs_cgpp")


def fig_rythme(t_iv, dossier):
    """Nouveaux cas confirmés par jour, moyenne de chaque intervalle entre deux SitReps (barres empilées)."""
    mois = {8: "août", 9: "sept.", 10: "oct.", 11: "nov.", 12: "déc."}
    def jour(d):
        d = pd.to_datetime(d)
        return f"{d.day} {mois.get(d.month, d.month)}"
    etiquettes = []
    for r in t_iv.itertuples():
        a, b = pd.to_datetime(r.debut), pd.to_datetime(r.fin)
        deb = f"{a.day}" if a.month == b.month else jour(a)
        etiquettes.append(f"{deb} au {jour(b)}\n({r.jours} j, SitReps {int(r.sitrep_debut):03d}-{int(r.sitrep_fin)})")
    fig, ax = plt.subplots(figsize=(6.6, 2.75))
    x = list(range(len(t_iv)))
    bas = [0.0] * len(t_iv)
    for terr, coul in [("Ituri", BLEU), ("Nord-Kivu", ORANGE), ("Autres provinces", VERT_EAU)]:
        v = [float(r.get(f"jour_cas_{terr}", 0) or 0) for r in t_iv.to_dict(orient="records")]
        ax.bar(x, v, bottom=bas, width=0.56, color=coul, edgecolor=SURFACE, linewidth=1.2, label=terr)
        for i, (b0, vv) in enumerate(zip(bas, v)):
            if vv >= 8:
                ax.text(i, b0 + vv / 2, fr(vv, 1), ha="center", va="center", fontsize=7.6,
                        color="white" if coul in (BLEU, ORANGE) else ENCRE, fontweight="semibold")
        bas = [b0 + vv for b0, vv in zip(bas, v)]
    for i, r in enumerate(t_iv.itertuples()):
        ax.text(i, bas[i] + 1.2, f"{fr(r.jour_cas_National, 1)} cas/j", ha="center", va="bottom", fontsize=8,
                color=ENCRE, fontweight="semibold")
    ax.set_xticks(x, etiquettes, fontsize=7.4)
    ax.set_ylim(0, max(bas) * 1.18)
    ax.set_ylabel("cas par jour", fontsize=7.6)
    axes_propres(ax, grille="y")
    ax.legend(loc="upper right", fontsize=7.5, ncol=3, handlelength=1, columnspacing=1.2)
    ax.set_title("Nouveaux cas confirmés par jour, en moyenne sur chaque intervalle entre SitReps")
    enregistrer(fig, dossier, "f6_rythme")


def main(cfg: dict | None = None):
    cfg = cfg or charger_config()
    style()
    res, tables = calculer(cfg)
    dossier_fig = cfg["chemins"]["sorties_figures"]
    fig_parts_provinces(tables["provinces"], res, dossier_fig)
    fig_dumbbell_provinces(tables["provinces"].set_index("province"), res, dossier_fig)
    fig_hausse_zs(tables["zones"], res, dossier_fig)
    fig_contacts(tables["contacts"], res, cfg, dossier_fig)
    fig_zs_cgpp(tables["zs_cgpp"], res, dossier_fig)
    if len(tables["intervalles"]):
        fig_rythme(tables["intervalles"], dossier_fig)
    sortie = cfg["chemins"]["sorties_analyses"]
    with pd.ExcelWriter(sortie / "analyses.xlsx", engine="xlsxwriter") as xw:
        for nom, t in tables.items():
            t.to_excel(xw, sheet_name=nom[:31], index=False)
    with open(sortie / "resultats.json", "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=2, default=str)
    LOG.info("Analyses écrites dans %s ; figures dans %s", sortie, dossier_fig)
    return res, tables


if __name__ == "__main__":
    main()

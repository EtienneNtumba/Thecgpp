"""
generer_rapport.py - Produit le rapport PDF « Information stratégique CGPP, riposte MVE-17 ».

Lit la base Excel, les sorties du contrôle qualité et des analyses, la matrice des indicateurs,
puis assemble un document HTML (Jinja2) imprimé en PDF par Chromium (Playwright) :
  passe 1 : sommaire sans numéros, repères invisibles dans les titres, lecture des pages (pdfplumber) ;
  passe 2 : sommaire numéroté, contrôle que les titres n'ont pas bougé ;
  passe 3 : version finale sans repères, contrôle de chaque titre sur sa page ;
puis ajoute la page de garde, les signets, les liens du sommaire et la numérotation des pages (pypdf).

Usage : python outils/rapport/generer_rapport.py
"""
from __future__ import annotations

import io
import json
import logging
import re
import subprocess
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import pdfplumber
from jinja2 import Environment, FileSystemLoader
from markupsafe import Markup
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import Link

logging.getLogger("pdfminer").setLevel(logging.ERROR)

ICI = Path(__file__).resolve().parent
RACINE = ICI.parents[1]
sys.path.insert(0, str(RACINE / "scripts"))
from commun import SCHEMA, charger_config, lire_table  # noqa: E402
from s05_generer_sitrep import ecarts_publies, liste_fr  # noqa: E402

NBSP, NNBSP = " ", " "
DATE_DOC = date(2026, 10, 12)


def f(x, dec=0, signe=False):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "ND"
    s = f"{x:+,.{dec}f}" if signe else f"{x:,.{dec}f}"
    return s.replace(",", NNBSP).replace(".", ",")


def p(x, dec=1):
    return "ND" if x is None or (isinstance(x, float) and pd.isna(x)) else f(x, dec) + NBSP + "%"


def date_fr(d):
    mois = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre",
            "novembre", "décembre"]
    d = pd.to_datetime(d)
    return f"{d.day}{' ' if d.day != 1 else 'er '}{mois[d.month - 1]} {d.year}".replace("1er  ", "1er ")


def svg_en_ligne(chemin: Path) -> Markup:
    """Insère un SVG dans la page plutôt qu'en image : ses textes prennent alors la police Inter
    TrueType du document (incorporée proprement dans le PDF) au lieu d'une police système."""
    txt = chemin.read_text(encoding="utf-8")
    for motif in (r"<\?xml[^>]*\?>", r"<!DOCTYPE[^>]*>", r"<!--.*?-->", r"<metadata>.*?</metadata>"):
        txt = re.sub(motif, "", txt, flags=re.S)
    return Markup(txt.strip())


REGLES = {
    "SR-01": ("SitReps", "Somme des ZS égale au total de la province (cas et décès cumulés)"),
    "SR-02": ("SitReps", "Somme des provinces égale au chiffre de l'encadré national"),
    "SR-03": ("SitReps", "Létalité provinciale publiée égale à décès / cas"),
    "SR-04": ("SitReps", "Létalité par ZS recalculée (écart maximal 0,1 point)"),
    "SR-05": ("SitReps", "Nouveaux cas : ZS, province, total et faits saillants concordent"),
    "SR-06": ("SitReps", "Décès du jour : communautaires + intra-CTE = total, et concordance avec les faits saillants"),
    "SR-07": ("SitReps", "Nombre de ZS listées égal au nombre de ZS touchées du tableau 1"),
    "SR-08": ("SitReps", "Alertes : sommes internes, provinces non rapportantes, pourcentages du texte"),
    "SR-09": ("SitReps", "Contacts : sommes, taux recalculés, périmètre de l'encadré"),
    "SR-10": ("SitReps", "Laboratoire : échantillons et positifs concordent avec les faits saillants"),
    "SR-11": ("SitReps", "Chaque pourcentage écrit « x % (a/b) » ou « a sur b (x %) » est recalculé"),
    "SR-12": ("SitReps", "Patients en isolement : encadré égal à la somme des provinces"),
    "SR-13": ("SitReps", "Guéris du jour : résumé égal à la somme des guérisons par province"),
    "SR-14": ("SitReps", "Taux d'occupation : numérateur égal aux patients du jour (pas du SitRep précédent)"),
    "SR-15": ("SitReps", "Part des décès « à ventiler » inférieure à 5 %"),
    "SR-16": ("SitReps", "Compteur d'aires de santé mis à jour quand une ZS nouvelle est touchée"),
    "SR-17": ("SitReps", "Complétude de la courbe épidémique (date de début et statut vital)"),
    "SR-18": ("SitReps", "Date de la carte égale à la date du rapport"),
    "SR-19": ("SitReps", "PoE/PoC : valeur globale égale à la somme des provinces"),
    "SR-20": ("SitReps", "Ventilation « N patients (a confirmés et b suspects) » cohérente"),
    "TS-01": ("Entre SitReps", "Aucun cumul par ZS ne diminue d'un SitRep à l'autre"),
    "TS-02": ("Entre SitReps", "Hausse du cumul au moins égale aux nouveaux cas du dernier jour"),
    "TS-03": ("Entre SitReps", "Repérage des nouvelles ZS touchées"),
    "TS-04": ("Entre SitReps", "Comparaison du taux de suivi des contacts sur le même périmètre"),
    "CR-01": ("Réunions", "Pourcentages « soit x % » recalculés (et inversions repérées)"),
    "CR-02": ("Réunions", "Cumuls cités en réunion cohérents entre eux et avec les SitReps"),
    "CR-03": ("Réunions", "Dénominateurs (ZS par province) identiques à ceux du SitRep"),
    "CR-04": ("Réunions", "Chaque point d'action a un responsable et une échéance"),
    "CR-05": ("Réunions", "Liste de présence complète jointe"),
    "CR-06": ("Réunions", "Données personnelles masquées avant diffusion"),
    "IND-01": ("Classeur IP", "Complétude des cellules IP x ZS x semaine"),
    "IND-02": ("Classeur IP", "Aucun pourcentage additionné dans la compilation"),
    "IND-03": ("Classeur IP", "Une cible renseignée par indicateur"),
    "IND-04": ("Classeur IP", "Tous les IP des plans de travail présents"),
    "IND-05": ("Classeur IP", "Cas confirmés des ZS CGPP égaux aux chiffres du SitRep"),
    "IND-06": ("Classeur IP", "Données ventilées par ZS et par semaine"),
    "IND-07": ("Cadre global", "Intitulé et définition concordent"),
    "IND-08": ("Cadre global", "Une seule version du cadre, sans doublon"),
    "BAR-01": ("Matrice Barclay", "Identifiants uniques"),
    "BAR-02": ("Matrice Barclay", "Cibles dans un format numérique homogène"),
    "BAR-03": ("Matrice Barclay", "Aucun indicateur dupliqué"),
    "BAR-04": ("Matrice Barclay", "Résumé cohérent avec la matrice"),
    "BAR-05": ("Matrice Barclay", "Réalisations renseignées"),
    "BAR-06": ("Matrice Barclay", "Fréquence compatible avec la cible"),
    "PT-01": ("Plans", "Somme des activités égale au total du pilier"),
    "PT-02": ("Plans", "Somme des piliers égale au budget total"),
    "PT-03": ("Plans", "Détail des coûts égal au montant de l'activité"),
    "PT-04": ("Plans", "Une seule cible par extrant dans tout le document"),
    "PT-05": ("Plans", "Cibles DOS et Tranche 5 comparables"),
    "PT-06": ("Plans", "Budget Tranche 5 : somme des piliers égale au total"),
    "PT-07": ("Plans", "Quantité x coût unitaire plausible"),
    "SBC-01": ("Fiche SBC", "Chaque indicateur alimenté par la fiche a son champ"),
    "SBC-02": ("Fiche SBC", "Fenêtres temporelles identiques à celles des définitions"),
    "SBC-03": ("Fiche SBC", "Une valeur numérique par cellule"),
    "SBC-04": ("Fiche SBC", "Un seul canevas pour tous les IP"),
    "SBC-05": ("Fiche SBC", "Pas de feuille obsolète masquée"),
    "TS-05": ("Entre SitReps", "Série nationale : aucun cumul ne baisse d'un SitRep à l'autre (085 à 149)"),
    "TS-06": ("Entre SitReps", "Série nationale : somme des provinces égale au total à chaque date"),
    "WU-01": ("Weekly Updates", "Tableau national : somme des provinces égale à la ligne Total"),
    "WU-02": ("Weekly Updates", "Tableau national : létalité recalculée par province"),
    "WU-03": ("Weekly Updates", "Résumé exécutif concordant avec le tableau national"),
    "WU-04": ("Weekly Updates", "Tableau opérationnel : colonne Total égale à la somme des IP"),
    "WU-05": ("Weekly Updates", "Tableau opérationnel : taux de participation des RECO recalculé"),
    "WU-06": ("Weekly Updates", "Tous les partenaires présents dans le tableau opérationnel"),
    "WU-07": ("Weekly Updates", "Section « autres épidémies » mise à jour d'un rapport à l'autre"),
    "WU-08": ("Weekly Updates", "Pas de puce répétée dans une même section"),
    "WU-09": ("Weekly Updates", "Données opérationnelles de 7 jours au plus à la date du SitRep"),
    "WU-11": ("Weekly Updates", "Observations manuelles : chiffres et libellés incohérents (28/09)"),
    "WU-12": ("Weekly Updates", "Observations manuelles : définitions et cohérence texte-tableau (30/09)"),
}


def contexte():
    cfg = charger_config()
    base = cfg["chemins"]["base"]
    res = json.loads((cfg["chemins"]["sorties_analyses"] / "resultats.json").read_text(encoding="utf-8"))
    xa = cfg["chemins"]["sorties_analyses"] / "analyses.xlsx"
    T = {sh: pd.read_excel(xa, sheet_name=sh) for sh in pd.ExcelFile(xa).sheet_names}
    nat, cg, cv = res["national"], res["cgpp"], res["couverture"]
    qcf = sorted(cfg["chemins"]["sorties_qc"].glob("rapport_qc_*.xlsx"))[-1]
    tous = pd.read_excel(qcf, sheet_name="tous_les_controles")
    anom = pd.read_excel(qcf, sheet_name="anomalies")
    ref = lire_table(base, "ref_indicateurs")

    N = dict(cas=f(nat["cas"]), deces=f(nat["deces"]), gueris=f(nat["gueris"]), letalite=p(nat["letalite"]),
             cas_prec=f(nat["cas_prec"]), nouveaux_periode=f(nat["nouveaux_periode"]),
             deces_periode=f(nat["deces_periode"]), moyenne_jour=f(nat["moyenne_jour"], 1),
             moyenne_jour_085=f(nat["moyenne_jour_085"], 1), cas_085=f(nat["cas_085"]), deces_085=f(nat["deces_085"]),
             jours_085=nat["jours_085"], zs=nat["zs_touchees"], zs_total=nat["zs_total"], zs_085=nat["zs_085"],
             provinces=nat["provinces"], isolement=f(nat["isolement"]), nouveaux_24h=nat["nouveaux_24h"],
             facteur_085=f(nat["cas"] / nat["cas_085"], 1), sitrep=nat["sitrep"], sitrep_prec=nat["sitrep_prec"],
             jours=nat["jours"], contacts_085=f(nat["contacts_attendus_085"]),
             taux_contacts_085=p(100 * nat["contacts_vus_085"] / nat["contacts_attendus_085"]),
             alertes_085=f(nat["alertes_085"]),
             validation_085=p(100 * nat["alertes_validees_085"] / nat["alertes_085"]),
             investigation_085=p(100 * nat["suspects_investigues_085"] / nat["alertes_validees_085"]))
    C = dict(cas=f(cg["cas"]), deces=f(cg["deces"]), hausse=f(cg["hausse"]), part=p(cg["part_cas_pct"]),
             part_hausse=p(cg["part_hausse_pct"]), zs_touchees=cg["zs_touchees"], zs_total=cg["zs_total"],
             zs_sans_cas=cg["zs_sans_cas"], letalite=p(100 * cg["deces"] / cg["cas"]))
    V = dict(cv)

    # Provinces
    pv = T["provinces"]
    prov_rows = [dict(province=r.province, cas_085=f(r.cas_085) if pd.notna(r.cas_085) else "-",
                      cas_prec=f(r.cas_prec), cas=f(r.cas), hausse=f(r.hausse_periode, signe=True) if r.hausse_periode else "0",
                      part_085=p(r.part_085_pct) if pd.notna(r.part_085_pct) else "-", part=p(r.part_cas_pct),
                      facteur=("x " + f(r.facteur_depuis_085, 1)) if pd.notna(r.facteur_depuis_085) else "nouvelle",
                      letalite=p(r.letalite), zs=f"{int(r.zs_touchees)}/{int(r.zs_total)}") for r in pv.itertuples()]
    parts = {r.province: r for r in pv.itertuples()}
    # Décès communautaires
    dc = T["deces_communautaires"]
    dc_rows = [dict(sitrep=int(r.sitrep), province=r.province, nouveaux=int(r.nouveaux_cas), comm=int(r.deces_communautaires),
                    cte=int(r.deces_cte), part=p(r.part_post_mortem_pct), part_deces=p(r.part_comm_parmi_deces_pct)
                    if pd.notna(r.part_comm_parmi_deces_pct) else "-") for r in dc.itertuples()]
    dct = dc[dc.province == "Total"].set_index("sitrep")
    # Contacts
    ct = T["contacts"]
    ct_rows = []
    for prov in ["Ituri", "Nord-Kivu", "Haut-Uélé", "Tshopo", "Bas-Uélé", "Sud-Ubangi"]:
        a = ct[(ct.sitrep_no == nat["sitrep_prec"]) & (ct.province == prov)]
        b = ct[(ct.sitrep_no == nat["sitrep"]) & (ct.province == prov)]
        def fmt(x):
            if not len(x) or x.statut_donnees.iloc[0] == "ND":
                return "ND", "ND"
            r = x.iloc[0]
            return p(100 * r.contacts_vus / r.contacts_suivis), f"{f(r.contacts_vus)} / {f(r.contacts_suivis)}"
        ta, na = fmt(a)
        tb, nb = fmt(b)
        ct_rows.append(dict(province=prov, t146=ta, n146=na, t149=tb, n149=nb))
    c146 = ct[(ct.sitrep_no == nat["sitrep_prec"]) & (ct.statut_donnees == "OK")]
    tot_contacts_146 = f(c146.contacts_suivis.sum())
    # Alertes
    al = T["alertes"]
    al_rows = []
    for prov in ["Ituri", "Nord-Kivu", "Haut-Uélé", "Bas-Uélé", "Tshopo", "Sud-Ubangi"]:
        ligne = dict(province=prov)
        for no, k in [(nat["sitrep_prec"], "a"), (nat["sitrep"], "b")]:
            x = al[(al.sitrep_no == no) & (al.province == prov)]
            if not len(x) or x.statut_donnees.iloc[0] == "ND":
                ligne[k] = dict(total="ND", val="ND", trans="ND")
            else:
                r = x.iloc[0]
                ligne[k] = dict(total=f(r.recues_total), val=p(r.taux_validation), trans=p(r.taux_transfert)
                                if pd.notna(r.taux_transfert) else "-")
        al_rows.append(ligne)
    a146 = al[(al.sitrep_no == nat["sitrep_prec"])]
    val146 = p(100 * (a146.validees_vivants.sum() + a146.validees_deces.sum()) / a146.recues_total.sum())
    # Labo
    lb = T["laboratoire"]
    lab_rows = []
    for prov in ["Ituri", "Nord-Kivu", "Haut-Uélé", "Tshopo", "Bas-Uélé"]:
        ligne = dict(province=prov)
        for no, k in [(nat["sitrep_prec"], "a"), (nat["sitrep"], "b")]:
            x = lb[(lb.sitrep_no == no) & (lb.province == prov)]
            ligne[k] = "-" if not len(x) else f"{p(x.iloc[0].positivite)} ({int(x.iloc[0].positifs)}/{int(x.iloc[0].echantillons)})"
        lab_rows.append(ligne)
    # ZS
    zt = T["zones"]
    zs_top = [dict(zone=r.zone_sante, province=r.province, prec=f(r.cas_prec), cas=f(r.cas), hausse=f(r.hausse, signe=True),
                   pct=(f(r.hausse_pct, 0, signe=True) + NBSP + "%") if pd.notna(r.hausse_pct) else "-",
                   ip=r.ip_cgpp if pd.notna(r.ip_cgpp) else "", foyer=bool(r.foyer_actif)) for r in zt.head(15).itertuples()]
    zc = T["zs_cgpp"]
    zs_cgpp = [dict(zone=r.zone_sante, province=r.province, ip=r.ip, cas=f(r.cas), hausse=f(r.hausse, signe=True)
                    if r.hausse else "0", deces=f(r.deces), letalite=p(r.letalite) if r.cas else "-",
                    statut="touchée" if r.cas > 0 else "préparation") for r in zc.itertuples()]
    # Complétude du classeur
    comp = T["completude_classeur"]
    semaines = sorted(comp.semaine.unique())
    grille = []
    for (ip, zs), g in comp.groupby(["ip", "zone_sante"], sort=False):
        grille.append(dict(ip=ip, zone=zs, cases=[bool(g[g.semaine == s].renseigne.any()) for s in semaines]))
    grille.sort(key=lambda x: (["WV", "ADRA", "IMC"].index(x["ip"]), x["zone"]))
    realis = [dict(id=r["id"], libelle=r["libelle"], realise=f(r["realise"]), cible_dos=f(r["cible_dos"]),
                   cible_t5=f(r["cible_t5"]), pct_dos=p(r["pct_dos"]), pct_t5=p(r["pct_t5"]),
                   w_dos=min(100, r["pct_dos"]), w_t5=min(100, r["pct_t5"])) for r in res["realisations"]]

    # Contrôle qualité
    qc = dict(n=len(tous), ok=int((tous.statut == "OK").sum()), echecs=int((tous.statut == "ÉCHEC").sum()),
              verif=int((tous.statut == "À VÉRIFIER").sum()), info=int((tous.statut == "INFO").sum()),
              crit=int((anom.gravite == "Critique").sum()), maj=int((anom.gravite == "Majeure").sum()),
              mineures=int((anom.gravite == "Mineure").sum()), fichier=qcf.name)
    fam = []
    for fa, g in tous.groupby("famille"):
        a = anom[anom.famille == fa]
        fam.append(dict(famille=fa, n=len(g), ok=int((g.statut == "OK").sum()),
                        crit=int((a.gravite == "Critique").sum()), maj=int((a.gravite == "Majeure").sum()),
                        mineures=int((a.gravite == "Mineure").sum())))
    fam.sort(key=lambda x: -(x["crit"] * 100 + x["maj"] * 10 + x["mineures"]))
    def nettoie(x):
        return "" if pd.isna(x) else str(x)
    anomalies = [dict(id=r.id_regle, gravite=r.gravite, statut=r.statut, source=nettoie(r.source), objet=nettoie(r.objet),
                      attendu=nettoie(r.attendu), observe=nettoie(r.observe), reco=nettoie(r.recommandation))
                 for r in anom.itertuples()]
    regles = [dict(id=k, famille=v[0], description=v[1],
                   n=int((tous.id_regle == k).sum()), echecs=int(((tous.id_regle == k) & tous.statut.isin(["ÉCHEC", "À VÉRIFIER"])).sum()))
              for k, v in REGLES.items()]

    # Matrice des indicateurs
    ref["priorite"] = pd.to_numeric(ref["priorite"], errors="coerce")
    piliers = ["COORD", "SURV", "CREC", "PCI", "LOG", "MEL", "PREP", "DEPL"]
    noms_piliers = {"COORD": "Coordination", "SURV": "Surveillance", "CREC": "CREC (communication, engagement)",
                    "PCI": "PCI, EDS et référence", "LOG": "Logistique", "MEL": "MEL et conformité",
                    "PREP": "Préparation (Tranche 5)", "DEPL": "Déplacés (Tranche 5)"}
    mat_rows = []
    for pl in piliers:
        g = ref[ref.pilier == pl]
        mat_rows.append(dict(pilier=noms_piliers[pl], code=pl, existant=int((g.statut == "existant").sum()),
                             corriger=int((g.statut == "à corriger").sum()), nouveau=int((g.statut == "nouveau").sum()),
                             total=len(g)))
    mat_tot = dict(existant=sum(r["existant"] for r in mat_rows), corriger=sum(r["corriger"] for r in mat_rows),
                   nouveau=sum(r["nouveau"] for r in mat_rows), total=len(ref))
    t5_only = ref[ref.commentaire.fillna("").str.contains("Propre à la Tranche 5")]
    nouveaux = ref[(ref.statut == "nouveau") & ~ref.id_indicateur.isin(t5_only.id_indicateur)]
    def ligne_ind(r):
        return dict(id=r.id_indicateur, libelle=r.libelle_fr, type=r.type, num=nettoie(r.numerateur),
                    den=nettoie(r.denominateur), source=nettoie(r.source_donnees), champ=nettoie(r.champ_source),
                    freq=nettoie(r.frequence), seuil=nettoie(r.seuil_alerte), comm=nettoie(r.commentaire),
                    prio=int(r.priorite) if pd.notna(r.priorite) else "", cible_dos=nettoie(r.cible_dos),
                    cible_t5=nettoie(r.cible_tranche5), statut=r.statut, barclay=nettoie(r.id_barclay),
                    glob=nettoie(r.id_cadre_global))
    nouveaux_p1 = [ligne_ind(r) for r in nouveaux[nouveaux.priorite == 1].itertuples()]
    nouveaux_p23 = [ligne_ind(r) for r in nouveaux[nouveaux.priorite > 1].itertuples()]
    a_corriger = [ligne_ind(r) for r in ref[ref.statut == "à corriger"].itertuples()]
    annexe_ind = [ligne_ind(r) for r in ref.itertuples()]
    cibles = []
    from commun import cible_numerique
    for r in ref.itertuples():
        a, b = cible_numerique(r.cible_dos), cible_numerique(r.cible_tranche5)
        if a and b and r.type == "Nombre":
            cibles.append(dict(id=r.id_indicateur, libelle=r.libelle_fr, dos=f(a), t5=f(b), ratio="x " + f(b / a, 1)))
    cibles.sort(key=lambda x: -float(x["ratio"][2:].replace(",", ".").replace(NNBSP, "")))

    # Budgets
    sys.path.insert(0, str(RACINE / "scripts"))
    import docx as _docx
    d = _docx.Document(cfg["chemins"]["plan_dos"])
    lignes_act, totaux = {}, []
    from commun import nombre_fr
    for t in d.tables:
        for row in t.rows:
            cells = []
            for c in row.cells:
                tx = " ".join(c.text.split())
                if not cells or cells[-1] != tx:
                    cells.append(tx)
            if len(cells) >= 3 and re.match(r"^\d+\.\d+\.\d+$", cells[0]):
                k = cells[0].split(".")[0]
                mt = nombre_fr(cells[-1].replace("$", "").replace(",", ""))
                cle_ = (cells[0].rsplit(".", 1)[0], mt, cells[-2] if len(cells) >= 4 else "")
                lignes_act.setdefault(k, set()).add(cle_)
            if len(cells) >= 2 and cells[-2].strip() == "Total":
                totaux.append(nombre_fr(cells[-1].replace("$", "").replace(",", "")))
    noms_dos = ["Coordination et alignement des partenaires", "Surveillance (contacts, investigation, alertes)",
                "Communication sur les risques et engagement communautaire", "PCI, EDS et référence",
                "Logistique et appui opérationnel", "Gestion, MEL, conformité, finances, sécurité"]
    budget_dos = []
    for i, k in enumerate(sorted(lignes_act, key=int)):
        somme = sum(m or 0 for _, m, _ in lignes_act[k])
        aff = totaux[i] if i < len(totaux) else None
        budget_dos.append(dict(pilier=f"{k}. {noms_dos[i]}", affiche=f(aff) + " $", lignes=f(somme) + " $",
                               ecart=(f(aff - somme, signe=True) + " $") if aff and abs(aff - somme) > 1 else "0",
                               part=p(100 * aff / 4_000_000)))
    budget_tot = dict(affiche=f(sum(totaux)) + " $", lignes=f(sum(sum(m or 0 for _, m, _ in v) for v in lignes_act.values())) + " $")
    budget_t5 = [("Contact tracing et surveillance", 900_000), ("PCI/WASH, comblement des gaps et préparation", 700_000),
                 ("Communication sur les risques et engagement communautaire", 500_000),
                 ("Préparation des provinces et ZS en anneau", 500_000), ("Déplacés et milieux humanitaires", 200_000)]
    budget_t5 = [dict(pilier=f"{i}. {n}", montant=f(m) + " $", part=p(100 * m / 2_800_000)) for i, (n, m) in enumerate(budget_t5, 1)]

    # Modèle de données
    rempli = {"ref_zones": "s02 (SitReps + config)", "ref_partenaires": "s02 (config)", "ref_indicateurs": "s02 (matrice)",
              "ref_semaines": "s02 (calendrier ISO)", "sitrep_meta": "s01", "epi_province": "s01", "epi_zs": "s01",
              "alertes_province": "s01", "contacts_province": "s01", "labo_province": "s01", "pec_province": "s01",
              "poe_province": "s01", "sitrep_texte": "s01", "historique": "s02 (plan DOS, réunions) ; s02b (Weekly Updates)",
              "autres_epidemies": "s02b (codage des Weekly Updates)",
              "saisie_hebdo": "IP (vendredi) ; s02 (classeur) ; s02b (Weekly Updates)", "sbc_activites": "Export ODK",
              "sbc_synthese_as": "Export ODK", "reunions": "Représentant CGPP", "points_action": "Représentant CGPP",
              "sitrep_notes": "Équipe SI", "qc_suivi": "s03 puis équipe SI"}
    modele = [dict(nom=k, titre=v["titre"], cle=", ".join(v["cle"]), n=len(v["colonnes"]), rempli=rempli.get(k, ""))
              for k, v in SCHEMA.items()]
    n_lignes = {k: len(lire_table(base, k)) for k in SCHEMA}
    for m in modele:
        m["lignes"] = f(n_lignes[m["nom"]])

    # chiffres de l'extraction des SitReps (cités dans le texte)
    ez = lire_table(base, "epi_zs")
    a_vent = ez.est_a_ventiler.fillna(False).astype(bool)
    zs_num = ez[~a_vent].groupby("sitrep_no").size()
    zs_par_sitrep = " puis ".join(str(int(x)) for x in zs_num.sort_index().values)
    n_paragraphes = len(lire_table(base, "sitrep_texte"))
    n_colonnes = sum(len(v["colonnes"]) for v in SCHEMA.values())

    # série nationale (s04) et Weekly Updates
    q, ivs = res.get("quinzaine", {}), res.get("intervalles", [])
    d_ref = pd.to_datetime(q["date_ref"]) if q else None
    recents = [i for i in ivs if pd.to_datetime(i["debut"]) >= d_ref] if q else []
    long_iv = max(ivs, key=lambda i: i["jours"]) if ivs else None
    rythmes = [i["jour_cas_National"] for i in recents] or [nat["moyenne_jour"]]
    nk = [i["jour_cas_Nord-Kivu"] for i in recents] or [0]
    jm = lambda d: date_fr(d).rsplit(" ", 1)[0]  # noqa: E731
    Q = dict(cas=f(q.get("cas")), deces=f(q.get("deces")), jours=q.get("jours"), debut=jm(d_ref) if q else "",
             fin=jm(nat["date"]), par_jour=f(q["cas"] / q["jours"]) if q else "", min=f(min(rythmes)), max=f(max(rythmes)),
             long=f(long_iv["jour_cas_National"]) if long_iv else "",
             part_nk=p(100 * q["provinces"].get("Nord-Kivu", 0) / q["cas"], 0) if q else "",
             part_nk_long=p(100 * long_iv["cas_Nord-Kivu"] / long_iv["cas_National"], 0) if long_iv else "",
             nk_min=f(nk[0]), nk_max=f(nk[1] if len(nk) > 1 else nk[0]))
    ops = res.get("operations", {}).get("meta", {})
    fin_ops = pd.to_datetime(ops["fin"]) if ops else None
    wu = dict(ecarts=liste_fr(ecarts_publies(cfg, fin_ops)) if ops else "",
              retard=(pd.to_datetime(nat["date"]) - fin_ops).days if ops else "")

    fig = cfg["chemins"]["sorties_figures"]
    return dict(
        Q=Q, wu=wu, n_tables=len(SCHEMA),
        zs_par_sitrep=zs_par_sitrep, n_lignes_zs=f(len(ez)), n_paragraphes=f(n_paragraphes), n_colonnes=f(n_colonnes),
        date_doc=date_fr(DATE_DOC), N=N, C=C, V=V, prov_rows=prov_rows, dc_rows=dc_rows,
        post146=p(dct.loc[nat["sitrep_prec"]].part_post_mortem_pct), post149=p(dct.loc[nat["sitrep"]].part_post_mortem_pct),
        comm146=p(dct.loc[nat["sitrep_prec"]].part_comm_parmi_deces_pct), comm149=p(dct.loc[nat["sitrep"]].part_comm_parmi_deces_pct),
        part_nk=p(parts["Nord-Kivu"].part_cas_pct), part_nk_085=p(parts["Nord-Kivu"].part_085_pct),
        part_it=p(parts["Ituri"].part_cas_pct), part_it_085=p(parts["Ituri"].part_085_pct),
        ct_rows=ct_rows, tot_contacts_146=tot_contacts_146, al_rows=al_rows, val146=val146, lab_rows=lab_rows,
        zs_top=zs_top, zs_cgpp=zs_cgpp, semaines=semaines, grille=grille, completude=p(res["completude_classeur"]),
        cellules=f"{res['cellules_renseignees']} sur {res['cellules_attendues']}", realis=realis,
        qc=qc, familles=fam, anomalies=anomalies, regles=regles,
        crit=[a for a in anomalies if a["gravite"] == "Critique"],
        mat_rows=mat_rows, mat_tot=mat_tot, nouveaux_p1=nouveaux_p1, nouveaux_p23=nouveaux_p23, a_corriger=a_corriger,
        annexe_ind=annexe_ind, n_t5_only=len(t5_only), cibles=cibles, budget_dos=budget_dos, budget_tot=budget_tot,
        budget_t5=budget_t5, modele=modele,
        svg=dict(parts=svg_en_ligne(fig / "f1_parts_provinces.svg"), rythme=svg_en_ligne(fig / "f6_rythme.svg"),
                 croissance=svg_en_ligne(fig / "f2_croissance_provinces.svg"),
                 hausse=svg_en_ligne(fig / "f3_hausse_zs.svg"), contacts=svg_en_ligne(fig / "f4_suivi_contacts.svg"),
                 zs_cgpp=svg_en_ligne(fig / "f5_zs_cgpp.svg"), flux=svg_en_ligne(ICI / "flux_cible.svg"),
                 chaine=svg_en_ligne(ICI / "chaine_sitrep.svg")),
        date_donnees=pd.to_datetime(nat["date"]).strftime("%d/%m/%Y"),
        date_donnees_prec=pd.to_datetime(nat["date_prec"]).strftime("%d/%m/%Y"),
        date_qc=pd.to_datetime(re.search(r"(\d{4}-\d{2}-\d{2})", qcf.name).group(1)).strftime("%d/%m/%Y"),
    )


POLICE_BANDEAU = ICI / "polices" / "Inter-Regular.ttf"


def imprimer(html_path: Path, pdf_path: Path, marges: bool = True):
    """Impression Chromium. L'en-tête et le pied de page sont ajoutés ensuite par estampiller() :
    les modèles d'en-tête de Chromium n'ont pas accès aux polices du document."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        nav = pw.chromium.launch()
        page = nav.new_page()
        page.goto(html_path.as_uri(), wait_until="networkidle")
        page.evaluate("document.fonts.ready")
        m = dict(top="18mm", bottom="16mm", left="17mm", right="17mm") if marges else dict(top="0", bottom="0", left="0", right="0")
        page.pdf(path=str(pdf_path), format="A4", print_background=True, prefer_css_page_size=True, margin=m)
        nav.close()


def estampiller(corps_pdf: Path, ctx: dict) -> PdfReader:
    """Ajoute en-tête et pied de page (Inter TrueType, numéro de page) à chaque page du corps."""
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont as PoliceTT
    from reportlab.pdfgen import canvas
    pdfmetrics.registerFont(PoliceTT("InterBandeau", str(POLICE_BANDEAU)))
    lecteur = PdfReader(corps_pdf)
    n = len(lecteur.pages)
    gauche_haut = "CGPP RDC  |  Information stratégique, riposte MVE-17"
    droite_haut = f"Document de travail, {ctx['date_doc']}"
    gauche_bas = (f"Données au {ctx['date_donnees']} (SitRep national n° {ctx['N']['sitrep']}) ; "
                  f"base CGPP et contrôle qualité du {ctx['date_qc']}")
    tampon = io.BytesIO()
    c = canvas.Canvas(tampon, initialFontName="InterBandeau", initialFontSize=6.8)
    for i, pg in enumerate(lecteur.pages, start=1):
        larg, haut = float(pg.mediabox.width), float(pg.mediabox.height)
        c.setPageSize((larg, haut))
        c.setFont("InterBandeau", 6.8)
        c.setFillColorRGB(0x89 / 255, 0x87 / 255, 0x81 / 255)
        c.drawString(17 * mm, haut - 10.5 * mm, gauche_haut)
        c.drawRightString(larg - 17 * mm, haut - 10.5 * mm, droite_haut)
        c.drawString(17 * mm, 8.5 * mm, gauche_bas)
        c.drawRightString(larg - 17 * mm, 8.5 * mm, f"Page {i} / {n}")
        c.setStrokeColorRGB(0xd9 / 255, 0xd8 / 255, 0xd2 / 255)
        c.setLineWidth(0.4)
        c.line(17 * mm, haut - 12.2 * mm, larg - 17 * mm, haut - 12.2 * mm)
        c.showPage()
    c.save()
    calque = PdfReader(io.BytesIO(tampon.getvalue()))
    for pg, cal in zip(lecteur.pages, calque.pages):
        pg.merge_page(cal)
    return lecteur


def compact(t: str) -> str:
    """Texte sans espaces (y compris insécables) pour comparer titres et pages extraites du PDF."""
    return re.sub(r"[\s   ]+", "", t)


def pages_des_ancres(pdf_path: Path) -> dict:
    """Page (numérotation du corps) de chaque repère invisible « ⟦Hk⟧ » placé dans les titres."""
    pages = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, pg in enumerate(pdf.pages, start=1):
            for k in re.findall(r"⟦H(\d+)⟧", compact(pg.extract_text() or "")):
                pages.setdefault(int(k), i)
    return pages


def titres_absents(pdf_path: Path, titres: list, pages: dict) -> list:
    """Titres introuvables sur la page annoncée par le sommaire (contrôle de la passe finale)."""
    with pdfplumber.open(pdf_path) as pdf:
        textes = [compact(pg.extract_text() or "") for pg in pdf.pages]
    manquants = []
    for niv, k, t in titres:
        pg = pages.get(k)
        if not pg or compact(t)[:30] not in textes[pg - 1]:
            manquants.append(t)
    return manquants


def liens_sommaire(pdf_path: Path, titres: list, pages: dict, decalage: int) -> list:
    """Zones cliquables du sommaire : (index de la page du sommaire, rectangle PDF, index de la page cible)."""
    zones = []
    attendus = [(compact(t), pages[k]) for niv, k, t in titres if k in pages]
    j = 0
    with pdfplumber.open(pdf_path) as pdf:
        for ip, pg in enumerate(pdf.pages):
            if j >= len(attendus):
                break
            mots = pg.extract_words(keep_blank_chars=False, use_text_flow=False)
            lignes = {}
            for m in mots:
                lignes.setdefault(round(m["top"]), []).append(m)
            for top in sorted(lignes):
                ms = sorted(lignes[top], key=lambda m: m["x0"])
                texte = compact("".join(m["text"] for m in ms))
                if j < len(attendus) and texte.startswith(attendus[j][0][:24]) and texte.endswith(str(attendus[j][1])):
                    x0, x1 = min(m["x0"] for m in ms), max(m["x1"] for m in ms)
                    haut, bas = min(m["top"] for m in ms), max(m["bottom"] for m in ms)
                    h = float(pg.height)
                    zones.append((ip, (x0 - 1, h - bas - 1.5, x1 + 1, h - haut + 1.5), attendus[j][1] - 1 + decalage))
                    j += 1
            if zones and j < len(attendus) and not any(z[0] == ip for z in zones):
                break
    return zones


def main():
    cfg = charger_config()
    out = RACINE / "sorties" / "rapport"
    out.mkdir(parents=True, exist_ok=True)
    # aperçu de toutes les pages du SitRep CGPP généré (grille 2 x 2 dans la partie 8)
    srep = sorted(cfg["chemins"]["sorties_sitrep"].glob("SitRep_CGPP_*.pdf"))[-1]
    for ancien in ICI.glob("_sitrep*"):
        ancien.unlink()
    subprocess.run(["pdftoppm", "-jpeg", "-jpegopt", "quality=88", "-r", "150", str(srep), str(ICI / "_sitrep")],
                   check=True)
    apercus = sorted(ICI.glob("_sitrep-*.jpg"), key=lambda x: int(re.search(r"-(\d+)\.jpg$", x.name).group(1)))
    ctx = contexte()
    ctx["apercus"] = [x.as_uri() for x in apercus[:4]]
    n_pages = len(PdfReader(srep).pages)
    ctx["pages_sitrep"] = {2: "deux", 3: "trois", 4: "quatre", 5: "cinq", 6: "six"}.get(n_pages, str(n_pages))
    ctx["sitrep_fichier"] = srep.name
    env = Environment(loader=FileSystemLoader(str(ICI)), autoescape=True)
    corps = env.get_template("rapport.html.j2")
    garde = env.get_template("garde.html.j2")
    (out / "_garde.html").write_text(garde.render(**ctx), encoding="utf-8")
    imprimer(out / "_garde.html", out / "_garde.pdf", marges=False)

    # passe 0 : liste des titres ; passe 1 : sommaire sans numéros ; passe 2 : sommaire numéroté
    ctx["toc"] = []
    html0 = corps.render(**ctx)
    titres = [(int(niv), int(k), re.sub(r"<[^>]+>", "", t).strip())
              for niv, k, t in re.findall(r'<h([12])\b[^>]*?\bdata-toc="(\d+)"[^>]*>(.*?)<span class="ancre"', html0, flags=re.S)]
    titres = [(niv, k, t.replace("&#39;", "'").replace("&amp;", "&")) for niv, k, t in titres]
    ctx["toc"] = [dict(niv=niv, titre=t, page="00") for niv, k, t in titres]
    (out / "_corps.html").write_text(corps.render(**ctx), encoding="utf-8")
    imprimer(out / "_corps.html", out / "_corps.pdf")
    pages = pages_des_ancres(out / "_corps.pdf")
    ctx["toc"] = [dict(niv=niv, titre=t, page=pages.get(k, "")) for niv, k, t in titres]
    html2 = corps.render(**ctx)
    (out / "_corps.html").write_text(html2, encoding="utf-8")
    imprimer(out / "_corps.html", out / "_corps.pdf")
    verif = pages_des_ancres(out / "_corps.pdf")
    ecarts = [t for niv, k, t in titres if verif.get(k) != pages.get(k)]
    # passe 3 : version finale sans repères (hors flux : la mise en page ne bouge pas)
    (out / "_corps.html").write_text(re.sub(r'<span class="ancre">⟦H\d+⟧</span>', "", html2), encoding="utf-8")
    imprimer(out / "_corps.html", out / "_corps.pdf")
    absents = titres_absents(out / "_corps.pdf", titres, pages)
    n_corps = len(PdfReader(out / "_corps.pdf").pages)

    # assemblage : garde + corps, numérotation affichée par le lecteur, signets, liens du sommaire
    corps_final = estampiller(out / "_corps.pdf", ctx)
    w = PdfWriter()
    for pg in PdfReader(out / "_garde.pdf").pages:
        w.add_page(pg)
    for pg in corps_final.pages:
        w.add_page(pg)
    for pg in w.pages:          # le calque fusionné laisse des flux de contenu non compressés
        pg.compress_content_streams()
    w.set_page_label(0, 0, prefix="Couverture")
    w.set_page_label(1, n_corps, style="/D", start=1)
    w.add_outline_item("Sommaire", 1)
    parent = None
    for niv, k, t in titres:
        if k not in pages:
            continue
        if niv == 1:
            parent = w.add_outline_item(t, pages[k])
        else:
            w.add_outline_item(t, pages[k], parent=parent)
    w.page_mode = "/UseOutlines"
    zones = liens_sommaire(out / "_corps.pdf", titres, pages, decalage=1)
    for ip, rect, cible in zones:
        w.add_annotation(page_number=ip + 1, annotation=Link(rect=rect, target_page_index=cible, border=[0, 0, 0]))
    w.add_metadata({"/Title": "Information stratégique CGPP, riposte MVE-17 : données, qualité, indicateurs et SitRep automatisé",
                    "/Author": "Etienne Ntumba Kabongo, Senior Data Analyst (SI), CGPP RDC",
                    "/Subject": "Analyse des documents de la riposte, contrôle qualité, cadre d'indicateurs, automatisation du SitRep",
                    "/Keywords": "CGPP, MVE-17, Bundibugyo, SitRep, contrôle qualité, indicateurs, RDC",
                    "/Creator": "outils/rapport/generer_rapport.py (Jinja2, Chromium, pypdf)"})
    w.compress_identical_objects()
    final = out / f"Rapport_SI_CGPP_MVE17_{DATE_DOC:%Y-%m-%d}.pdf"
    with open(final, "wb") as fh:
        w.write(fh)
    for tmp in ("_garde.pdf", "_corps.pdf"):
        (out / tmp).unlink(missing_ok=True)
    print("PDF :", final, "| pages :", len(PdfReader(final).pages), "| titres :", len(titres),
          "| sans page :", [t for niv, k, t in titres if k not in pages], "| décalés entre passes :", ecarts,
          "| absents de leur page :", absents, "| liens du sommaire :", f"{len(zones)}/{len(titres)}")


if __name__ == "__main__":
    main()

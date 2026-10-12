"""
pipeline.py - Chaîne SI CGPP complète : canevas, extraction, contrôle qualité, analyses, SitRep CGPP.

Usage :
  python pipeline.py                        toutes les étapes
  python pipeline.py --etapes qc,sitrep     étapes choisies (canevas, sitreps, indicateurs, weekly, qc, analyses, sitrep)
  python pipeline.py --reinitialiser        reconstruit la base à partir du canevas (efface les saisies !)
  python pipeline.py --date 2026-10-10      SitRep CGPP pour une date de données précise

Chaque étape lit et écrit la base Excel data/base/base_CGPP_MVE.xlsx ; un journal est écrit dans
sorties/journal_pipeline.log.
"""
from __future__ import annotations

import argparse
import importlib
import logging
import sys
import time
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE / "scripts"))

from commun import charger_config, journal  # noqa: E402

ETAPES = [
    ("canevas", "s00_construire_canevas", "Canevas Excel de la base"),
    ("sitreps", "s01_extraire_sitreps", "Extraction des SitReps nationaux"),
    ("indicateurs", "s02_extraire_indicateurs", "Référentiels et saisies des IP"),
    ("weekly", "s02b_extraire_weekly_updates", "Weekly Updates CGPP (série nationale, données opérationnelles)"),
    ("qc", "s03_controle_qualite", "Contrôle qualité"),
    ("analyses", "s04_analyses", "Analyses et graphiques"),
    ("sitrep", "s05_generer_sitrep", "SitRep CGPP (Word et PDF)"),
]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--etapes", default="", help="liste d'étapes séparées par des virgules")
    ap.add_argument("--reinitialiser", action="store_true", help="repartir d'une base vide (canevas)")
    ap.add_argument("--date", default=None, help="date des données du SitRep CGPP (AAAA-MM-JJ)")
    args = ap.parse_args()

    cfg = charger_config()
    log = journal("pipeline")
    fh = logging.FileHandler(RACINE / "sorties" / "journal_pipeline.log", encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s | %(name)s | %(levelname)s | %(message)s"))
    for nom in ("pipeline", "sitreps", "indicateurs", "weekly", "qc", "analyses", "sitrep", "canevas"):
        logging.getLogger(nom).addHandler(fh)

    choisies = [e.strip() for e in args.etapes.split(",") if e.strip()] or [e[0] for e in ETAPES]
    if args.reinitialiser and cfg["chemins"]["base"].exists():
        cfg["chemins"]["base"].unlink()
        log.info("Base supprimée : elle sera recréée à partir du canevas.")
    if "canevas" not in choisies and not cfg["chemins"]["canevas_base"].exists():
        choisies.insert(0, "canevas")

    for code, module, libelle in ETAPES:
        if code not in choisies:
            continue
        t0 = time.time()
        log.info("=== %s", libelle)
        try:
            m = importlib.import_module(module)
            if code == "sitrep":
                m.main(cfg, args.date)
            else:
                m.main(cfg)
        except PermissionError as e:
            log.error("Fichier verrouillé : %s. Le fermer dans Excel ou Word, puis relancer.", e.filename or e)
            sys.exit(1)
        except FileNotFoundError as e:
            log.error("Fichier ou programme introuvable : %s", e)
            sys.exit(1)
        except Exception:
            log.exception("Échec de l'étape « %s » : voir le détail ci-dessous.", libelle)
            sys.exit(1)
        log.info("    terminé en %.1f s", time.time() - t0)
    log.info("Chaîne terminée.")


if __name__ == "__main__":
    main()

# Chaîne SI CGPP : riposte MVE-17 (virus Bundibugyo)

Boîte à outils de l'équipe Information stratégique (SI) de CGPP RDC. Elle transforme les fichiers de la riposte
(SitReps nationaux, Weekly Updates CGPP, classeur hebdomadaire des IP, plans de travail, comptes rendus, fiche SBC)
en une base Excel
normalisée, contrôle la qualité des données, calcule les analyses et produit le SitRep CGPP en Word et en PDF.

Auteur : Etienne Ntumba Kabongo, Senior Data Analyst (SI), CGPP RDC. Version 1.0 du 12 octobre 2026.

## 1. Installation (une fois)

1. Installer Python 3.11 ou plus (python.org ; sous Windows, cocher « Add python.exe to PATH »).
2. Installer LibreOffice (libreoffice.org). Il sert à lire les SitReps `.doc` et à produire les PDF.
3. Ouvrir un terminal dans le dossier `cgpp_si`, puis :

```
python -m venv .venv
.venv\Scripts\activate            (Windows)
source .venv/bin/activate         (Linux, macOS)
pip install -r requirements.txt
```

LibreOffice est cherché automatiquement (PATH, puis `C:/Program Files/LibreOffice/program/soffice.exe`).
S'il n'est pas trouvé, indiquer son chemin dans `config.yaml`, clé `libreoffice`.

## 2. Utilisation courante

1. Déposer les fichiers dans `data/entree/` : SitReps nationaux dans `sitreps/`, Weekly Updates CGPP dans
   `weekly_updates/`, comptes rendus dans `reunions/`.
   Le nom des fichiers est libre : le numéro et la date sont lus dans l'en-tête du SitRep
   (« SitRep N°149 /MVE-BDBV/10/10/2026 »). Les SitReps déjà présents dans la base sont mis à jour, pas dupliqués.
2. Lancer la chaîne :

```
python pipeline.py                         toute la chaîne
python pipeline.py --etapes qc,sitrep      seulement le contrôle qualité et le SitRep
python pipeline.py --date 2026-10-10       SitRep CGPP pour une date de données précise
```

   Sous Windows, un double-clic sur `lancer.bat` fait la même chose (journal dans `sorties/journal_lancer.log`).
3. Lire les résultats dans `sorties/` :

| Dossier | Contenu |
|---|---|
| `sorties/sitrep/` | SitRep CGPP en `.docx` (à relire, modifiable) et `.pdf` (à diffuser) |
| `sorties/qc/` | rapport de contrôle qualité : synthèse, anomalies, tous les contrôles, pourcentages des SitReps |
| `sorties/analyses/` | tableaux d'analyse (`analyses.xlsx`) et chiffres clés (`resultats.json`) |
| `sorties/figures/` | graphiques en SVG et PNG |
| `sorties/journal_pipeline.log` | journal de chaque exécution |

Étapes de la chaîne (chacune lit et écrit `data/base/base_CGPP_MVE.xlsx`) :

| Code | Script | Rôle |
|---|---|---|
| `canevas` | `s00_construire_canevas.py` | canevas Excel vide : listes déroulantes, dictionnaire, tableau de suivi par formules |
| `sitreps` | `s01_extraire_sitreps.py` | extraction des SitReps nationaux (tableaux, encadrés, texte, pourcentages) |
| `indicateurs` | `s02_extraire_indicateurs.py` | référentiels, classeur des IP en format long, matrice des indicateurs, historique |
| `weekly` | `s02b_extraire_weekly_updates.py` | Weekly Updates : tableau national (série des SitReps), tableau opérationnel des IP, activités codées, autres épidémies |
| `qc` | `s03_controle_qualite.py` | contrôles automatiques (261 au 12/10), rapport QC, suivi des anomalies |
| `analyses` | `s04_analyses.py` | tableaux, chiffres clés et graphiques |
| `sitrep` | `s05_generer_sitrep.py` | SitRep CGPP à partir du canevas Word `modeles/canevas_sitrep_CGPP.docx` |

## 3. Saisies manuelles dans la base Excel

La base `data/base/base_CGPP_MVE.xlsx` compte 22 tables (une feuille chacune), décrites colonne par colonne dans
la feuille `dictionnaire` du canevas `modeles/canevas_base_CGPP_MVE.xlsx`. Quatre feuilles se remplissent à la main ;
elles sont conservées d'une exécution à l'autre :

- `saisie_hebdo` (points focaux IP, chaque vendredi) : une ligne par semaine, partenaire, ZS, indicateur et
  désagrégation ; saisir le numérateur et le dénominateur, jamais un pourcentage. La colonne `valeur` se calcule seule.
- `points_action` (représentant CGPP après chaque réunion) : action, responsable, échéance, statut.
- `sitrep_notes` (équipe SI) : faits marquants, activités, sécurité, défis et actions repris dans le SitRep.
- `qc_suivi` (équipe SI) : responsable, statut et date de résolution de chaque anomalie détectée.

Les Weekly Updates sont lus automatiquement pour leurs tableaux (situation nationale, données opérationnelles des IP).
Leurs passages narratifs chiffrés se codent dans deux fichiers relus par l'équipe SI : `data/ref/weekly_updates_activites.csv`
(une ligne par activité : indicateur, valeur, période, extrait du texte source) et `data/ref/autres_epidemies.csv`
(choléra, mpox, rougeole). Les constats qualité non calculables se notent dans `data/ref/observations_manuelles.csv`.

Règles : cellule vide = non rapporté ; zéro = rapporté nul ; pas de cellules fusionnées ni de lignes de total ;
choisir les provinces, ZS, partenaires et indicateurs dans les listes. Fermer la base dans Excel avant de lancer
la chaîne (un fichier ouvert est verrouillé sous Windows).

L'option `python pipeline.py --reinitialiser` reconstruit la base à partir du canevas et efface les saisies : ne
l'utiliser qu'après une sauvegarde.

## 4. Planification

Sous Windows, pour produire le SitRep chaque samedi à 18 h :

```
schtasks /create /sc weekly /d SAT /st 18:00 /tn SitRepCGPP /tr "C:\cgpp_si\lancer.bat"
```

## 5. Faire évoluer la chaîne

- Ajouter un indicateur : une ligne dans `scripts/referentiel_indicateurs.py` (identifiant texte stable, type,
  numérateur, dénominateur, source, cibles) ; il apparaît dans `ref_indicateurs` et dans le tableau de suivi.
- Ajouter une règle de contrôle : une fonction dans `scripts/s03_controle_qualite.py` qui appelle `res(...)` avec
  l'identifiant de la règle, la gravité et la recommandation.
- Changer la mise en page du SitRep : ouvrir `modeles/canevas_sitrep_CGPP.docx` dans Word et modifier le texte ou
  les styles en gardant chaque balise `{{ ... }}` ou `{% ... %}` d'un seul tenant (ne pas la couper ni la mettre
  en forme à moitié). Le canevas peut aussi être reconstruit avec `node modeles/outils/construire_canevas_sitrep.js
  modeles/canevas_sitrep_CGPP.docx` (bibliothèque `docx` de Node.js).
- Ajouter une ZS ou un partenaire : bloc `zones_cgpp` de `config.yaml`.
- Seuils (suivi des contacts, cibles, tolérance des pourcentages, définition d'un foyer) : bloc `seuils` de `config.yaml`.
- Nouvelle maquette du SitRep national : les tableaux sont repérés par leurs en-têtes ; adapter
  `trouver_tableaux` dans `scripts/s01_extraire_sitreps.py`.

## 6. Rapport PDF d'analyse (facultatif)

Le rapport « Riposte à la MVE-17 : données, qualité, indicateurs et SitRep automatisé » se régénère à partir des
sorties de la chaîne :

```
pip install -r outils/rapport/requirements-rapport.txt
python -m playwright install chromium
python outils/rapport/generer_rapport.py
```

Le PDF est écrit dans `sorties/rapport/`. Les polices Inter (licence SIL OFL, `outils/rapport/polices/`) sont fournies.

## 7. Données personnelles

La base d'indicateurs ne contient ni noms ni numéros de téléphone. Les listes de présence des réunions (photos des
comptes rendus) et les listes de contacts contiennent des données personnelles : ne pas les copier dans la base ni
les joindre aux documents diffusés.

## 8. Dépannage

| Message | Cause et solution |
|---|---|
| `Fichier verrouillé` | la base ou le SitRep est ouvert dans Excel ou Word : le fermer et relancer |
| `LibreOffice introuvable` | installer LibreOffice ou renseigner `libreoffice` dans `config.yaml` |
| `Aucun SitRep national trouvé` | vérifier le dossier `data/entree/sitreps/` (fichiers `.doc` ou `.docx`) |
| `SitRep ignoré (...) : Numéro de SitRep introuvable` | l'en-tête du SitRep n'a pas la forme « SitRep N°149 /MVE-BDBV/10/10/2026 » |
| Valeurs « ND » dans le SitRep | la province n'a pas rapporté ce jour-là : c'est voulu, le périmètre est signalé |
| Section vide dans le SitRep | la feuille correspondante (`sitrep_notes`, `points_action`) n'a pas de ligne pour la date |
| `Weekly Update ignoré (...)` | le titre ne porte pas de date en anglais (« September 30th, 2026 ») ou le fichier n'est pas un .docx |
| « Données opérationnelles : aucune » | aucun tableau consolidé (« Indicator / ADRA (Ituri) / ... ») dans les Weekly Updates déposés |

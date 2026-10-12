// construire_canevas_sitrep.js
// Construit le canevas Word du SitRep CGPP (balises docxtpl/Jinja entre {{ }} et {% %}).
// Le canevas se modifie ensuite directement dans Word : garder chaque balise d'un seul tenant.
// Usage : node modeles/outils/construire_canevas_sitrep.js modeles/canevas_sitrep_CGPP.docx
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  HeadingLevel, BorderStyle, Header, Footer, PageNumber, LevelFormat, TabStopType, VerticalAlign, Tab,
} = require("docx");

const POLICE = "Arial";
const BLEU = "1F3864", BLEU_PALE = "DCE6F2", GRIS_TUILE = "F3F5F8", JAUNE = "FFF2CC";
const ENCRE = "0B0B0B", ENCRE2 = "52514E", BORD = "BFBFBF";
const LARGEUR = 9638; // A4, marges de 2 cm

const t = (text, o = {}) => new TextRun({ text, font: POLICE, size: o.size || 19, bold: o.bold, italics: o.italics,
  color: o.color || ENCRE });
const p = (runs, o = {}) => new Paragraph({ children: Array.isArray(runs) ? runs : [runs],
  alignment: o.align, spacing: { before: o.before || 0, after: o.after === undefined ? 60 : o.after, line: o.line },
  numbering: o.puce ? { reference: "puces", level: 0 } : undefined, keepNext: o.keepNext });
const balise = (texte) => p(t(texte, { size: 4, color: "FFFFFF" }), { after: 0 }); // paragraphe de contrôle docxtpl
const bordures = (c = BORD) => ({ top: { style: BorderStyle.SINGLE, size: 4, color: c },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: c }, left: { style: BorderStyle.SINGLE, size: 4, color: c },
  right: { style: BorderStyle.SINGLE, size: 4, color: c } });
const sansBord = { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } };

function cellule(contenu, largeur, o = {}) {
  const enfants = (Array.isArray(contenu) ? contenu : [contenu]).map((c) =>
    c instanceof Paragraph ? c : p(t(String(c), { size: o.size || 17, bold: o.bold, color: o.color }),
      { align: o.align, after: 0 }));
  return new TableCell({ children: enfants, width: { size: largeur, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: { top: o.mt || 36, bottom: o.mb || 36, left: 90, right: 90 }, borders: o.borders || bordures(),
    verticalAlign: o.valign || VerticalAlign.CENTER, columnSpan: o.span });
}

function tableauBoucle(entetes, largeurs, champs, boucle, totaux, gauche = [0]) {
  // entetes : libellés ; champs : balises de la ligne répétée ; boucle : "p in provinces"
  const n = largeurs.length;
  const lignes = [
    new TableRow({ tableHeader: true, children: entetes.map((e, i) =>
      cellule(e, largeurs[i], { bold: true, color: "FFFFFF", fill: BLEU, size: 16, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })) }),
    new TableRow({ children: [cellule(`{%tr for ${boucle} %}`, LARGEUR, { span: n, size: 4 })] }),
    new TableRow({ children: champs.map((c, i) => cellule(c, largeurs[i], { align: gauche.includes(i) ? AlignmentType.LEFT : AlignmentType.CENTER })) }),
    new TableRow({ children: [cellule("{%tr endfor %}", LARGEUR, { span: n, size: 4 })] }),
  ];
  if (totaux) lignes.push(new TableRow({ children: totaux.map((c, i) =>
    cellule(c, largeurs[i], { bold: true, fill: BLEU_PALE, align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })) }));
  return new Table({ width: { size: LARGEUR, type: WidthType.DXA }, columnWidths: largeurs, rows: lignes });
}

function tuile(i) {
  return cellule([
    p(t(`{{ t${i}.label }}`, { size: 15, color: ENCRE2 }), { after: 20 }),
    p(t(`{{ t${i}.valeur }}`, { size: 30, bold: true }), { after: 10 }),
    p(t(`{{ t${i}.sous }}`, { size: 15, color: ENCRE2 }), { after: 0 }),
  ], LARGEUR / 4, { fill: GRIS_TUILE, borders: bordures("FFFFFF"), mt: 90, mb: 90, valign: VerticalAlign.TOP });
}

const titre = (texte) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(texte, { size: 23, bold: true, color: BLEU })],
  spacing: { before: 200, after: 80 }, keepNext: true });
const sousTitre = (texte) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t(texte, { size: 19, bold: true, color: BLEU })],
  spacing: { before: 120, after: 40 }, keepNext: true });
const listeBoucle = (variable, liste) => [balise(`{%p for ${variable} in ${liste} %}`),
  p(t(`{{ ${variable} }}`, { size: 18 }), { puce: true, after: 30 }), balise("{%p endfor %}")];

const bandeau = new Table({ width: { size: LARGEUR, type: WidthType.DXA }, columnWidths: [LARGEUR], rows: [new TableRow({ children: [
  cellule([
    p(t("SITREP CGPP N° {{ numero }}", { size: 34, bold: true, color: "FFFFFF" }), { after: 30 }),
    p(t("Riposte à la maladie à virus Ebola (virus Bundibugyo), 17e épidémie, République démocratique du Congo",
      { size: 18, color: "FFFFFF" }), { after: 30 }),
    p(t("Données nationales au {{ date_donnees }} (SitRep n° {{ sitrep_national }})  |  Données CGPP du {{ periode_ops }}  |  Publié le {{ date_publication }}",
      { size: 15, color: "C9D8F0" }), { after: 0 }),
  ], LARGEUR, { fill: BLEU, borders: sansBord, mt: 150, mb: 150 })] })] });

const completude = new Table({ width: { size: LARGEUR, type: WidthType.DXA }, columnWidths: [LARGEUR], rows: [new TableRow({ children: [
  cellule([p([t("Sources et complétude : ", { size: 16, bold: true }), t("{{ completude_texte }}", { size: 16 })], { after: 0 })],
    LARGEUR, { fill: JAUNE, borders: bordures("F2C200"), mt: 60, mb: 60 })] })] });

const tuiles = new Table({ width: { size: LARGEUR, type: WidthType.DXA }, columnWidths: [2410, 2410, 2409, 2409], rows: [
  new TableRow({ children: [1, 2, 3, 4].map(tuile) }), new TableRow({ children: [5, 6, 7, 8].map(tuile) })] });

const note = (texte, o = {}) => p(t(texte, { size: 15, italics: true, color: ENCRE2 }), { before: 40, after: o.after || 80 });

const provinces = tableauBoucle(
  ["Province", "Nouveaux (24 h)", "Cas cumulés", "Hausse {{ j14 }} j", "Hausse {{ jours }} j", "Décès", "Létalité", "ZS touchées", "Contacts vus (24 h)"],
  [1300, 1000, 1000, 1000, 1000, 900, 900, 1000, 1538],
  ["{{ p.nom }}", "{{ p.nouveaux }}", "{{ p.cas }}", "{{ p.h14 }}", "{{ p.hausse }}", "{{ p.deces }}", "{{ p.letalite }}", "{{ p.zs }}",
    "{% cellbg p.couleur_contacts %}{{ p.contacts }}"],
  "p in provinces",
  ["Total", "{{ tot.nouveaux }}", "{{ tot.cas }}", "{{ tot.h14 }}", "{{ tot.hausse }}", "{{ tot.deces }}", "{{ tot.letalite }}", "{{ tot.zs }}", "{{ tot.contacts }}"]);

const zsCgpp = tableauBoucle(
  ["Zone de santé", "Province", "IP", "Cas cumulés", "Hausse ({{ jours }} j)", "Nouveaux (24 h)", "Décès", "Létalité"],
  [1800, 1300, 800, 1200, 1100, 1200, 1100, 1138],
  ["{{ z.nom }}", "{{ z.province }}", "{{ z.ip }}", "{{ z.cas }}", "{% cellbg z.couleur_hausse %}{{ z.hausse }}", "{{ z.nouveaux }}", "{{ z.deces }}", "{{ z.letalite }}"],
  "z in zs_cgpp",
  ["Total ZS CGPP", "", "", "{{ tot_cgpp.cas }}", "{{ tot_cgpp.hausse }}", "{{ tot_cgpp.nouveaux }}", "{{ tot_cgpp.deces }}", "{{ tot_cgpp.letalite }}"]);

const indicateurs = tableauBoucle(
  ["ID", "Indicateur", "Dernière valeur", "Période et source", "Cible DOS", "Lecture"],
  [900, 3000, 1500, 1638, 1000, 1600],
  ["{{ i.id }}", "{{ i.libelle }}", "{{ i.valeur }}", "{{ i.periode }}", "{{ i.cible }}", "{% cellbg i.couleur %}{{ i.statut }}"],
  "i in indicateurs", null, [0, 1]);

const epidemies = tableauBoucle(
  ["Maladie", "Province", "Cas (décès)", "ZS notifiantes", "ZS CGPP concernées", "Cas suspects référés par les RECO CGPP"],
  [1000, 1050, 1100, 1000, 2500, 2988],
  ["{{ e.maladie }}", "{{ e.province }}", "{{ e.cas }}", "{{ e.zs }}", "{{ e.cgpp }}", "{{ e.ref }}"],
  "e in epidemies", null, [0, 1, 4, 5]);

const actions = tableauBoucle(
  ["Réunion", "Point d'action", "Responsable", "Échéance", "Statut"],
  [1500, 4338, 1500, 1100, 1200],
  ["{{ a.reunion }}", "{{ a.action }}", "{{ a.responsable }}", "{{ a.echeance }}", "{% cellbg a.couleur %}{{ a.statut }}"],
  "a in actions", null, [0, 1]);

const doc = new Document({
  creator: "Unité SI CGPP RDC", title: "SitRep CGPP riposte MVE-17", description: "Canevas docxtpl du SitRep CGPP",
  styles: { default: { document: { run: { font: POLICE, size: 19 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: POLICE, size: 23, bold: true, color: BLEU }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: POLICE, size: 19, bold: true, color: BLEU }, paragraph: { spacing: { before: 120, after: 40 }, outlineLevel: 1 } },
    ] },
  numbering: { config: [{ reference: "puces", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
    alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 900, left: 1134, right: 1134, header: 500, footer: 450 } } },
    headers: { default: new Header({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: LARGEUR }],
      children: [t("CGPP RDC  |  SitRep CGPP riposte MVE-17", { size: 14, color: ENCRE2 }),
        new TextRun({ font: POLICE, size: 14, color: ENCRE2, children: [new Tab(), "N° {{ numero }}, données au {{ date_donnees }}"] })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: LARGEUR }],
      children: [t("Généré automatiquement le {{ date_generation }} à partir de la base CGPP (scripts SI) ; à relire avant diffusion", { size: 13, color: ENCRE2 }),
        new TextRun({ font: POLICE, size: 13, color: ENCRE2, children: [new Tab(), "Page "] }),
        new TextRun({ font: POLICE, size: 13, color: ENCRE2, children: [PageNumber.CURRENT] }),
        new TextRun({ font: POLICE, size: 13, color: ENCRE2, children: [" / "] }),
        new TextRun({ font: POLICE, size: 13, color: ENCRE2, children: [PageNumber.TOTAL_PAGES] })] })] }) },
    children: [
      bandeau, p(t(""), { after: 60 }), completude, p(t(""), { after: 60 }), tuiles,
      titre("Messages clés"), ...listeBoucle("m", "messages"),
      titre("1. Situation épidémiologique nationale"), p(t("{{ texte_epi }}"), { after: 80 }),
      p(t("{{ figure_rythme }}"), { align: AlignmentType.CENTER, after: 20 }),
      note("Figure 1. {{ legende_figure1 }}", { after: 100 }),
      provinces, note("{{ note_provinces }}", { after: 60 }),
      titre("2. Foyers et zones de santé du dispositif CGPP"), p(t("{{ texte_cgpp }}"), { after: 80 }), zsCgpp,
      note("{{ note_cgpp }}", { after: 40 }),
      balise("{%p if notes_faits and notes_faits[0] %}"), sousTitre("Faits marquants"), ...listeBoucle("n", "notes_faits"),
      balise("{%p endif %}"),
      titre("3. Riposte CGPP : données opérationnelles et activités"),
      sousTitre("3.1 Données opérationnelles consolidées des partenaires"), p(t("{{ texte_ops }}"), { after: 80 }),
      p(t("{{ marque_operations }}"), { after: 0 }), note("{{ note_ops }}", { after: 60 }),
      sousTitre("3.2 Activités de la période"), ...listeBoucle("n", "notes_activites"),
      sousTitre("3.3 Performance contre les cibles"), p(t("{{ texte_performance }}"), { after: 80 }), indicateurs,
      titre("4. Sécurité, défis et actions prioritaires"),
      sousTitre("Sécurité et accès"), ...listeBoucle("n", "notes_secu"),
      sousTitre("Défis"), ...listeBoucle("n", "notes_defis"),
      sousTitre("Actions prioritaires"), ...listeBoucle("n", "notes_actions"),
      titre("5. Autres épidémies dans les zones d'intervention"), p(t("{{ texte_epidemies }}"), { after: 80 }), epidemies,
      note("{{ note_epidemies }}", { after: 60 }),
      titre("6. Suivi des points d'action des réunions"), actions,
      titre("7. Qualité des données"), p(t("{{ texte_qc }}"), { after: 60 }), ...listeBoucle("q", "anomalies"),
      p(t("Sources : {{ sources }}", { size: 15, italics: true, color: ENCRE2 }), { before: 120, after: 40 }),
      p(t("Sigles : {{ sigles }}", { size: 15, italics: true, color: ENCRE2 }), { after: 0 }),
    ],
  }],
});

const sortie = process.argv[2] || "canevas_sitrep_CGPP.docx";
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(sortie, b); console.log("Canevas écrit :", sortie); });

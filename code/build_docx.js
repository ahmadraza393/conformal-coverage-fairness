const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  ImageRun, PageNumber, Footer, Header, PageBreak, ExternalHyperlink,
  LevelFormat, convertInchesToTwip
} = require("docx");

const FIG_DIR = "/home/claude/research/figures";

// ---------- helpers ----------
function h1(text) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 } });
}
function h2(text) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 } });
}
function p(text, opts = {}) {
  return new Paragraph({
    children: Array.isArray(text) ? text : [new TextRun(text)],
    spacing: { after: 160, line: 276 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    ...opts.pOpts,
  });
}
function italic(t) { return new TextRun({ text: t, italics: true }); }
function bold(t) { return new TextRun({ text: t, bold: true }); }
function run(t, opts = {}) { return new TextRun({ text: t, ...opts }); }

function caption(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 20 })],
    spacing: { before: 80, after: 240 },
    alignment: AlignmentType.CENTER,
  });
}

function cell(text, opts = {}) {
  const { bold: b = false, width, shading, align = AlignmentType.LEFT, size = 18 } = opts;
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold: b, size })],
    })],
  });
}

function makeTable(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((htext, i) => cell(htext, { bold: true, width: widths[i], shading: "D9E2F3", align: AlignmentType.CENTER })),
  });
  const bodyRows = rows.map((r) => new TableRow({
    children: r.map((val, i) => cell(val, { width: widths[i], align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
  }));
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...bodyRows],
  });
}

function figureImage(filename, widthPx, heightPx) {
  const data = fs.readFileSync(`${FIG_DIR}/${filename}`);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200 },
    children: [new ImageRun({ data, type: "png", transformation: { width: widthPx, height: heightPx } })],
  });
}

// ---------- title page ----------
const titlePage = [
  new Paragraph({ text: "", spacing: { after: 800 } }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
    children: [new TextRun({
      text: "Coverage Fairness in Conformal Prediction Under Class Imbalance: When Class-Conditional Calibration Helps and When It Backfires",
      bold: true, size: 30,
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [new TextRun({ text: "A Cross-Model, Cross-Regime Empirical Study on Tabular Classification", italics: true, size: 24 })],
  }),
  new Paragraph({ text: "", spacing: { after: 500 } }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
    children: [new TextRun({ text: "Ahmad Raza", bold: true, size: 24 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
    children: [new TextRun({ text: "Independent Researcher", size: 22 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 500 },
    children: [new TextRun({ text: "ahmadrz3333@gmail.com", size: 22 })],
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---------- abstract & keywords ----------
const abstractSection = [
  h1("Abstract"),
  p(
    "Split conformal prediction provides distribution-free marginal coverage guarantees, but marginal coverage can mask severe under-coverage of minority classes when the underlying label distribution is imbalanced. Class-conditional (Mondrian) conformal calibration is the standard remedy, and recent large-scale benchmarks report that it reliably restores minority-class coverage in high-stakes, severely imbalanced settings. What remains comparatively under-examined is whether this correction is uniformly beneficial, or whether its value depends jointly on the severity of the imbalance and on the number of calibration examples available per class. This study investigates that question empirically across four widely used tabular model families \u2014 logistic regression, random forest, histogram-based gradient boosting, and a radial-basis-function support vector machine \u2014 on two public datasets under three imbalance regimes: a mildly imbalanced binary dataset, a severely imbalanced binary dataset obtained by controlled minority-class subsampling, and a mildly imbalanced three-class dataset with small per-class calibration samples. Over 30 independent random data splits per condition, class-conditional calibration reduced the maximum-minimum per-class coverage gap relative to marginal calibration by a large, statistically significant margin under severe binary imbalance (mean gap reduction of 0.16\u20130.18, Holm-adjusted p < 0.01 for all four models), consistent with prior reports. Under mild binary imbalance the correction produced no significant fairness benefit while still inflating the average prediction-set size. Under mild three-class imbalance with small per-class calibration samples, class-conditional calibration significantly increased the coverage gap in three of four model families (Holm-adjusted p < 0.05), reversing its intended effect. These results indicate that class-conditional conformal calibration is not a uniformly safe default: its benefit is concentrated in regimes of severe imbalance with adequate per-class calibration data, and it can be counterproductive when classes are numerous, close to balanced, or calibration data are scarce. The study contributes a joint, cross-model characterization of this trade-off and provides a fully reproducible experimental pipeline for auditing the fairness-efficiency behavior of conformal calibration schemes before deployment.",
  ),
  h2("Keywords"),
  p("conformal prediction; class-conditional calibration; Mondrian conformal prediction; class imbalance; uncertainty quantification; coverage fairness; tabular classification"),
];

fs.writeFileSync("/tmp/sections_placeholder.txt", "ok");
console.log("helpers ready");
module.exports = { titlePage, abstractSection };

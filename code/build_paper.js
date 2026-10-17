const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType,
  ImageRun, PageNumber, Footer, Header, PageBreak,
} = require("docx");

const FIG_DIR = "/home/claude/research/figures";
const OUT = "/home/claude/research/Coverage_Fairness_Conformal_Prediction_Paper.docx";

// ---------------- helpers ----------------
function h1(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 } }); }
function h2(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 } }); }
function p(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [new TextRun(children)],
    spacing: { after: 160, line: 276 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
  });
}
function tr(t, o = {}) { return new TextRun({ text: t, ...o }); }
function caption(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 20 })],
    spacing: { before: 80, after: 240 },
    alignment: AlignmentType.CENTER,
  });
}
function cell(text, opts = {}) {
  const { bold: b = false, width, shading, align = AlignmentType.LEFT, size = 17, italics = false } = opts;
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: { top: 50, bottom: 50, left: 70, right: 70 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold: b, italics, size })],
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
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [headerRow, ...bodyRows] });
}
function figureImage(filename, widthPx, heightPx) {
  const data = fs.readFileSync(`${FIG_DIR}/${filename}`);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200 },
    children: [new ImageRun({ data, type: "png", transformation: { width: widthPx, height: heightPx } })],
  });
}

// ================= TITLE PAGE =================
const titlePage = [
  new Paragraph({ text: "", spacing: { after: 900 } }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 300 },
    children: [new TextRun({ text: "Coverage Fairness in Conformal Prediction Under Class Imbalance: When Class-Conditional Calibration Helps and When It Backfires", bold: true, size: 30 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 60 },
    children: [new TextRun({ text: "A Cross-Model, Cross-Regime Empirical Study on Tabular Classification", italics: true, size: 24 })],
  }),
  new Paragraph({ text: "", spacing: { after: 600 } }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Ahmad Raza", bold: true, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Independent Researcher", size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 900 }, children: [new TextRun({ text: "ahmadrz3333@gmail.com", size: 22 })] }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ================= ABSTRACT =================
const abstractSection = [
  h1("Abstract"),
  p("Split conformal prediction provides distribution-free marginal coverage guarantees, but marginal coverage can mask severe under-coverage of minority classes when the underlying label distribution is imbalanced. Class-conditional (Mondrian) conformal calibration is the standard remedy, and recent large-scale benchmarks report that it reliably restores minority-class coverage in high-stakes, severely imbalanced settings. What remains comparatively under-examined is whether this correction is uniformly beneficial, or whether its value depends jointly on the severity of the imbalance and on the number of calibration examples available per class."),
  p("This study investigates that question empirically across four widely used tabular model families \u2014 logistic regression, random forest, histogram-based gradient boosting, and a radial-basis-function support vector machine \u2014 on two public datasets under three imbalance regimes: a mildly imbalanced binary dataset, a severely imbalanced binary dataset obtained by controlled minority-class subsampling, and a mildly imbalanced three-class dataset with small per-class calibration samples. Over 30 independent random data splits per condition, class-conditional calibration reduced the maximum-minimum per-class coverage gap relative to marginal calibration by a large, statistically significant margin under severe binary imbalance (mean gap reduction of 0.16\u20130.18 across models, Holm-adjusted p < 0.01 for all four model families), consistent with prior reports. Under mild binary imbalance the correction produced no significant fairness benefit while still inflating the average prediction-set size. Under mild three-class imbalance with small per-class calibration samples, class-conditional calibration significantly increased the coverage gap in three of four model families (Holm-adjusted p < 0.05), reversing its intended effect."),
  p("These results indicate that class-conditional conformal calibration is not a uniformly safe default: its benefit is concentrated in regimes of severe imbalance with adequate per-class calibration data, and it can be counterproductive when classes are numerous, close to balanced, or calibration data are scarce. The study contributes a joint, cross-model characterization of this trade-off and provides a fully reproducible experimental pipeline for auditing the fairness-efficiency behavior of conformal calibration schemes before deployment."),
  h2("Keywords"),
  p("conformal prediction; class-conditional calibration; Mondrian conformal prediction; class imbalance; uncertainty quantification; coverage fairness; tabular classification"),
];

// ================= INTRODUCTION =================
const introSection = [
  h1("1. Introduction"),
  p("Reliable uncertainty quantification is increasingly required of machine learning systems deployed in consequential settings such as clinical decision support, credit risk assessment, and industrial safety monitoring. Split conformal prediction (Vovk, Gammerman, & Shafer, 2005) has become a popular framework for this purpose because it produces prediction sets with a distribution-free, finite-sample marginal coverage guarantee: under the exchangeability assumption, the true label falls inside the predicted set with a user-specified probability, regardless of the underlying model or data distribution."),
  p("The guarantee that conformal prediction provides, however, is marginal: it holds on average over the full population, not for every subgroup or class individually. When the class distribution is imbalanced, a classifier's probability estimates for the minority class are typically less reliable, and a single, population-wide conformal threshold calibrated mostly on majority-class examples can leave the minority class substantially under-covered even though the overall, marginal coverage target is met. The broader problem of learning from imbalanced class distributions has been studied extensively in the machine learning literature (He & Garcia, 2009), and its specific interaction with conformal calibration has been documented across several application domains in recent work, including virtual screening (Tursunbadalov & Tursunbadalov, 2026), high-stakes tabular decision support spanning credit scoring, fraud detection, and healthcare (Singh, Srikantha, & Lakhanpal, 2026), and anomaly detection under extreme imbalance (Althani, 2026). The standard remedy is class-conditional, or Mondrian, conformal calibration (Vovk, Gammerman, & Shafer, 2005), which computes a separate calibration threshold for each class so that coverage is guaranteed conditionally on class membership rather than only on average."),
  p("These recent benchmarks consistently report that class-conditional calibration restores minority-class coverage, at a generally modest cost in prediction-set size. That is a valuable and increasingly well-replicated finding: Singh, Srikantha, and Lakhanpal (2026) report a 61.7 percentage-point average minority-coverage improvement over marginal calibration across 15 imbalanced tabular datasets and 7 model families; Tursunbadalov and Tursunbadalov (2026) report the same qualitative pattern on molecular screening datasets; and Althani (2026) reports a 37.65 percentage-point coverage improvement on a severely imbalanced (1:345) anomaly-detection benchmark. What has received comparatively little direct, controlled attention is the boundary of this benefit: class-conditional calibration itself relies on estimating a separate quantile from a smaller, class-restricted calibration sample, and quantile estimation from few observations is inherently noisy. It is therefore plausible that the correction helps when imbalance is severe and per-class calibration samples remain large in absolute terms, but becomes unreliable, or even actively harmful, when the number of classes grows relative to the calibration budget, or when imbalance is mild enough that the marginal method was not badly miscalibrated for any class in the first place."),
  p("This study investigates that boundary directly. Rather than proposing a new conformal-prediction algorithm, it conducts a controlled, cross-model empirical comparison of marginal and class-conditional split conformal prediction across three imbalance regimes that differ specifically in imbalance severity and per-class calibration sample size, holding the underlying conformal machinery (a standard least-ambiguous-set-valued, or LAC, nonconformity score) fixed. The central question is not whether class-conditional calibration can fix minority under-coverage \u2014 that is now well established \u2014 but under what conditions it does so reliably, and under what conditions it should not be treated as a safe default correction."),
];

// ================= RELATED WORK / GAP =================
const relatedWorkSection = [
  h1("2. Related Work and Research Gap"),
  h2("2.1 Split conformal prediction and set-valued classification"),
  p("Split (inductive) conformal prediction constructs prediction sets with finite-sample, distribution-free marginal coverage by calibrating a nonconformity score on a held-out calibration set (Vovk, Gammerman, & Shafer, 2005). For classification, the least-ambiguous-set-valued (LAC) score of Sadinle, Lei, and Wasserman (2019) defines the nonconformity of an example as one minus the predicted probability of its true class, and yields prediction sets that are optimal in expected size subject to the marginal coverage constraint under a correctly specified model. An alternative, adaptive prediction sets (APS) approach accumulates sorted class probabilities until a coverage threshold is reached (Romano, Sesia, & Cand\u00e8s, 2020), trading some efficiency for improved conditional coverage behavior. The present study uses the LAC score because it isolates the effect of the calibration scheme (marginal versus class-conditional) from the effect of the scoring function itself."),
  h2("2.2 Class-conditional (Mondrian) calibration"),
  p("Mondrian conformal prediction (Vovk, Gammerman, & Shafer, 2005) generalizes split conformal prediction by partitioning the calibration set into categories \u2014 here, class labels \u2014 and computing a separate threshold within each category, which yields a coverage guarantee conditional on category membership. Shi, Ong, and Leckie (2013) applied class-conditional calibration to multi-class classification and showed it can improve per-class coverage balance relative to the marginal scheme. Ding, Angelopoulos, Bates, Jordan, and Tibshirani (2023) studied the opposite end of the class-count spectrum, showing that with very many classes (hundreds to thousands, as in large-scale image classification), naively estimating one threshold per class from limited per-class calibration data degrades prediction-set efficiency, and proposed clustering classes to share calibration statistics as a remedy."),
  p("Several 2026 empirical studies document the practical severity of minority under-coverage under marginal calibration and confirm that class-conditional calibration corrects it: Singh, Srikantha, and Lakhanpal (2026) benchmarked fifteen imbalanced tabular datasets, seven model families, and three calibration techniques and found that class-conditional calibration improved minority-class coverage by an average of 61.7 percentage points over marginal calibration; Tursunbadalov and Tursunbadalov (2026) found the same qualitative pattern in a virtual-screening application, with marginal minority-class coverage falling as low as 4.2% on one dataset before correction; and Althani (2026) studied anomaly detection under a 1:345 imbalance ratio and found marginal calibration achieved only 52.94% anomaly-class coverage against a 90% target, restored to 90.59% by class-conditional calibration."),
  h2("2.3 Research gap"),
  p("Taken together, this literature establishes two boundary conditions under which the naive per-class thresholding used in class-conditional calibration is known to be fragile: an extremely large number of classes with correspondingly thin per-class calibration samples (Ding et al., 2023), and, at the other extreme, severe binary or few-class imbalance with generous per-class calibration samples, where the correction is reported to work well (Singh, Srikantha, & Lakhanpal, 2026; Tursunbadalov & Tursunbadalov, 2026; Althani, 2026). What has not been directly characterized is the intermediate case that is arguably most common in ordinary tabular practice: a small number of classes (as few as three) that are only mildly imbalanced, combined with modest, rather than abundant or vanishingly small, per-class calibration samples. It is not established from prior work whether class-conditional calibration remains a safe, low-risk default in this intermediate regime, or whether the added estimation variance from splitting an already-small calibration set outweighs any fairness gain when the marginal method was not badly miscalibrated to begin with. This study addresses that gap by directly contrasting three regimes \u2014 mild binary imbalance, severe binary imbalance, and mild multi-class imbalance with small per-class calibration samples \u2014 within a single controlled experimental design, holding the model families, nonconformity score, and evaluation protocol fixed across regimes so that differences in outcome can be attributed to the interaction between imbalance severity and calibration sample size rather than to confounding methodological differences."),
];

// ================= RESEARCH QUESTIONS / HYPOTHESES =================
const rqSection = [
  h1("3. Research Questions and Hypotheses"),
  p([tr("RQ1: ", { bold: true }), tr("Does class-conditional (Mondrian) conformal calibration reduce the per-class coverage gap relative to marginal conformal calibration under class imbalance in tabular classification, and is this effect consistent across model families?")]),
  p([tr("RQ2: ", { bold: true }), tr("Is the fairness benefit of class-conditional calibration moderated by imbalance severity and by the number of calibration examples available per class, to the point of reversing (increasing rather than decreasing the coverage gap) in some regimes?")]),
  p([tr("RQ3: ", { bold: true }), tr("What efficiency cost, measured as the change in average prediction-set size, accompanies any fairness gain (or loss) from class-conditional calibration, and does this cost differ systematically across model families or imbalance regimes?")]),
  p("Corresponding to these questions, the study tests three pre-specified hypotheses:"),
  p([tr("H1: ", { bold: true }), tr("Under severe class imbalance with an adequate absolute number of minority-class calibration examples, class-conditional calibration will significantly reduce the per-class coverage gap relative to marginal calibration, consistently across model families.")]),
  p([tr("H2: ", { bold: true }), tr("Under mild imbalance combined with small per-class calibration samples (as in a near-balanced multi-class setting), class-conditional calibration will fail to significantly reduce, and may significantly increase, the coverage gap relative to marginal calibration.")]),
  p([tr("H3: ", { bold: true }), tr("Any coverage-gap reduction achieved by class-conditional calibration will be accompanied by a measurable increase in average prediction-set size, reflecting a fairness\u2013efficiency trade-off.")]),
];

console.log("Part 1 sections built.");
module.exports = {
  titlePage, abstractSection, introSection, relatedWorkSection, rqSection,
  h1, h2, p, tr, caption, cell, makeTable, figureImage,
};

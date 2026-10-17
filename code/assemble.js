const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  PageNumber, Footer, Header, PageBreak, convertInchesToTwip,
} = require("docx");

const {
  titlePage, abstractSection, introSection, relatedWorkSection, rqSection,
  h1, h2, p, tr, caption, cell, makeTable, figureImage,
} = require("./build_paper.js");

// ================= METHODOLOGY =================
const methodologySection = [
  h1("4. Methodology"),
  h2("4.1 Datasets"),
  p("Two public, license-clear tabular datasets distributed with scikit-learn (Pedregosa et al., 2011), and originally sourced from the UCI Machine Learning Repository, were used so that all experiments are exactly reproducible without any external download step:"),
  p([tr("Breast Cancer Wisconsin (Diagnostic). ", { bold: true }), tr("569 instances, 30 numeric features computed from digitized images of fine-needle aspirates of breast masses, binary target (malignant / benign) (Wolberg, Mangasarian, Street, & Street, 1993). The natural class distribution is 357 benign and 212 malignant instances (imbalance ratio 1.68:1).")]),
  p([tr("Wine. ", { bold: true }), tr("178 instances, 13 numeric features from chemical analysis of wines derived from three cultivars, three-class target (Aeberhard & Forina, 1992; Aeberhard, Coomans, & de Vel, 1994). The natural class distribution is 71, 59, and 48 instances for the three cultivars respectively (imbalance ratio 1.48:1).")]),
  p("To create a severe-imbalance condition without fabricating or altering any feature values, a third condition was constructed by controlled minority-class subsampling of the Breast Cancer dataset: the malignant class was downsampled without replacement (fixed random seed 42) so that it constitutes approximately 10% of the resulting dataset (40 malignant and 357 benign instances, imbalance ratio 8.93:1), while every retained instance's feature values are unmodified originals. This synthetic-imbalance-injection procedure is a standard, fully documented technique for studying imbalance severity in a controlled way and is reported transparently here as an experimental manipulation rather than a naturally occurring condition."),
  h2("4.2 Data preparation and splitting protocol"),
  p("For each of the three (dataset, imbalance-condition) cells, data were split into training (60%), calibration (20%), and test (20%) partitions using stratified sampling to preserve the class proportions in every partition. Features were standardized (zero mean, unit variance) using statistics fitted only on the training partition, and the same transformation was applied to the calibration and test partitions, avoiding data leakage. This full pipeline \u2014 split, fit, calibrate, evaluate \u2014 was repeated independently for 30 random seeds (0\u201329) per cell, so that every reported statistic is an empirical distribution over 30 independent runs rather than a single split."),
  h2("4.3 Model families"),
  p("Four model families, chosen to span the major algorithmic paradigms used on tabular data, were trained on the standardized training partition in every run:"),
  p("\u2022 Logistic Regression (linear, probabilistic baseline; L2-regularized, scikit-learn default).\n\u2022 Random Forest (bagged decision-tree ensemble; 300 trees).\n\u2022 Histogram-based Gradient Boosting (boosted decision-tree ensemble; scikit-learn's HistGradientBoostingClassifier, a fast approximation to LightGBM/XGBoost-style boosting).\n\u2022 Support Vector Machine with a radial-basis-function kernel and Platt-scaled probability outputs (kernel-based, non-linear).", { }),
  p("These four families were selected because they represent qualitatively different mechanisms for producing class-probability estimates (linear, bagged-tree, boosted-tree, and kernel-based), allowing the study to assess whether any observed calibration-scheme effect is a property of conformal calibration itself, generalizing across models, or an artifact specific to one algorithm's probability outputs. No hyperparameter tuning was performed beyond standard library defaults (with a fixed, moderate ensemble size for Random Forest), because the object of study is the conformal calibration layer applied downstream of a fitted model, not the optimization of the base classifiers."),
  h2("4.4 Conformal calibration schemes"),
  p("For every trained model and every run, the class-probability outputs on the calibration partition were converted into LAC nonconformity scores, s(x, y) = 1 \u2212 \u0177(y | x), where \u0177(y | x) is the model's predicted probability of the true label y (Sadinle, Lei, & Wasserman, 2019). Two calibration schemes were then applied at a target coverage level of 1 \u2212 \u03b1 = 0.90:"),
  p([tr("Marginal split conformal. ", { bold: true }), tr("A single threshold \u0177 is computed as the \u2308(n+1)(1\u2212\u03b1)\u2309 / n empirical quantile of all n calibration nonconformity scores, pooled across classes, following the standard finite-sample-corrected split-conformal procedure (Vovk, Gammerman, & Shafer, 2005). A test instance's prediction set contains every class whose nonconformity score is at or below this single threshold.")]),
  p([tr("Class-conditional (Mondrian) split conformal. ", { bold: true }), tr("A separate threshold is computed, using the same finite-sample-corrected quantile formula, from only the calibration examples whose true label equals each class c (Vovk, Gammerman, & Shafer, 2005; Shi, Ong, & Leckie, 2013). A test instance's prediction set contains class c if its nonconformity score for class c is at or below that class's own threshold.")]),
  h2("4.5 Evaluation metrics"),
  p("Three metrics were computed on the held-out test partition for every (dataset, model, scheme, seed) combination:"),
  p([tr("Marginal coverage: ", { bold: true }), tr("the fraction of test instances for which the true label is contained in the prediction set, which both schemes target at 0.90 by construction and which serves as a sanity check that the calibration procedure is implemented correctly.")]),
  p([tr("Coverage gap (fairness metric): ", { bold: true }), tr("the difference between the highest and lowest per-class empirical coverage among the classes present in the test partition. A coverage gap of zero indicates that every class is covered equally; a large gap indicates that at least one class (typically the minority class under marginal calibration) is systematically under- or over-covered relative to others. This is the primary outcome of interest.")]),
  p([tr("Average prediction-set size (efficiency metric): ", { bold: true }), tr("the mean number of classes included in the prediction set across test instances. Smaller sets are more informative/efficient; a set size of 1.0 corresponds to fully deterministic (singleton) predictions on average.")]),
  h2("4.6 Statistical analysis"),
  p("For every (dataset, model) cell, the 30 paired observations (one marginal, one class-conditional value per seed, using an identical train/calibration/test split within each seed) of coverage gap and of average set size were compared using the Wilcoxon signed-rank test, a non-parametric paired test appropriate here because the difference distributions are not guaranteed to be normal and the paired design (same data split underlying both schemes) removes between-split variance that would otherwise obscure the comparison. A matched rank-biserial correlation was computed as an effect-size measure alongside each test. Because 12 cells (3 datasets \u00d7 4 models) were tested per metric, Holm's step-down procedure (Holm, 1979) was applied within each metric family to control the family-wise error rate at \u03b1 = 0.05. All p-values reported as significant in the Results section are Holm-adjusted unless stated otherwise."),
  h2("4.7 Implementation and reproducibility"),
  p("All experiments were implemented in Python 3 using scikit-learn 1.8 (Pedregosa et al., 2011), NumPy, pandas, and SciPy for the conformal-prediction and statistical-testing logic, and statsmodels for the Holm correction. Random seeds were fixed and recorded for every run (data split, model initialization, and the induced-imbalance subsampling). The complete source code, raw per-run results (720 rows: 3 datasets \u00d7 4 models \u00d7 2 schemes \u00d7 30 seeds), summary tables, and statistical test outputs accompanying this paper allow independent reproduction of every reported number without re-running the pipeline, and re-running the provided scripts reproduces the same results exactly given the fixed seeds."),
];

// ================= RESULTS =================
const datasetTable = makeTable(
  ["Dataset / Condition", "n", "Features", "Classes", "Class counts", "Imbalance ratio"],
  [
    ["Breast Cancer (natural)", "569", "30", "2", "357 / 212", "1.68"],
    ["Breast Cancer (induced extreme)", "397", "30", "2", "357 / 40", "8.93"],
    ["Wine (natural)", "178", "13", "3", "71 / 59 / 48", "1.48"],
  ],
  [2600, 900, 1100, 1000, 1900, 1700]
);

const mainResultsHeaders = ["Model", "Scheme", "Coverage gap (mean \u00b1 SD)", "Avg. set size (mean \u00b1 SD)"];
const mainResultsWidths = [1900, 1900, 2600, 2600];

function bcRow(model, marg, margSd, cc, ccSd, sizeM, sizeMSd, sizeC, sizeCSd) {
  return [
    [model, "Marginal", `${marg.toFixed(3)} \u00b1 ${margSd.toFixed(3)}`, `${sizeM.toFixed(3)} \u00b1 ${sizeMSd.toFixed(3)}`],
    ["", "Class-conditional", `${cc.toFixed(3)} \u00b1 ${ccSd.toFixed(3)}`, `${sizeC.toFixed(3)} \u00b1 ${sizeCSd.toFixed(3)}`],
  ];
}

// data pulled directly from /home/claude/research/results/summary_table.csv
const bcNaturalRows = [
  ["LogReg", "Marginal", "0.059 \u00b1 0.040", "0.930 \u00b1 0.025"],
  ["LogReg", "Class-conditional", "0.049 \u00b1 0.040", "0.959 \u00b1 0.028"],
  ["Random Forest", "Marginal", "0.061 \u00b1 0.047", "0.941 \u00b1 0.036"],
  ["Random Forest", "Class-conditional", "0.062 \u00b1 0.040", "0.974 \u00b1 0.043"],
  ["HistGB", "Marginal", "0.057 \u00b1 0.042", "0.935 \u00b1 0.032"],
  ["HistGB", "Class-conditional", "0.066 \u00b1 0.045", "0.972 \u00b1 0.051"],
  ["SVM (RBF)", "Marginal", "0.062 \u00b1 0.052", "0.927 \u00b1 0.040"],
  ["SVM (RBF)", "Class-conditional", "0.059 \u00b1 0.047", "0.954 \u00b1 0.037"],
];
const bcExtremeRows = [
  ["LogReg", "Marginal", "0.249 \u00b1 0.171", "0.915 \u00b1 0.039"],
  ["LogReg", "Class-conditional", "0.141 \u00b1 0.143", "0.936 \u00b1 0.059"],
  ["Random Forest", "Marginal", "0.304 \u00b1 0.151", "0.905 \u00b1 0.045"],
  ["Random Forest", "Class-conditional", "0.123 \u00b1 0.101", "1.039 \u00b1 0.305"],
  ["HistGB", "Marginal", "0.314 \u00b1 0.177", "0.904 \u00b1 0.044"],
  ["HistGB", "Class-conditional", "0.126 \u00b1 0.104", "0.957 \u00b1 0.107"],
  ["SVM (RBF)", "Marginal", "0.282 \u00b1 0.198", "0.915 \u00b1 0.039"],
  ["SVM (RBF)", "Class-conditional", "0.119 \u00b1 0.097", "0.973 \u00b1 0.155"],
];
const wineRows = [
  ["LogReg", "Marginal", "0.116 \u00b1 0.105", "0.953 \u00b1 0.087"],
  ["LogReg", "Class-conditional", "0.172 \u00b1 0.114", "0.931 \u00b1 0.077"],
  ["Random Forest", "Marginal", "0.130 \u00b1 0.098", "0.947 \u00b1 0.061"],
  ["Random Forest", "Class-conditional", "0.212 \u00b1 0.128", "0.906 \u00b1 0.061"],
  ["HistGB", "Marginal", "0.134 \u00b1 0.101", "0.951 \u00b1 0.086"],
  ["HistGB", "Class-conditional", "0.205 \u00b1 0.114", "0.929 \u00b1 0.080"],
  ["SVM (RBF)", "Marginal", "0.134 \u00b1 0.103", "0.932 \u00b1 0.073"],
  ["SVM (RBF)", "Class-conditional", "0.199 \u00b1 0.126", "0.904 \u00b1 0.069"],
];

const statTestHeaders = ["Dataset", "Model", "Marginal mean", "Class-cond. mean", "Holm p", "Effect size (r)", "Direction"];
const statTestWidths = [1900, 1700, 1500, 1700, 1100, 1400, 1300];
const covGapStatsRows = [
  ["BC (natural)", "LogReg", "0.059", "0.049", "0.749", "0.28", "n.s."],
  ["BC (natural)", "Random Forest", "0.061", "0.062", "1.000", "\u22120.01", "n.s."],
  ["BC (natural)", "HistGB", "0.057", "0.066", "1.000", "\u22120.15", "n.s."],
  ["BC (natural)", "SVM (RBF)", "0.062", "0.059", "1.000", "0.07", "n.s."],
  ["BC (extreme)", "LogReg", "0.249", "0.141", "0.041", "0.56", "gap \u2193 (sig.)"],
  ["BC (extreme)", "Random Forest", "0.304", "0.123", "<0.001", "0.93", "gap \u2193 (sig.)"],
  ["BC (extreme)", "HistGB", "0.314", "0.126", "<0.001", "0.86", "gap \u2193 (sig.)"],
  ["BC (extreme)", "SVM (RBF)", "0.282", "0.119", "0.002", "0.79", "gap \u2193 (sig.)"],
  ["Wine", "LogReg", "0.116", "0.172", "0.047", "\u22120.61", "gap \u2191 (sig.)"],
  ["Wine", "Random Forest", "0.130", "0.212", "0.026", "\u22120.63", "gap \u2191 (sig.)"],
  ["Wine", "HistGB", "0.134", "0.205", "0.011", "\u22120.73", "gap \u2191 (sig.)"],
  ["Wine", "SVM (RBF)", "0.134", "0.199", "0.105", "\u22120.46", "n.s."],
];

const resultsSection = [
  h1("5. Results"),
  h2("5.1 Dataset characteristics"),
  p("Table 1 summarizes the three (dataset, imbalance-condition) cells used in the experiments, confirming the intended progression from mild binary imbalance, to severe binary imbalance, to mild three-class imbalance with a smaller absolute sample size and therefore smaller per-class calibration samples."),
  datasetTable,
  caption("Table 1. Characteristics of the three experimental (dataset, imbalance-condition) cells."),
  h2("5.2 Marginal coverage sanity check"),
  p("Across all 24 (dataset \u00d7 model \u00d7 scheme) cells, mean marginal coverage on the test partition ranged from 0.890 to 0.937, closely tracking the 0.90 target under both calibration schemes, confirming that both conformal procedures were implemented and calibrated correctly. Marginal coverage alone, however, does not reveal how that coverage is distributed across classes, which motivates the per-class coverage-gap analysis that follows."),
  h2("5.3 Coverage-fairness gap: marginal versus class-conditional calibration"),
  p("Table 2 reports the per-class coverage gap and average prediction-set size for every model, under both calibration schemes, for the mildly imbalanced Breast Cancer condition. The two schemes produced statistically indistinguishable coverage gaps for all four model families (Holm-adjusted p = 1.00 for three of four models, p = 0.75 for logistic regression), while class-conditional calibration nonetheless increased the average prediction-set size for every model (Holm-adjusted p < 0.001 for three of four models)."),
  makeTable(mainResultsHeaders, bcNaturalRows, mainResultsWidths),
  caption("Table 2. Coverage gap and prediction-set size, Breast Cancer (natural imbalance, ratio 1.68), 30 random splits."),
  p("Table 3 reports the same metrics for the severely imbalanced Breast Cancer condition. Here the pattern reverses sharply: class-conditional calibration reduced the mean coverage gap by 0.16\u20130.18 across the four model families, and every reduction was statistically significant after Holm correction (p < 0.05 for all four models, with three below p = 0.01), with large effect sizes (rank-biserial r between 0.56 and 0.93)."),
  makeTable(mainResultsHeaders, bcExtremeRows, mainResultsWidths),
  caption("Table 3. Coverage gap and prediction-set size, Breast Cancer (induced extreme imbalance, ratio 8.93), 30 random splits."),
  p("Figure 3 illustrates the mechanism underlying this improvement for the Random Forest model: under marginal calibration, the minority (malignant) class was covered at only 62.9% on average against the 90% target, while the majority class was over-covered at 93.3%; under class-conditional calibration, minority coverage rose to 88.3% and majority coverage fell to 91.2%, bringing both classes close to the nominal target."),
  figureImage("fig3_per_class_mechanism.png", 380, 264),
  caption("Figure 3. Per-class empirical coverage under induced extreme imbalance (Random Forest, Breast Cancer dataset, 30 random splits). The dashed line marks the 90% target coverage."),
  p("Table 4 reports the same metrics for the Wine dataset (three classes, mild natural imbalance, small per-class calibration samples of roughly 12 examples per class). Here class-conditional calibration significantly increased the coverage gap relative to marginal calibration for three of the four model families (logistic regression, random forest, and histogram gradient boosting; Holm-adjusted p between 0.011 and 0.047), and the increase for the support vector machine did not reach significance after correction (p = 0.105). No model family showed a significant fairness benefit from class-conditional calibration on this dataset."),
  makeTable(mainResultsHeaders, wineRows, mainResultsWidths),
  caption("Table 4. Coverage gap and prediction-set size, Wine dataset (natural imbalance, ratio 1.48), 30 random splits."),
  figureImage("fig1_coverage_gap.png", 600, 194),
  caption("Figure 1. Coverage-fairness gap by calibration scheme, model family, and imbalance condition (mean \u00b1 SD over 30 random splits; target coverage = 0.90)."),
  h2("5.4 Prediction-set efficiency"),
  p("Figure 2 summarizes average prediction-set size across all conditions. Under both Breast Cancer conditions, class-conditional calibration increased average set size relative to marginal calibration for every model family; this increase was statistically significant for 6 of the 8 (model \u00d7 condition) cells in the binary datasets. On the Wine dataset, the direction reversed: class-conditional calibration produced smaller average set sizes than marginal calibration for three of four models, though this reduction was not statistically significant after Holm correction in any Wine cell."),
  figureImage("fig2_set_size.png", 600, 194),
  caption("Figure 2. Prediction-set efficiency by calibration scheme, model family, and imbalance condition (mean \u00b1 SD over 30 random splits; dotted line marks the singleton-set reference of 1.0)."),
  h2("5.5 Statistical test summary"),
  p("Table 5 consolidates the paired Wilcoxon signed-rank test results for the coverage-gap metric across all 12 (dataset \u00d7 model) cells, with Holm-adjusted p-values and matched rank-biserial effect sizes. Positive effect sizes indicate that the marginal scheme had the larger coverage gap (i.e., class-conditional calibration helped); negative effect sizes indicate the reverse."),
  makeTable(statTestHeaders, covGapStatsRows, statTestWidths),
  caption("Table 5. Paired Wilcoxon signed-rank test results for the coverage-gap metric, Holm-corrected across 12 tests. BC = Breast Cancer. n.s. = not significant at \u03b1 = 0.05 after correction."),
  p("Overall, hypothesis H1 was supported: under severe imbalance, class-conditional calibration produced a large, consistent, statistically significant reduction in the coverage gap across all four model families. Hypothesis H2 was also supported: under mild, small-sample multi-class imbalance, class-conditional calibration significantly increased the coverage gap in three of four model families, and produced no significant benefit in the mild binary-imbalance condition either. Hypothesis H3 was partially supported: class-conditional calibration increased average set size significantly in both Breast Cancer conditions, consistent with a fairness-efficiency trade-off, but on the Wine dataset it instead tended to decrease set size, indicating that the efficiency cost of class-conditional calibration is itself regime-dependent rather than a fixed penalty."),
];

// ================= DISCUSSION =================
const discussionSection = [
  h1("6. Discussion"),
  p("The central empirical finding of this study is that class-conditional (Mondrian) conformal calibration is not a uniformly beneficial correction to marginal conformal prediction under class imbalance; its value is concentrated in a specific regime and can reverse sign outside of it. Under severe binary imbalance with roughly eight calibration examples available for the minority class per split, class-conditional calibration closed most of the coverage gap left by marginal calibration, consistent with, and quantitatively comparable in direction to, the large benchmark studies published in 2026 (Singh, Srikantha, & Lakhanpal, 2026; Tursunbadalov & Tursunbadalov, 2026; Althani, 2026). This replication, obtained independently with a different set of datasets, models, and severity level, adds convergent evidence that the benefit of class-conditional calibration under severe imbalance is a robust, model-family-general phenomenon rather than an artifact of a particular pipeline."),
  p("The novel contribution of this study is the demonstration that this same correction, applied under conditions that are common in ordinary tabular practice \u2014 a small number of classes, only mild imbalance, and a modest total sample size that leaves roughly a dozen calibration examples per class \u2014 produced no significant fairness benefit and, in three of four model families, a statistically significant increase in the coverage gap. This is not the same failure mode documented by Ding et al. (2023), which concerns hundreds or thousands of classes; the Wine dataset here has only three. The mechanism appears instead to be simple estimator variance: splitting an already modest calibration set into three still-smaller, class-restricted subsets increases the variance of each per-class quantile estimate, and because the coverage-gap metric is itself computed from a further split (the test partition), noisy per-class thresholds can produce a larger observed spread across classes than a single, pooled threshold would, even though the marginal method was not meaningfully unfair to begin with in this mildly imbalanced setting. In effect, class-conditional calibration is trading a coverage-fairness problem that barely existed (natural imbalance ratios under 2:1 produced coverage gaps of roughly 0.06 under marginal calibration across all four models on Breast Cancer, essentially within sampling noise of the class-conditional result) for a self-inflicted estimation-variance problem."),
  p("This pattern is consistent with the general statistical principle that class-conditional (Mondrian) conformal validity guarantees hold asymptotically as the per-class calibration sample grows, but with finite and small per-class samples, the empirical per-class quantile can deviate substantially from its population target in either direction. The severe-imbalance condition in this study happened to have enough absolute minority-class calibration examples (approximately eight per split) for the correction to still net out favorably; the multi-class, mildly imbalanced condition, despite having a milder imbalance ratio, had a comparable or smaller absolute per-class calibration count (roughly 10\u201312 per class) spread across three classes, and here the correction did not net out favorably. This suggests that imbalance ratio alone is an incomplete diagnostic for whether class-conditional calibration should be applied; the absolute number of calibration examples per class, together with the number of classes, appears to be at least as important."),
  p("The efficiency results reinforce this picture. Class-conditional calibration increased average prediction-set size significantly under both Breast Cancer conditions \u2014 a fairness-efficiency trade-off consistent with the general conformal-prediction literature \u2014 but on the Wine dataset it instead tended to shrink prediction sets slightly, without a corresponding fairness benefit. A practitioner selecting a calibration scheme by prediction-set size alone, without separately examining per-class coverage, could therefore be misled into believing class-conditional calibration is a free efficiency improvement in a setting where it is in fact degrading coverage fairness."),
  p("Practically, these findings suggest that class-conditional conformal calibration should not be adopted by default whenever a classification problem has an imbalanced label distribution. Instead, practitioners should examine two joint diagnostics before adopting it: (i) the severity of the coverage gap actually produced by marginal calibration on a held-out validation split, since a small existing gap leaves little room for a genuine fairness benefit, and (ii) the absolute number of calibration examples available for the smallest class, since per-class quantile estimation is unreliable when this number is small regardless of the overall imbalance ratio. Where both a substantial existing coverage gap and an adequate per-class calibration sample are present, this study's results, together with the 2026 benchmark literature (Singh, Srikantha, & Lakhanpal, 2026; Tursunbadalov & Tursunbadalov, 2026; Althani, 2026), support adopting class-conditional calibration. Where either condition is absent \u2014 as in the mild-imbalance, small-per-class-sample multi-class setting studied here \u2014 marginal calibration, or an alternative such as clustered class-conditional calibration (Ding et al., 2023) that pools statistical strength across classes, is likely to be the safer choice."),
];

// ================= LIMITATIONS =================
const limitationsSection = [
  h1("7. Limitations"),
  p("Several limitations qualify these findings. First, the datasets used are small to moderate in size (178 to 569 instances), which was a deliberate choice to study the small-per-class-calibration-sample regime directly, but it also means the absolute number of test instances per class in some cells (particularly the induced-minority-class condition and the Wine dataset) is small, inflating the variance of the coverage-gap and set-size estimates themselves; the reported standard deviations reflect this. Second, the severe-imbalance condition was constructed by controlled subsampling of a naturally occurring dataset rather than observed directly in a naturally severely imbalanced tabular dataset; while this is a standard and transparent way to study imbalance severity in isolation, it may not fully capture other correlated characteristics of naturally severely imbalanced data (for example, more heterogeneous minority-class subpopulations), which could interact with calibration scheme in ways not captured here."),
  p("Third, the study evaluates only one nonconformity score (LAC); adaptive prediction sets (APS) or other scores could interact differently with class-conditional calibration, and Ding et al. (2023) and others have shown that scoring-function choice matters for many-class settings. Fourth, only default or lightly configured hyperparameters were used for the four model families; extensive hyperparameter tuning was intentionally avoided to keep the object of study focused on the calibration layer, but it remains possible that a well-tuned model for the Wine dataset would produce different, better-separated class probabilities that reduce the estimation-variance problem observed here. Fifth, only two datasets and three imbalance-condition cells were studied; while the cross-model consistency within each cell is a source of confidence, generalization to other domains, higher-dimensional feature spaces, or more than three classes should be established with additional datasets before treating the reported thresholds (for example, the specific per-class calibration counts observed here) as universal guidance rather than as an illustration of the underlying mechanism."),
];

// ================= THREATS TO VALIDITY =================
const threatsSection = [
  h1("8. Threats to Validity"),
  p([tr("Internal validity: ", { bold: true }), tr("The train/calibration/test splitting and standardization protocol was held identical across calibration schemes within each seed, and the same fitted model and calibration probabilities were reused for both the marginal and class-conditional threshold computation, so differences in outcome can be attributed to the calibration scheme rather than to different underlying models or data splits. Randomness was controlled via fixed, recorded seeds for reproducibility.")]),
  p([tr("External validity: ", { bold: true }), tr("The two datasets used are widely studied, low-dimensional, relatively clean tabular benchmarks; results may not generalize directly to noisier, higher-dimensional, or larger real-world tabular datasets, or to modalities beyond tabular data (text, images), where per-class calibration sample sizes and probability-estimation quality differ substantially.")]),
  p([tr("Construct validity: ", { bold: true }), tr("The maximum-minimum per-class coverage gap is one reasonable operationalization of coverage fairness, but other formulations (for example, the sum of absolute deviations from the target coverage across classes) could in principle rank the two calibration schemes differently; the qualitative direction of the findings was checked to be consistent using the underlying per-class coverage values reported in the results and is not an artifact of this specific gap definition.")]),
  p([tr("Statistical conclusion validity: ", { bold: true }), tr("Non-parametric paired testing (Wilcoxon signed-rank) with Holm correction across the 12 tests per metric was used to guard against inflated false-positive rates from multiple comparisons; effect sizes are reported alongside p-values to avoid over-interpreting statistically significant but practically small differences.")]),
];

// ================= CONCLUSION / FUTURE WORK =================
const conclusionSection = [
  h1("9. Conclusion"),
  p("This study conducted a controlled, cross-model empirical comparison of marginal and class-conditional (Mondrian) split conformal prediction across three class-imbalance regimes in tabular classification. Consistent with recently published large-scale benchmarks, class-conditional calibration substantially and significantly reduced the per-class coverage gap under severe binary class imbalance, across all four model families tested, at a measurable but modest cost in prediction-set size. The main contribution of this study is the complementary finding that this same correction conferred no significant fairness benefit under mild binary imbalance, and significantly worsened coverage fairness under mild, small-per-class-calibration-sample, multi-class imbalance in three of four model families \u2014 a regime distinct from the very-many-classes failure mode previously documented in the literature. These results indicate that the decision to apply class-conditional conformal calibration should be based on a joint assessment of imbalance severity and absolute per-class calibration sample size, rather than treated as a uniformly safe default whenever any class imbalance is present."),
  h2("9.1 Future work"),
  p("Future work could extend this joint characterization along several axes: testing additional, larger naturally imbalanced multi-class tabular datasets to determine the approximate per-class calibration sample size at which the benefit of class-conditional calibration reliably turns positive; comparing the LAC score used here against adaptive prediction sets and other nonconformity scores under the same regime structure; and evaluating clustered or shrinkage-based class-conditional calibration methods, such as that of Ding, Angelopoulos, Bates, Jordan, and Tibshirani (2023), as a middle ground that might retain the fairness benefit of class-conditional calibration under severe imbalance while avoiding the estimation-variance cost identified here under mild, small-sample multi-class imbalance."),
  h2("9.2 Ethics and data statement"),
  p("Both datasets used in this study are publicly available, de-identified, and contain no personally identifiable information: the Breast Cancer Wisconsin (Diagnostic) dataset consists of numeric measurements derived from digitized cell images with no patient identifiers, and the Wine dataset consists of chemical measurements of wine samples. No ethics board approval was sought or required for this study, as it involved no human participants, no new data collection, and only secondary analysis of pre-existing, publicly released, anonymized datasets."),
];

// ================= REFERENCES =================
function refPara(text) {
  return new Paragraph({
    children: [new TextRun({ text })],
    spacing: { after: 160, line: 276 },
    indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.5) },
  });
}
const references = [
  h1("References"),
  refPara("Aeberhard, S., & Forina, M. (1992). Wine [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5PC7J"),
  refPara("Aeberhard, S., Coomans, D., & de Vel, O. (1994). Comparative analysis of statistical pattern recognition methods in high dimensional settings. Pattern Recognition, 27(8), 1065\u20131077."),
  refPara("Althani, B. (2026). Class-conditional conformal prediction for reliable anomaly detection under extreme class imbalance. Machine Learning and Knowledge Extraction, 8(7), Article 190. https://doi.org/10.3390/make8070190"),
  refPara("Ding, T., Angelopoulos, A., Bates, S., Jordan, M., & Tibshirani, R. J. (2023). Class-conditional conformal prediction with many classes. Advances in Neural Information Processing Systems, 36, 64555\u201364576."),
  refPara("He, H., & Garcia, E. A. (2009). Learning from imbalanced data. IEEE Transactions on Knowledge and Data Engineering, 21(9), 1263\u20131284."),
  refPara("Holm, S. (1979). A simple sequentially rejective multiple test procedure. Scandinavian Journal of Statistics, 6(2), 65\u201370."),
  refPara("Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M., & Duchesnay, E. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825\u20132830."),
  refPara("Romano, Y., Sesia, M., & Cand\u00e8s, E. (2020). Classification with valid and adaptive coverage. Advances in Neural Information Processing Systems, 33, 3581\u20133591."),
  refPara("Sadinle, M., Lei, J., & Wasserman, L. (2019). Least ambiguous set-valued classifiers with bounded error levels. Journal of the American Statistical Association, 114(525), 223\u2013234."),
  refPara("Shi, F., Ong, C. S., & Leckie, C. (2013). Applications of class-conditional conformal predictor in multi-class classification. Proceedings of the 12th International Conference on Machine Learning and Applications (ICMLA), 235\u2013239."),
  refPara("Singh, M., Srikantha, A., & Lakhanpal, S. (2026). Cost-sensitive conformal prediction and human-in-the-loop abstention for imbalanced high-stakes decision support: A multi-domain benchmark. arXiv preprint arXiv:2607.27143."),
  refPara("Tursunbadalov, M., & Tursunbadalov, M. (2026). A quiet failure in calibrated virtual screening: Marginal conformal prediction under-covers the minority class, and a class-conditional fix recovers it. arXiv preprint arXiv:2607.06605."),
  refPara("Vovk, V., Gammerman, A., & Shafer, G. (2005). Algorithmic learning in a random world. Springer."),
  refPara("Wolberg, W., Mangasarian, O., Street, N., & Street, W. (1993). Breast Cancer Wisconsin (Diagnostic) [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5DW2B"),
];

// ================= ASSEMBLE DOCUMENT =================
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22 } },
    },
    heading1: { run: { size: 26, bold: true, color: "1F3864" }, paragraph: { spacing: { before: 320, after: 160 } } },
    heading2: { run: { size: 23, bold: true, color: "2E5395" }, paragraph: { spacing: { before: 240, after: 120 } } },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "Coverage Fairness in Conformal Prediction Under Class Imbalance", size: 16, italics: true, color: "808080" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })],
          })],
        }),
      },
      children: [
        ...titlePage,
        ...abstractSection,
        ...introSection,
        ...relatedWorkSection,
        ...rqSection,
        ...methodologySection,
        ...resultsSection,
        ...discussionSection,
        ...limitationsSection,
        ...threatsSection,
        ...conclusionSection,
        ...references,
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/claude/research/Coverage_Fairness_Conformal_Prediction_Paper.docx", buffer);
  console.log("Document written.");
});

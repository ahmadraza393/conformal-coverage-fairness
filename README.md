# Coverage Fairness in Conformal Prediction Under Class Imbalance

Reproducibility package for the paper *"Coverage Fairness in Conformal Prediction
Under Class Imbalance: When Class-Conditional Calibration Helps and When It
Backfires"* (Ahmad Raza, Independent Researcher).

## What this study does

Compares **marginal** vs. **class-conditional (Mondrian)** split conformal
prediction (LAC nonconformity score, target coverage 90%) across:

- 4 model families: Logistic Regression, Random Forest, Histogram Gradient
  Boosting, RBF-kernel SVM
- 3 imbalance regimes: Breast Cancer Wisconsin (Diagnostic) natural imbalance
  (ratio 1.68), Breast Cancer induced extreme imbalance (ratio 8.93, via
  documented minority-class subsampling, seed 42), and Wine natural imbalance
  (3 classes, ratio 1.48)
- 30 independent random train/calibration/test splits (seeds 0–29) per cell

and evaluates per-class coverage gap (fairness) and average prediction-set
size (efficiency), with paired Wilcoxon signed-rank tests and Holm correction
across the 12 (dataset × model) cells per metric.

## Directory structure

```
code/
  run_experiments.py   # main experimental pipeline (fits models, runs both
                        # conformal calibration schemes, computes metrics,
                        # runs statistical tests) -> writes results/*.csv
  make_figures.py       # generates all three figures from results/*.csv
  build_paper.js         # docx-js: title page, abstract, intro, related work, RQs
  assemble.js             # docx-js: methodology, results, discussion, references;
                           # assembles and writes the final .docx
results/
  raw_results.csv       # 720 rows: 3 datasets x 4 models x 2 schemes x 30 seeds
  summary_table.csv     # per (dataset, model, scheme) mean/SD of all metrics
  statistical_tests.csv # paired Wilcoxon tests, Holm-adjusted p-values, effect sizes
  mechanism_breast_cancer_extreme_rf.csv  # per-seed per-class coverage (Figure 3 data)
figures/
  fig1_coverage_gap.png
  fig2_set_size.png
  fig3_per_class_mechanism.png
Coverage_Fairness_Conformal_Prediction_Paper.docx   # final paper
requirements.txt
README.md (this file)
```

## Reproducing the results

```bash
pip install -r requirements.txt --break-system-packages
cd code
python3 run_experiments.py   # writes ../results/*.csv (deterministic given fixed seeds)
python3 make_figures.py      # writes ../figures/*.png
node assemble.js             # writes ../Coverage_Fairness_Conformal_Prediction_Paper.docx
```

All random seeds (data splitting, model initialization, and the induced-imbalance
subsampling) are fixed and recorded, so re-running `run_experiments.py` reproduces
`raw_results.csv` exactly, and every number quoted in the paper can be traced back
to a cell in `summary_table.csv` or `statistical_tests.csv`.

## Datasets

Both datasets are loaded from scikit-learn's bundled copies (no external
download required) and are originally sourced from the UCI Machine Learning
Repository:

- Breast Cancer Wisconsin (Diagnostic): Wolberg, Mangasarian, Street, & Street
  (1993), https://doi.org/10.24432/C5DW2B
- Wine: Aeberhard & Forina (1992), https://doi.org/10.24432/C5PC7J

## Ethics

Both datasets are public, de-identified, and contain no personally
identifiable information. No new data collection or human-subjects research
was involved.

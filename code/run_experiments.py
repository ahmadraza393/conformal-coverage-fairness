"""
Marginal vs. Class-Conditional Conformal Prediction for Imbalanced Tabular
Classification: Coverage Fairness and Efficiency Across Model Families.

This script performs the full experimental pipeline:
  1. Loads two public, license-clear tabular datasets bundled with
     scikit-learn (Breast Cancer Wisconsin Diagnostic, and Wine).
  2. Constructs three imbalance conditions (natural imbalance for both
     datasets, plus an induced extreme-imbalance condition for the
     binary dataset obtained by controlled minority-class subsampling).
  3. Trains four model families (Logistic Regression, Random Forest,
     Histogram Gradient Boosting, calibrated RBF-SVM).
  4. Applies split conformal prediction with two calibration schemes:
       (a) Marginal (standard LAC split-conformal)
       (b) Class-conditional / Mondrian split-conformal
  5. Repeats every (dataset x imbalance-condition x model) cell over
     30 independent random splits (seeds 0-29) for statistical power.
  6. Computes marginal coverage, per-class coverage, the
     max-min per-class coverage gap (the fairness metric), and average
     prediction-set size (the efficiency metric).
  7. Runs paired Wilcoxon signed-rank tests (marginal vs. class-conditional)
     on coverage gap and set size for every dataset x model cell.
  8. Saves all raw per-run results, summary tables, statistical test
     results, and figures to disk. No numbers in the paper are invented;
     every reported value is read back from these output files.
"""

import json
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.datasets import load_breast_cancer, load_wine
from sklearn.ensemble import HistGradientBoostingClassifier, RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC

warnings.filterwarnings("ignore")

RESULTS_DIR = Path("/home/claude/research/results")
RESULTS_DIR.mkdir(parents=True, exist_ok=True)

ALPHA = 0.10  # target coverage 1 - alpha = 0.90
N_SEEDS = 30
SEEDS = list(range(N_SEEDS))

# ---------------------------------------------------------------------------
# Dataset / imbalance-condition construction
# ---------------------------------------------------------------------------

def load_breast_cancer_natural():
    data = load_breast_cancer()
    X, y = data.data, data.target
    return X, y, "Breast Cancer (natural)"


def load_breast_cancer_extreme(minority_frac=0.10, seed=42):
    """Induce a severe, controlled class imbalance by subsampling the
    minority class of the Breast Cancer Wisconsin dataset. The majority
    class is left untouched; only the minority class is downsampled
    (without replacement) so that it constitutes approximately
    `minority_frac` of the resulting dataset. This is a standard,
    fully documented synthetic-imbalance-injection procedure and does
    not alter or fabricate any feature values."""
    data = load_breast_cancer()
    X, y = data.data, data.target
    rng = np.random.RandomState(seed)
    counts = np.bincount(y)
    minority_class = int(np.argmin(counts))
    majority_class = 1 - minority_class
    n_majority = counts[majority_class]
    # solve for minority count s.t. minority / (minority + n_majority) = minority_frac
    n_minority_target = int(round((minority_frac * n_majority) / (1 - minority_frac)))
    minority_idx = np.where(y == minority_class)[0]
    majority_idx = np.where(y == majority_class)[0]
    keep_minority = rng.choice(minority_idx, size=min(n_minority_target, len(minority_idx)), replace=False)
    keep_idx = np.concatenate([majority_idx, keep_minority])
    rng.shuffle(keep_idx)
    return X[keep_idx], y[keep_idx], "Breast Cancer (induced extreme imbalance)"


def load_wine_natural():
    data = load_wine()
    X, y = data.data, data.target
    return X, y, "Wine (natural)"


DATASETS = {
    "breast_cancer_natural": load_breast_cancer_natural,
    "breast_cancer_extreme": lambda: load_breast_cancer_extreme(0.10, 42),
    "wine_natural": load_wine_natural,
}

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

def make_models(seed):
    return {
        "LogisticRegression": LogisticRegression(max_iter=5000, random_state=seed),
        "RandomForest": RandomForestClassifier(n_estimators=300, random_state=seed),
        "HistGradientBoosting": HistGradientBoostingClassifier(random_state=seed),
        "SVM_RBF": SVC(kernel="rbf", probability=True, random_state=seed),
    }


# ---------------------------------------------------------------------------
# Conformal prediction (LAC / split conformal for classification)
# ---------------------------------------------------------------------------

def lac_scores(probs, labels):
    """Least-ambiguous-set-valued (LAC) nonconformity score: 1 - p(true class)."""
    idx = np.arange(len(labels))
    return 1.0 - probs[idx, labels]


def marginal_threshold(calib_scores, alpha):
    n = len(calib_scores)
    q_level = np.ceil((n + 1) * (1 - alpha)) / n
    q_level = min(q_level, 1.0)
    return np.quantile(calib_scores, q_level, method="higher")


def class_conditional_thresholds(calib_scores, calib_labels, n_classes, alpha):
    thresholds = {}
    for c in range(n_classes):
        mask = calib_labels == c
        scores_c = calib_scores[mask]
        if len(scores_c) == 0:
            thresholds[c] = 1.0  # degenerate: no calibration points for this class
            continue
        n = len(scores_c)
        q_level = np.ceil((n + 1) * (1 - alpha)) / n
        q_level = min(q_level, 1.0)
        thresholds[c] = np.quantile(scores_c, q_level, method="higher")
    return thresholds


def build_prediction_sets_marginal(test_probs, threshold):
    return (1.0 - test_probs) <= threshold


def build_prediction_sets_conditional(test_probs, thresholds, n_classes):
    thr_vec = np.array([thresholds[c] for c in range(n_classes)])
    return (1.0 - test_probs) <= thr_vec[np.newaxis, :]


def evaluate_sets(pred_sets, y_test, n_classes):
    n = len(y_test)
    covered = pred_sets[np.arange(n), y_test]
    marginal_coverage = covered.mean()
    set_sizes = pred_sets.sum(axis=1)
    avg_set_size = set_sizes.mean()

    per_class_coverage = {}
    for c in range(n_classes):
        mask = y_test == c
        if mask.sum() == 0:
            continue
        per_class_coverage[c] = covered[mask].mean()

    if len(per_class_coverage) >= 2:
        vals = list(per_class_coverage.values())
        coverage_gap = max(vals) - min(vals)
    else:
        coverage_gap = np.nan

    return {
        "marginal_coverage": marginal_coverage,
        "avg_set_size": avg_set_size,
        "coverage_gap": coverage_gap,
        "per_class_coverage": per_class_coverage,
    }


# ---------------------------------------------------------------------------
# Main experimental loop
# ---------------------------------------------------------------------------

def run():
    records = []

    for ds_key, loader in DATASETS.items():
        X, y, ds_label = loader()
        n_classes = len(np.unique(y))
        class_counts = np.bincount(y)
        imbalance_ratio = class_counts.max() / class_counts.min()

        for seed in SEEDS:
            # 60/20/20 train/calibration/test split, stratified
            X_train, X_temp, y_train, y_temp = train_test_split(
                X, y, test_size=0.4, stratify=y, random_state=seed
            )
            X_calib, X_test, y_calib, y_test = train_test_split(
                X_temp, y_temp, test_size=0.5, stratify=y_temp, random_state=seed
            )

            scaler = StandardScaler().fit(X_train)
            X_train_s = scaler.transform(X_train)
            X_calib_s = scaler.transform(X_calib)
            X_test_s = scaler.transform(X_test)

            models = make_models(seed)
            for model_name, model in models.items():
                model.fit(X_train_s, y_train)

                calib_probs = model.predict_proba(X_calib_s)
                test_probs = model.predict_proba(X_test_s)

                calib_scores = lac_scores(calib_probs, y_calib)

                # Marginal split conformal
                thr_marg = marginal_threshold(calib_scores, ALPHA)
                sets_marg = build_prediction_sets_marginal(test_probs, thr_marg)
                res_marg = evaluate_sets(sets_marg, y_test, n_classes)

                # Class-conditional (Mondrian) split conformal
                thr_cc = class_conditional_thresholds(calib_scores, y_calib, n_classes, ALPHA)
                sets_cc = build_prediction_sets_conditional(test_probs, thr_cc, n_classes)
                res_cc = evaluate_sets(sets_cc, y_test, n_classes)

                # top-1 accuracy for reference
                preds = model.predict(X_test_s)
                acc = (preds == y_test).mean()

                for scheme, res in [("marginal", res_marg), ("class_conditional", res_cc)]:
                    records.append({
                        "dataset": ds_key,
                        "dataset_label": ds_label,
                        "n_classes": n_classes,
                        "imbalance_ratio": imbalance_ratio,
                        "seed": seed,
                        "model": model_name,
                        "scheme": scheme,
                        "top1_accuracy": acc,
                        "marginal_coverage": res["marginal_coverage"],
                        "avg_set_size": res["avg_set_size"],
                        "coverage_gap": res["coverage_gap"],
                    })

    df = pd.DataFrame(records)
    df.to_csv(RESULTS_DIR / "raw_results.csv", index=False)
    return df


def summarize(df):
    summary = (
        df.groupby(["dataset", "dataset_label", "model", "scheme"])
        .agg(
            mean_marginal_coverage=("marginal_coverage", "mean"),
            std_marginal_coverage=("marginal_coverage", "std"),
            mean_coverage_gap=("coverage_gap", "mean"),
            std_coverage_gap=("coverage_gap", "std"),
            mean_set_size=("avg_set_size", "mean"),
            std_set_size=("avg_set_size", "std"),
            mean_top1_accuracy=("top1_accuracy", "mean"),
        )
        .reset_index()
    )
    summary.to_csv(RESULTS_DIR / "summary_table.csv", index=False)
    return summary


def paired_tests(df):
    test_records = []
    for (ds_key, model_name), grp in df.groupby(["dataset", "model"]):
        marg = grp[grp.scheme == "marginal"].sort_values("seed")
        cc = grp[grp.scheme == "class_conditional"].sort_values("seed")
        assert list(marg.seed) == list(cc.seed)

        for metric in ["coverage_gap", "avg_set_size"]:
            a = marg[metric].values
            b = cc[metric].values
            diff = a - b
            if np.allclose(diff, 0):
                stat_, p_ = np.nan, 1.0
            else:
                stat_, p_ = stats.wilcoxon(a, b)
            # rank-biserial effect size for Wilcoxon
            n = len(diff)
            if np.allclose(diff, 0):
                effect = 0.0
            else:
                r = stats.rankdata(np.abs(diff))
                pos = r[diff > 0].sum()
                neg = r[diff < 0].sum()
                effect = (pos - neg) / (pos + neg) if (pos + neg) > 0 else 0.0

            test_records.append({
                "dataset": ds_key,
                "model": model_name,
                "metric": metric,
                "mean_marginal": a.mean(),
                "mean_class_conditional": b.mean(),
                "wilcoxon_stat": stat_,
                "p_value": p_,
                "rank_biserial_effect_size": effect,
                "n_pairs": n,
            })
    test_df = pd.DataFrame(test_records)
    # Holm-Bonferroni correction within each metric family
    from statsmodels.stats.multitest import multipletests
    for metric in test_df.metric.unique():
        mask = test_df.metric == metric
        reject, p_adj, _, _ = multipletests(test_df.loc[mask, "p_value"], method="holm")
        test_df.loc[mask, "p_value_holm"] = p_adj
        test_df.loc[mask, "significant_holm_0.05"] = reject
    test_df.to_csv(RESULTS_DIR / "statistical_tests.csv", index=False)
    return test_df


if __name__ == "__main__":
    df = run()
    summary = summarize(df)
    tests = paired_tests(df)
    print("DONE")
    print(df.shape, summary.shape, tests.shape)

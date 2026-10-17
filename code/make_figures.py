import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

FIG_DIR = "/home/claude/research/figures"
plt.rcParams.update({"font.size": 10, "figure.dpi": 150})

df = pd.read_csv("/home/claude/research/results/raw_results.csv")
summary = pd.read_csv("/home/claude/research/results/summary_table.csv")

DATASET_ORDER = ["breast_cancer_natural", "breast_cancer_extreme", "wine_natural"]
DATASET_TITLES = {
    "breast_cancer_natural": "Breast Cancer\n(natural, ratio 1.68)",
    "breast_cancer_extreme": "Breast Cancer\n(induced extreme, ratio 8.93)",
    "wine_natural": "Wine\n(natural, ratio 1.48)",
}
MODEL_ORDER = ["LogisticRegression", "RandomForest", "HistGradientBoosting", "SVM_RBF"]

# ---------------------------------------------------------------------
# Figure 1: Coverage gap, marginal vs class-conditional, grouped bars
# ---------------------------------------------------------------------
fig, axes = plt.subplots(1, 3, figsize=(13, 4.2), sharey=True)
width = 0.35
x = np.arange(len(MODEL_ORDER))

for ax, ds in zip(axes, DATASET_ORDER):
    sub = summary[summary.dataset == ds]
    marg = [sub[(sub.model == m) & (sub.scheme == "marginal")].mean_coverage_gap.values[0] for m in MODEL_ORDER]
    marg_sd = [sub[(sub.model == m) & (sub.scheme == "marginal")].std_coverage_gap.values[0] for m in MODEL_ORDER]
    cc = [sub[(sub.model == m) & (sub.scheme == "class_conditional")].mean_coverage_gap.values[0] for m in MODEL_ORDER]
    cc_sd = [sub[(sub.model == m) & (sub.scheme == "class_conditional")].std_coverage_gap.values[0] for m in MODEL_ORDER]

    ax.bar(x - width/2, marg, width, yerr=marg_sd, capsize=3, label="Marginal", color="#4C72B0")
    ax.bar(x + width/2, cc, width, yerr=cc_sd, capsize=3, label="Class-conditional", color="#DD8452")
    ax.set_title(DATASET_TITLES[ds], fontsize=10)
    ax.set_xticks(x)
    ax.set_xticklabels(["LogReg", "RF", "HGB", "SVM"], rotation=0)
    ax.set_xlabel("Model")

axes[0].set_ylabel("Max\u2212min per-class coverage gap")
axes[0].legend(loc="upper right", fontsize=8)
fig.suptitle("Figure 1. Coverage-fairness gap by calibration scheme, model, and imbalance condition\n(mean \u00b1 SD over 30 random splits; target coverage = 0.90)", fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.90])
fig.savefig(f"{FIG_DIR}/fig1_coverage_gap.png")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure 2: Average prediction-set size (efficiency), grouped bars
# ---------------------------------------------------------------------
fig, axes = plt.subplots(1, 3, figsize=(13, 4.2), sharey=True)
for ax, ds in zip(axes, DATASET_ORDER):
    sub = summary[summary.dataset == ds]
    marg = [sub[(sub.model == m) & (sub.scheme == "marginal")].mean_set_size.values[0] for m in MODEL_ORDER]
    marg_sd = [sub[(sub.model == m) & (sub.scheme == "marginal")].std_set_size.values[0] for m in MODEL_ORDER]
    cc = [sub[(sub.model == m) & (sub.scheme == "class_conditional")].mean_set_size.values[0] for m in MODEL_ORDER]
    cc_sd = [sub[(sub.model == m) & (sub.scheme == "class_conditional")].std_set_size.values[0] for m in MODEL_ORDER]

    ax.bar(x - width/2, marg, width, yerr=marg_sd, capsize=3, label="Marginal", color="#4C72B0")
    ax.bar(x + width/2, cc, width, yerr=cc_sd, capsize=3, label="Class-conditional", color="#DD8452")
    ax.set_title(DATASET_TITLES[ds], fontsize=10)
    ax.set_xticks(x)
    ax.set_xticklabels(["LogReg", "RF", "HGB", "SVM"], rotation=0)
    ax.set_xlabel("Model")
    ax.axhline(1.0, color="grey", linestyle=":", linewidth=1)

axes[0].set_ylabel("Average prediction-set size")
axes[0].legend(loc="upper left", fontsize=8)
fig.suptitle("Figure 2. Prediction-set efficiency by calibration scheme, model, and imbalance condition\n(mean \u00b1 SD over 30 random splits; dotted line = singleton-set reference)", fontsize=10)
fig.tight_layout(rect=[0, 0, 1, 0.90])
fig.savefig(f"{FIG_DIR}/fig2_set_size.png")
plt.close(fig)

# ---------------------------------------------------------------------
# Figure 3: Mechanism illustration - per-class coverage under extreme
# imbalance, Random Forest, marginal vs class-conditional
# ---------------------------------------------------------------------
import run_experiments as R
import sys
sys.path.insert(0, "/home/claude/research/code")

X, y, _ = R.load_breast_cancer_extreme(0.10, 42)
n_classes = 2
maj_cov_marg, min_cov_marg, maj_cov_cc, min_cov_cc = [], [], [], []

from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

for seed in R.SEEDS:
    X_train, X_temp, y_train, y_temp = train_test_split(X, y, test_size=0.4, stratify=y, random_state=seed)
    X_calib, X_test, y_calib, y_test = train_test_split(X_temp, y_temp, test_size=0.5, stratify=y_temp, random_state=seed)
    scaler = StandardScaler().fit(X_train)
    model = R.make_models(seed)["RandomForest"]
    model.fit(scaler.transform(X_train), y_train)
    calib_probs = model.predict_proba(scaler.transform(X_calib))
    test_probs = model.predict_proba(scaler.transform(X_test))
    calib_scores = R.lac_scores(calib_probs, y_calib)

    thr_m = R.marginal_threshold(calib_scores, R.ALPHA)
    sets_m = R.build_prediction_sets_marginal(test_probs, thr_m)
    res_m = R.evaluate_sets(sets_m, y_test, n_classes)

    thr_cc = R.class_conditional_thresholds(calib_scores, y_calib, n_classes, R.ALPHA)
    sets_cc = R.build_prediction_sets_conditional(test_probs, thr_cc, n_classes)
    res_cc = R.evaluate_sets(sets_cc, y_test, n_classes)

    counts = np.bincount(y)
    minority_class = int(np.argmin(counts))
    majority_class = 1 - minority_class

    maj_cov_marg.append(res_m["per_class_coverage"].get(majority_class, np.nan))
    min_cov_marg.append(res_m["per_class_coverage"].get(minority_class, np.nan))
    maj_cov_cc.append(res_cc["per_class_coverage"].get(majority_class, np.nan))
    min_cov_cc.append(res_cc["per_class_coverage"].get(minority_class, np.nan))

fig, ax = plt.subplots(figsize=(6.5, 4.5))
groups = ["Majority\n(marginal)", "Minority\n(marginal)", "Majority\n(class-cond.)", "Minority\n(class-cond.)"]
data = [maj_cov_marg, min_cov_marg, maj_cov_cc, min_cov_cc]
colors = ["#4C72B0", "#4C72B0", "#DD8452", "#DD8452"]
bp = ax.boxplot(data, labels=groups, patch_artist=True, widths=0.55)
for patch, c in zip(bp["boxes"], colors):
    patch.set_facecolor(c)
    patch.set_alpha(0.7)
ax.axhline(0.90, color="black", linestyle="--", linewidth=1, label="Target coverage (0.90)")
ax.set_ylabel("Per-class empirical coverage")
ax.set_title("Figure 3. Per-class coverage under induced extreme imbalance\n(Random Forest, breast cancer dataset, 30 random splits)", fontsize=10)
ax.legend(fontsize=8, loc="lower right")
fig.tight_layout()
fig.savefig(f"{FIG_DIR}/fig3_per_class_mechanism.png")
plt.close(fig)

# save the mechanism data for exact figures-in-text numbers
mech_df = pd.DataFrame({
    "seed": R.SEEDS,
    "majority_coverage_marginal": maj_cov_marg,
    "minority_coverage_marginal": min_cov_marg,
    "majority_coverage_class_conditional": maj_cov_cc,
    "minority_coverage_class_conditional": min_cov_cc,
})
mech_df.to_csv("/home/claude/research/results/mechanism_breast_cancer_extreme_rf.csv", index=False)

print("Mean minority coverage (marginal):", np.mean(min_cov_marg))
print("Mean minority coverage (class-conditional):", np.mean(min_cov_cc))
print("Mean majority coverage (marginal):", np.mean(maj_cov_marg))
print("Mean majority coverage (class-conditional):", np.mean(maj_cov_cc))
print("Figures saved.")

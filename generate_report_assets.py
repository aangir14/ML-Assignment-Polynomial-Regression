"""
generate_report_assets.py
Generates all plots needed for the report and saves them as PNGs.
Run this BEFORE building the report DOCX.
"""

import os, warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import Ridge
from sklearn.model_selection import KFold, cross_validate
from sklearn.metrics import mean_squared_error, r2_score
warnings.filterwarnings("ignore")

# ── paths ──────────────────────────────────────────────────────
DATA_DIR  = r"c:\Users\Aangir Doshi\Downloads\BT2024078 (8)\BT2024078"
PLOT_DIR  = os.path.join(DATA_DIR, "report_plots")
os.makedirs(PLOT_DIR, exist_ok=True)

ROLLNO = "BT2024078"
def dpath(split, var): return os.path.join(DATA_DIR, f"{ROLLNO}_{split}_{var}.csv")

# ── load data ──────────────────────────────────────────────────
train1 = pd.read_csv(dpath("train","var1")); test1 = pd.read_csv(dpath("test","var1"))
train2 = pd.read_csv(dpath("train","var2")); test2 = pd.read_csv(dpath("test","var2"))
feat1 = ["x1","x2","x3","x4","x5","x6"]; feat2 = ["x1","x2","x3"]
X1,y1 = train1[feat1].values, train1["y"].values
X2,y2 = train2[feat2].values, train2["y"].values

# ── final configs ──────────────────────────────────────────────
CFG = {
    "var1": {"deg":5, "alpha":2.477076, "feat":feat1, "X":X1, "y":y1,
             "label":"Var1 – Steam Turbine Power Score"},
    "var2": {"deg":12,"alpha":0.231013, "feat":feat2, "X":X2, "y":y2,
             "label":"Var2 – Thermal Anomaly Score"},
}
KF = KFold(n_splits=5, shuffle=True, random_state=42)

STYLE = dict(dpi=150, bbox_inches="tight")
PLT_COLOR = "#2E86AB"
PLT_COLOR2 = "#E84855"
GRID_COLOR = "#e0e0e0"

def savefig(name):
    path = os.path.join(PLOT_DIR, name)
    plt.savefig(path, **STYLE)
    plt.close("all")
    print(f"  saved: {name}")
    return path

# ─────────────────────────────────────────────────────────────────
# FIGURE 1 – Degree vs CV MSE for both problems (side by side)
# ─────────────────────────────────────────────────────────────────
print("Generating Figure 1: Degree search curves …")

def degree_curve(X, y, degrees, alpha_fixed, label, ax, mark_deg):
    mses = []
    for d in degrees:
        poly = PolynomialFeatures(degree=d, include_bias=False)
        Xp   = poly.fit_transform(X)
        res  = cross_validate(Ridge(alpha=alpha_fixed), Xp, y, cv=KF,
                               scoring="neg_mean_squared_error")
        mses.append(-res["test_score"].mean())
    ax.plot(degrees, mses, "o-", color=PLT_COLOR, lw=2, ms=6)
    ax.axvline(mark_deg, color=PLT_COLOR2, ls="--", lw=1.5, label=f"Selected degree={mark_deg}")
    ax.set_xlabel("Polynomial Degree", fontsize=11)
    ax.set_ylabel("5-Fold CV MSE", fontsize=11)
    ax.set_title(label, fontsize=12, fontweight="bold")
    ax.legend(fontsize=9)
    ax.grid(True, color=GRID_COLOR)
    ax.set_xticks(degrees)
    return mses

fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
fig.suptitle("Degree Selection: Cross-Validation MSE vs Polynomial Degree", fontsize=13, fontweight="bold", y=1.01)

# Var1: use the fine-tuned alpha
degs1 = list(range(2,7))
degree_curve(X1, y1, degs1, 2.477076, "Var1 – Steam Turbine (6 features, α=2.477)", axes[0], 5)

# Var2: show degrees 4-14
degs2 = list(range(4,15))
degree_curve(X2, y2, degs2, 0.231013, "Var2 – Thermal Reservoir (3 features, α=0.231)", axes[1], 12)

plt.tight_layout()
FIG1 = savefig("fig1_degree_selection.png")

# ─────────────────────────────────────────────────────────────────
# FIGURE 2 – Alpha (regularisation) selection curves
# ─────────────────────────────────────────────────────────────────
print("Generating Figure 2: Alpha selection curves …")

def alpha_curve(X, y, deg, alphas, label, ax, mark_alpha):
    mses = []
    poly = PolynomialFeatures(degree=deg, include_bias=False)
    Xp   = poly.fit_transform(X)
    for a in alphas:
        res = cross_validate(Ridge(alpha=a), Xp, y, cv=KF,
                             scoring="neg_mean_squared_error")
        mses.append(-res["test_score"].mean())
    ax.semilogx(alphas, mses, "o-", color=PLT_COLOR, lw=2, ms=5)
    ax.axvline(mark_alpha, color=PLT_COLOR2, ls="--", lw=1.5, label=f"Selected α={mark_alpha:.3f}")
    ax.set_xlabel("Ridge Alpha (log scale)", fontsize=11)
    ax.set_ylabel("5-Fold CV MSE", fontsize=11)
    ax.set_title(label, fontsize=12, fontweight="bold")
    ax.legend(fontsize=9)
    ax.grid(True, color=GRID_COLOR)

fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
fig.suptitle("Regularisation Strength: CV MSE vs Ridge Alpha", fontsize=13, fontweight="bold", y=1.01)
alphas_plot = np.logspace(-3, 3, 30)
alpha_curve(X1, y1, 5,  alphas_plot, "Var1 – Degree 5",  axes[0], 2.477076)
alpha_curve(X2, y2, 12, alphas_plot, "Var2 – Degree 12", axes[1], 0.231013)
plt.tight_layout()
FIG2 = savefig("fig2_alpha_selection.png")

# ─────────────────────────────────────────────────────────────────
# FIGURE 3 – Actual vs Predicted (full-train fit, both problems)
# ─────────────────────────────────────────────────────────────────
print("Generating Figure 3: Actual vs Predicted …")

fig, axes = plt.subplots(1, 2, figsize=(12, 5))
fig.suptitle("Actual vs Predicted — Full Training Data", fontsize=13, fontweight="bold", y=1.01)

for ax, tag, X, y, deg, alpha, lbl in [
    (axes[0], "var1", X1, y1, 5,  2.477076, "Var1 – Steam Turbine"),
    (axes[1], "var2", X2, y2, 12, 0.231013, "Var2 – Thermal Reservoir"),
]:
    poly  = PolynomialFeatures(degree=deg, include_bias=False)
    Xp    = poly.fit_transform(X)
    model = Ridge(alpha=alpha); model.fit(Xp, y)
    yhat  = model.predict(Xp)
    mse   = mean_squared_error(y, yhat)
    r2    = r2_score(y, yhat)

    ax.scatter(y, yhat, alpha=0.35, s=14, color=PLT_COLOR, edgecolors="none")
    lims = [min(y.min(), yhat.min())-1, max(y.max(), yhat.max())+1]
    ax.plot(lims, lims, "r--", lw=1.5, label="Perfect fit")
    ax.set_xlabel("Actual y", fontsize=11); ax.set_ylabel("Predicted ŷ", fontsize=11)
    ax.set_title(f"{lbl}\nTrain MSE={mse:.4f}, R²={r2:.4f}", fontsize=11, fontweight="bold")
    ax.legend(fontsize=9); ax.grid(True, color=GRID_COLOR)

plt.tight_layout()
FIG3 = savefig("fig3_actual_vs_predicted.png")

# ─────────────────────────────────────────────────────────────────
# FIGURE 4 – Residual plots
# ─────────────────────────────────────────────────────────────────
print("Generating Figure 4: Residuals …")

fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))
fig.suptitle("Residuals — Full Training Data", fontsize=13, fontweight="bold", y=1.01)

for ax, X, y, deg, alpha, lbl in [
    (axes[0], X1, y1, 5,  2.477076, "Var1 – Steam Turbine"),
    (axes[1], X2, y2, 12, 0.231013, "Var2 – Thermal Reservoir"),
]:
    poly  = PolynomialFeatures(degree=deg, include_bias=False)
    Xp    = poly.fit_transform(X)
    model = Ridge(alpha=alpha); model.fit(Xp, y)
    yhat  = model.predict(Xp)
    resid = y - yhat

    ax.scatter(yhat, resid, alpha=0.35, s=14, color=PLT_COLOR, edgecolors="none")
    ax.axhline(0, color=PLT_COLOR2, ls="--", lw=1.5)
    ax.set_xlabel("Predicted ŷ", fontsize=11); ax.set_ylabel("Residual (y − ŷ)", fontsize=11)
    ax.set_title(f"{lbl}\nResid std={resid.std():.4f}", fontsize=11, fontweight="bold")
    ax.grid(True, color=GRID_COLOR)

plt.tight_layout()
FIG4 = savefig("fig4_residuals.png")

# ─────────────────────────────────────────────────────────────────
# FIGURE 5 – Per-fold CV bar chart
# ─────────────────────────────────────────────────────────────────
print("Generating Figure 5: Per-fold CV MSE …")

fold_data = {
    "Var1 (deg=5, α=2.477)": [0.4523, 0.4332, 0.4472, 0.6200, 0.5096],
    "Var2 (deg=12, α=0.231)": [0.2493, 0.2906, 0.2685, 0.2273, 0.2781],
}

fig, axes = plt.subplots(1, 2, figsize=(12, 4))
fig.suptitle("Per-Fold Validation MSE (5-Fold CV)", fontsize=13, fontweight="bold", y=1.01)
colors = [PLT_COLOR]*4 + [PLT_COLOR2]

for ax, (name, vals) in zip(axes, fold_data.items()):
    folds = [f"Fold {i+1}" for i in range(5)]
    bars = ax.bar(folds, vals, color=[PLT_COLOR,PLT_COLOR,PLT_COLOR,PLT_COLOR,PLT_COLOR],
                  edgecolor="white", width=0.6)
    mean_v = np.mean(vals)
    ax.axhline(mean_v, color=PLT_COLOR2, ls="--", lw=1.8, label=f"Mean={mean_v:.4f}")
    ax.set_ylabel("Validation MSE", fontsize=11)
    ax.set_title(name, fontsize=11, fontweight="bold")
    ax.legend(fontsize=9); ax.grid(True, axis="y", color=GRID_COLOR)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x()+bar.get_width()/2, v+0.003, f"{v:.4f}",
                ha="center", va="bottom", fontsize=8.5)

plt.tight_layout()
FIG5 = savefig("fig5_cv_folds.png")

# ─────────────────────────────────────────────────────────────────
# FIGURE 6 – Target distribution (train y histograms)
# ─────────────────────────────────────────────────────────────────
print("Generating Figure 6: Target distributions …")

fig, axes = plt.subplots(1, 2, figsize=(12, 4))
fig.suptitle("Target Variable Distribution (Training Data)", fontsize=13, fontweight="bold", y=1.01)

for ax, y, lbl, color in [
    (axes[0], y1, "Var1 – Steam Turbine (y)", PLT_COLOR),
    (axes[1], y2, "Var2 – Thermal Reservoir (y)", PLT_COLOR2),
]:
    ax.hist(y, bins=40, color=color, alpha=0.8, edgecolor="white")
    ax.axvline(y.mean(), color="black", ls="--", lw=1.5, label=f"Mean={y.mean():.2f}")
    ax.set_xlabel("y", fontsize=11); ax.set_ylabel("Count", fontsize=11)
    ax.set_title(f"{lbl}\nstd={y.std():.2f}, skew={pd.Series(y).skew():.3f}", fontsize=11, fontweight="bold")
    ax.legend(fontsize=9); ax.grid(True, color=GRID_COLOR)

plt.tight_layout()
FIG6 = savefig("fig6_target_distribution.png")

print("\nAll figures saved to:", PLOT_DIR)
print("Files:")
for f in sorted(os.listdir(PLOT_DIR)):
    print(f"  {f}")

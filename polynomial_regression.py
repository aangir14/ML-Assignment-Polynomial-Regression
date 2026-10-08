"""
Polynomial Regression Assignment — BT2024078
============================================
Phase 1 : Steam Turbine Power Score  (var1, 6 features, degree up to 10)
Phase 2 : Thermal Anomaly Score      (var2, 3 features, degree up to 20)

Strategy
--------
1. Load & sanity-check datasets.
2. Exhaustive degree x alpha grid search via 5-fold CV.
3. Validate best config with RidgeCV (efficient built-in alpha search).
4. Fine-grain alpha sweep around the winner.
5. Refit on full training data.
6. Generate prediction CSVs.
"""

import os
import warnings
import numpy as np
import pandas as pd

from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import Ridge, RidgeCV
from sklearn.pipeline import Pipeline
from sklearn.model_selection import KFold, cross_validate
from sklearn.metrics import mean_squared_error, r2_score

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────
DATA_DIR = r"c:\Users\Aangir Doshi\Downloads\BT2024078 (8)\BT2024078"
ROLLNO   = "BT2024078"

def path(split, var):
    return os.path.join(DATA_DIR, f"{ROLLNO}_{split}_{var}.csv")

def out_path(var):
    return os.path.join(DATA_DIR, f"{ROLLNO}_pred_{var}.csv")

# ─────────────────────────────────────────────
# 1. Load data
# ─────────────────────────────────────────────
print("=" * 60)
print("LOADING DATASETS")
print("=" * 60)

train1 = pd.read_csv(path("train", "var1"))
test1  = pd.read_csv(path("test",  "var1"))
train2 = pd.read_csv(path("train", "var2"))
test2  = pd.read_csv(path("test",  "var2"))

feat1 = ["x1", "x2", "x3", "x4", "x5", "x6"]
feat2 = ["x1", "x2", "x3"]

X_tr1, y_tr1 = train1[feat1].values, train1["y"].values
X_te1        = test1[feat1].values

X_tr2, y_tr2 = train2[feat2].values, train2["y"].values
X_te2        = test2[feat2].values

print(f"Var1 train: {X_tr1.shape}, test: {X_te1.shape}")
print(f"Var2 train: {X_tr2.shape}, test: {X_te2.shape}")
print(f"Var1 y: mean={y_tr1.mean():.4f}, std={y_tr1.std():.4f}, min={y_tr1.min():.4f}, max={y_tr1.max():.4f}")
print(f"Var2 y: mean={y_tr2.mean():.4f}, std={y_tr2.std():.4f}, min={y_tr2.min():.4f}, max={y_tr2.max():.4f}")
print()

# ─────────────────────────────────────────────
# 2. CV utility
# ─────────────────────────────────────────────
KF = KFold(n_splits=5, shuffle=True, random_state=42)

def cv_ridge(X, y, degree, alpha):
    """Return (cv_mse, cv_r2) for a given (degree, alpha) combo."""
    pipe = Pipeline([
        ("poly",  PolynomialFeatures(degree=degree, include_bias=False)),
        ("ridge", Ridge(alpha=alpha, fit_intercept=True)),
    ])
    res = cross_validate(pipe, X, y, cv=KF,
                         scoring={"mse": "neg_mean_squared_error", "r2": "r2"},
                         return_train_score=False)
    return -res["test_mse"].mean(), res["test_r2"].mean()


# ─────────────────────────────────────────────
# 3. Grid search helper
# ─────────────────────────────────────────────
def grid_search(X, y, degrees, alphas):
    print(f"{'Deg':>4}  {'Alpha':>10}  {'CV MSE':>10}  {'CV R2':>8}  {'#Feat':>6}")
    print("-" * 50)
    best = {"mse": np.inf, "r2": -np.inf, "deg": None, "alpha": None}
    rows = []
    dummy = PolynomialFeatures(include_bias=False)
    for deg in degrees:
        n_feat = PolynomialFeatures(degree=deg, include_bias=False).fit(X[:1]).n_output_features_
        for alpha in alphas:
            mse, r2 = cv_ridge(X, y, deg, alpha)
            marker = " <-- BEST" if mse < best["mse"] else ""
            print(f"{deg:>4}  {alpha:>10.2e}  {mse:>10.4f}  {r2:>8.4f}  {n_feat:>6}{marker}")
            rows.append({"deg": deg, "alpha": alpha, "cv_mse": mse, "cv_r2": r2, "n_feat": n_feat})
            if mse < best["mse"]:
                best.update({"mse": mse, "r2": r2, "deg": deg, "alpha": alpha})
    return best, pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════
# VAR 1 — Steam Turbine (6 features, degree up to 10)
# ═══════════════════════════════════════════════════════════════
print("=" * 60)
print("VAR1 — STEAM TURBINE POWER SCORE")
print("=" * 60)

print("\n[Var1] Coarse degree + alpha grid search ...")
degrees_v1     = [2, 3, 4, 5, 6]
alphas_coarse  = np.logspace(-3, 3, 13)
best_v1_coarse, df_v1_coarse = grid_search(X_tr1, y_tr1, degrees_v1, alphas_coarse)

print(f"\nCoarse best -> degree={best_v1_coarse['deg']}, alpha={best_v1_coarse['alpha']:.4f}, "
      f"CV MSE={best_v1_coarse['mse']:.4f}, CV R2={best_v1_coarse['r2']:.4f}")

# Fine alpha sweep around the best degree
best_deg_v1   = best_v1_coarse["deg"]
coarse_a_v1   = best_v1_coarse["alpha"]
print(f"\n[Var1] Fine alpha sweep for degree={best_deg_v1} ...")
fine_alphas_v1 = np.unique(np.concatenate([
    np.linspace(max(1e-4, coarse_a_v1 / 20), coarse_a_v1 * 20, 40),
    np.logspace(-2, 2, 30),
]).round(6))

best_v1_fine = {"mse": np.inf, "r2": -np.inf, "alpha": None}
print(f"{'Alpha':>12}  {'CV MSE':>10}  {'CV R2':>8}")
print("-" * 36)
for alpha in fine_alphas_v1:
    mse, r2 = cv_ridge(X_tr1, y_tr1, best_deg_v1, alpha)
    marker = " <--" if mse < best_v1_fine["mse"] else ""
    print(f"{alpha:>12.4f}  {mse:>10.4f}  {r2:>8.4f}{marker}")
    if mse < best_v1_fine["mse"]:
        best_v1_fine.update({"mse": mse, "r2": r2, "alpha": alpha})

# Check neighboring degrees at fine-tuned alpha
print(f"\n[Var1] Neighboring degrees with alpha={best_v1_fine['alpha']:.4f} ...")
neighbor_results_v1 = {}
for d in sorted(set([max(2, best_deg_v1 - 1), best_deg_v1, best_deg_v1 + 1])):
    n_feat = PolynomialFeatures(degree=d, include_bias=False).fit(X_tr1[:1]).n_output_features_
    if n_feat > 6000:
        print(f"  degree={d}: {n_feat} features -- skipping (too large)")
        continue
    mse, r2 = cv_ridge(X_tr1, y_tr1, d, best_v1_fine["alpha"])
    neighbor_results_v1[d] = (mse, r2)
    print(f"  degree={d} ({n_feat} feat): CV MSE={mse:.4f}, CV R2={r2:.4f}")

# Pick the best degree from neighbors
best_deg_v1_final = min(neighbor_results_v1, key=lambda d: neighbor_results_v1[d][0])
if neighbor_results_v1[best_deg_v1_final][0] < best_v1_fine["mse"]:
    best_deg_v1 = best_deg_v1_final
    print(f"  -> Updated best degree to {best_deg_v1}")

# RidgeCV sanity check
print(f"\n[Var1] RidgeCV confirmation for degree={best_deg_v1} ...")
poly_v1_tmp = PolynomialFeatures(degree=best_deg_v1, include_bias=False)
Xp_v1_tmp   = poly_v1_tmp.fit_transform(X_tr1)
alphas_rcv   = np.logspace(-3, 4, 100)
rcv1 = RidgeCV(alphas=alphas_rcv, cv=KF, scoring="neg_mean_squared_error")
rcv1.fit(Xp_v1_tmp, y_tr1)
print(f"  RidgeCV best alpha = {rcv1.alpha_:.6f}")
mse_rcv1, r2_rcv1 = cv_ridge(X_tr1, y_tr1, best_deg_v1, rcv1.alpha_)
print(f"  RidgeCV -> CV MSE={mse_rcv1:.4f}, CV R2={r2_rcv1:.4f}")

# Choose best between fine sweep and RidgeCV
if mse_rcv1 < best_v1_fine["mse"]:
    FINAL_ALPHA_V1 = rcv1.alpha_
    final_mse_v1   = mse_rcv1
    final_r2_v1    = r2_rcv1
else:
    FINAL_ALPHA_V1 = best_v1_fine["alpha"]
    final_mse_v1   = best_v1_fine["mse"]
    final_r2_v1    = best_v1_fine["r2"]
FINAL_DEG_V1 = best_deg_v1

print(f"\n{'='*60}")
print(f"VAR1 FINAL CONFIG: degree={FINAL_DEG_V1}, alpha={FINAL_ALPHA_V1:.6f}")
print(f"  Final CV MSE = {final_mse_v1:.4f}")
print(f"  Final CV R2  = {final_r2_v1:.4f}")
print(f"{'='*60}\n")


# ═══════════════════════════════════════════════════════════════
# VAR 2 — Thermal Reservoir (3 features, degree up to 20)
# ═══════════════════════════════════════════════════════════════
print("=" * 60)
print("VAR2 — THERMAL ANOMALY SCORE")
print("=" * 60)

print("\n[Var2] Coarse degree + alpha grid search ...")
degrees_v2       = list(range(4, 14))      # 4..13
alphas_coarse_v2 = np.logspace(-4, 2, 13) # 1e-4 .. 1e2

best_v2_coarse, df_v2_coarse = grid_search(X_tr2, y_tr2, degrees_v2, alphas_coarse_v2)
print(f"\nCoarse best -> degree={best_v2_coarse['deg']}, alpha={best_v2_coarse['alpha']:.4f}, "
      f"CV MSE={best_v2_coarse['mse']:.4f}, CV R2={best_v2_coarse['r2']:.4f}")

# Fine alpha sweep around best degree
best_deg_v2  = best_v2_coarse["deg"]
coarse_a_v2  = best_v2_coarse["alpha"]
print(f"\n[Var2] Fine alpha sweep for degree={best_deg_v2} ...")
fine_alphas_v2 = np.unique(np.concatenate([
    np.linspace(max(1e-5, coarse_a_v2 / 20), coarse_a_v2 * 20, 40),
    np.logspace(-4, 1, 30),
]).round(7))

best_v2_fine = {"mse": np.inf, "r2": -np.inf, "alpha": None}
print(f"{'Alpha':>12}  {'CV MSE':>10}  {'CV R2':>8}")
print("-" * 36)
for alpha in fine_alphas_v2:
    mse, r2 = cv_ridge(X_tr2, y_tr2, best_deg_v2, alpha)
    marker = " <--" if mse < best_v2_fine["mse"] else ""
    print(f"{alpha:>12.6f}  {mse:>10.4f}  {r2:>8.4f}{marker}")
    if mse < best_v2_fine["mse"]:
        best_v2_fine.update({"mse": mse, "r2": r2, "alpha": alpha})

# Check neighboring degrees
print(f"\n[Var2] Neighboring degrees with alpha={best_v2_fine['alpha']:.6f} ...")
neighbor_results_v2 = {}
for d in sorted(set([max(4, best_deg_v2 - 2), best_deg_v2 - 1,
                     best_deg_v2, best_deg_v2 + 1, best_deg_v2 + 2])):
    if d < 4:
        continue
    n_feat = PolynomialFeatures(degree=d, include_bias=False).fit(X_tr2[:1]).n_output_features_
    mse, r2 = cv_ridge(X_tr2, y_tr2, d, best_v2_fine["alpha"])
    neighbor_results_v2[d] = (mse, r2)
    print(f"  degree={d} ({n_feat} feat): CV MSE={mse:.4f}, CV R2={r2:.4f}")

best_deg_v2_final = min(neighbor_results_v2, key=lambda d: neighbor_results_v2[d][0])
if neighbor_results_v2[best_deg_v2_final][0] < best_v2_fine["mse"]:
    best_deg_v2 = best_deg_v2_final
    print(f"  -> Updated best degree to {best_deg_v2}")

# RidgeCV sanity check
print(f"\n[Var2] RidgeCV confirmation for degree={best_deg_v2} ...")
poly_v2_tmp = PolynomialFeatures(degree=best_deg_v2, include_bias=False)
Xp_v2_tmp   = poly_v2_tmp.fit_transform(X_tr2)
alphas_rcv2  = np.logspace(-5, 3, 100)
rcv2 = RidgeCV(alphas=alphas_rcv2, cv=KF, scoring="neg_mean_squared_error")
rcv2.fit(Xp_v2_tmp, y_tr2)
print(f"  RidgeCV best alpha = {rcv2.alpha_:.6f}")
mse_rcv2, r2_rcv2 = cv_ridge(X_tr2, y_tr2, best_deg_v2, rcv2.alpha_)
print(f"  RidgeCV -> CV MSE={mse_rcv2:.4f}, CV R2={r2_rcv2:.4f}")

if mse_rcv2 < best_v2_fine["mse"]:
    FINAL_ALPHA_V2 = rcv2.alpha_
    final_mse_v2   = mse_rcv2
    final_r2_v2    = r2_rcv2
else:
    FINAL_ALPHA_V2 = best_v2_fine["alpha"]
    final_mse_v2   = best_v2_fine["mse"]
    final_r2_v2    = best_v2_fine["r2"]
FINAL_DEG_V2 = best_deg_v2

print(f"\n{'='*60}")
print(f"VAR2 FINAL CONFIG: degree={FINAL_DEG_V2}, alpha={FINAL_ALPHA_V2:.6f}")
print(f"  Final CV MSE = {final_mse_v2:.4f}")
print(f"  Final CV R2  = {final_r2_v2:.4f}")
print(f"{'='*60}\n")


# ═══════════════════════════════════════════════════════════════
# 4. Per-fold breakdown (for reporting)
# ═══════════════════════════════════════════════════════════════
print("=" * 60)
print("PER-FOLD CV BREAKDOWN (final configs)")
print("=" * 60)

for label, X, y, deg, alpha in [
    ("Var1", X_tr1, y_tr1, FINAL_DEG_V1, FINAL_ALPHA_V1),
    ("Var2", X_tr2, y_tr2, FINAL_DEG_V2, FINAL_ALPHA_V2),
]:
    pipe = Pipeline([
        ("poly",  PolynomialFeatures(degree=deg, include_bias=False)),
        ("ridge", Ridge(alpha=alpha, fit_intercept=True)),
    ])
    res = cross_validate(pipe, X, y, cv=KF,
                         scoring={"mse": "neg_mean_squared_error", "r2": "r2"},
                         return_train_score=True)
    print(f"\n{label} (degree={deg}, alpha={alpha:.6f})")
    for i in range(5):
        print(f"  Fold {i+1}: train_MSE={-res['train_mse'][i]:.4f}  "
              f"val_MSE={-res['test_mse'][i]:.4f}  val_R2={res['test_r2'][i]:.4f}")
    mean_val_mse = -res["test_mse"].mean()
    mean_val_r2  =  res["test_r2"].mean()
    print(f"  MEAN:  val_MSE={mean_val_mse:.4f}  val_R2={mean_val_r2:.4f}")


# ═══════════════════════════════════════════════════════════════
# 5. Final fit on full training data
# ═══════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("FITTING FINAL MODELS ON FULL TRAINING DATA")
print("=" * 60)

# Var1
poly_v1  = PolynomialFeatures(degree=FINAL_DEG_V1, include_bias=False)
Xp_tr1   = poly_v1.fit_transform(X_tr1)
Xp_te1   = poly_v1.transform(X_te1)
model_v1 = Ridge(alpha=FINAL_ALPHA_V1, fit_intercept=True)
model_v1.fit(Xp_tr1, y_tr1)

tr_pred_v1 = model_v1.predict(Xp_tr1)
print(f"\nVar1 full-train MSE={mean_squared_error(y_tr1, tr_pred_v1):.4f}, "
      f"R2={r2_score(y_tr1, tr_pred_v1):.4f}")
print(f"  # poly features: {Xp_tr1.shape[1]}, intercept={model_v1.intercept_:.4f}, "
      f"|coef|max={np.abs(model_v1.coef_).max():.4f}")

pred_v1 = model_v1.predict(Xp_te1)
print(f"  Test predictions: min={pred_v1.min():.4f}, max={pred_v1.max():.4f}, "
      f"mean={pred_v1.mean():.4f}, std={pred_v1.std():.4f}")

# Var2
poly_v2  = PolynomialFeatures(degree=FINAL_DEG_V2, include_bias=False)
Xp_tr2   = poly_v2.fit_transform(X_tr2)
Xp_te2   = poly_v2.transform(X_te2)
model_v2 = Ridge(alpha=FINAL_ALPHA_V2, fit_intercept=True)
model_v2.fit(Xp_tr2, y_tr2)

tr_pred_v2 = model_v2.predict(Xp_tr2)
print(f"\nVar2 full-train MSE={mean_squared_error(y_tr2, tr_pred_v2):.4f}, "
      f"R2={r2_score(y_tr2, tr_pred_v2):.4f}")
print(f"  # poly features: {Xp_tr2.shape[1]}, intercept={model_v2.intercept_:.4f}, "
      f"|coef|max={np.abs(model_v2.coef_).max():.4f}")

pred_v2 = model_v2.predict(Xp_te2)
print(f"  Test predictions: min={pred_v2.min():.4f}, max={pred_v2.max():.4f}, "
      f"mean={pred_v2.mean():.4f}, std={pred_v2.std():.4f}")


# ═══════════════════════════════════════════════════════════════
# 6. Save prediction files
# ═══════════════════════════════════════════════════════════════
print("\n" + "=" * 60)
print("SAVING PREDICTION FILES")
print("=" * 60)

pd.DataFrame({"y": pred_v1}).to_csv(out_path("var1"), index=False)
pd.DataFrame({"y": pred_v2}).to_csv(out_path("var2"), index=False)

print(f"Saved: {out_path('var1')}")
print(f"Saved: {out_path('var2')}")

# Sanity check against sample submission shape
sample = pd.read_csv(os.path.join(DATA_DIR, "sample_submission.csv"))
out1   = pd.read_csv(out_path("var1"))
out2   = pd.read_csv(out_path("var2"))
assert out1.shape == sample.shape, f"Shape mismatch var1: {out1.shape} vs {sample.shape}"
assert out2.shape == sample.shape, f"Shape mismatch var2: {out2.shape} vs {sample.shape}"
assert list(out1.columns) == ["y"], "Column name mismatch var1"
assert list(out2.columns) == ["y"], "Column name mismatch var2"
print("Shape and column assertions passed.")

print("\n" + "=" * 60)
print("ALL DONE")
print(f"  Var1: degree={FINAL_DEG_V1}, alpha={FINAL_ALPHA_V1:.6f}, "
      f"CV MSE={final_mse_v1:.4f}, CV R2={final_r2_v1:.4f}")
print(f"  Var2: degree={FINAL_DEG_V2}, alpha={FINAL_ALPHA_V2:.6f}, "
      f"CV MSE={final_mse_v2:.4f}, CV R2={final_r2_v2:.4f}")
print("=" * 60)

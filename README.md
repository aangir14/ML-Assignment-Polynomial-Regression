# BT2024078 — Polynomial Regression Assignment

> **Machine Learning Assignment | Roll No: BT2024078**

This repository contains all code, predictions, and report for the Polynomial Regression assignment covering two geothermal energy prediction problems.

---

## 📋 Problem Overview

### Phase 1 — Steam Turbine Power Score (`var1`)
Predict the **Net Power Score** of a geothermal steam turbine from 6 operational parameters (x1–x6 representing percentage deviations in valve/pressure/flow settings).

### Phase 2 — Thermal Reservoir Anomaly Score (`var2`)
Predict the **Thermal Anomaly Score** at 3D spatial coordinates (x1=East-West, x2=North-South, x3=Vertical depth) to identify optimal geothermal drilling sites.

---

## 🗂️ Repository Structure

```
BT2024078/
├── polynomial_regression.py       # Main training + inference pipeline
├── generate_report_assets.py      # Generates all report figures (plots)
├── build_report.py                # Builds the DOCX report
├── requirements.txt               # Python dependencies
│
├── BT2024078_train_var1.csv       # Training data — Phase 1
├── BT2024078_test_var1.csv        # Test data   — Phase 1
├── BT2024078_train_var2.csv       # Training data — Phase 2
├── BT2024078_test_var2.csv        # Test data   — Phase 2
│
├── BT2024078_pred_var1.csv        # ✅ Predictions — Phase 1
├── BT2024078_pred_var2.csv        # ✅ Predictions — Phase 2
│
└── report_plots/                  # Auto-generated figures used in report
    ├── fig1_degree_selection.png
    ├── fig2_alpha_selection.png
    ├── fig3_actual_vs_predicted.png
    ├── fig4_residuals.png
    ├── fig5_cv_folds.png
    └── fig6_target_distribution.png
```

---

## 🔬 Approach & Results

| | **Var1 (Steam Turbine)** | **Var2 (Thermal Reservoir)** |
|---|---|---|
| Model | Polynomial + Ridge | Polynomial + Ridge |
| Polynomial Degree | **5** | **12** |
| Ridge Alpha (α) | **2.477** | **0.231** |
| # Polynomial Features | 461 | 454 |
| CV MSE (5-fold) | **0.4925** | **0.2627** |
| CV R² (5-fold) | **0.9516** | **0.9949** |
| Full-train R² | 0.9807 | 0.9967 |

### How degrees were selected
- **Exhaustive grid search**: Degrees × 13 log-spaced alpha values, evaluated with 5-fold CV
- **Fine alpha sweep**: ~50 dense alpha values around the coarse best
- **RidgeCV confirmation**: Independent sklearn RidgeCV with 100 alpha candidates
- **Neighbour degree check**: Adjacent degrees tested to confirm a genuine minimum

---

## ⚙️ How to Run

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the full training + prediction pipeline
```bash
python polynomial_regression.py
```
This will:
- Perform exhaustive hyperparameter tuning (degree + alpha) with 5-fold CV
- Print per-fold breakdown
- Fit final models on full training data
- Save `BT2024078_pred_var1.csv` and `BT2024078_pred_var2.csv`

### 3. Generate report figures
```bash
python generate_report_assets.py
```
Saves all plots to `report_plots/`.

### 4. Build the DOCX report
```bash
python build_report.py
```
Saves `BT2024078_Report.docx`.

---

## 📦 Dependencies

- Python ≥ 3.8
- scikit-learn
- numpy
- pandas
- matplotlib
- python-docx

---

## 📊 Key Findings

**Var1**: The data follows a **degree-5 polynomial** in 6 features. Without regularisation, degree 5 overfits severely. With Ridge α ≈ 2.5, it achieves CV R² = 0.9516, outperforming degree 4 by ~30% in MSE.

**Var2**: The thermal anomaly follows a **degree-12 polynomial** in 3D space. The 3-feature space grows slowly enough (454 features at degree 12) that a moderate Ridge penalty (α ≈ 0.23) achieves exceptional CV R² = 0.9949.

---

## 📁 Output Files

| File | Description |
|---|---|
| `BT2024078_pred_var1.csv` | 1000 predictions for Phase 1 test set |
| `BT2024078_pred_var2.csv` | 1000 predictions for Phase 2 test set |
| `BT2024078_Report.docx` | Full written report (4–5 pages) |

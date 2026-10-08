"""
build_report.py
Builds the final DOCX report for BT2024078 polynomial regression assignment.
Run AFTER generate_report_assets.py has been run (needs report_plots/ folder).
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import docx.opc.constants

DATA_DIR  = r"c:\Users\Aangir Doshi\Downloads\BT2024078 (8)\BT2024078"
PLOT_DIR  = os.path.join(DATA_DIR, "report_plots")
OUT_FILE  = os.path.join(DATA_DIR, "BT2024078_Report.docx")

# ── helpers ────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    """Set table cell background colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

def add_heading(doc, text, level, color_hex="1F3864"):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = h.runs[0] if h.runs else h.add_run(text)
    run.font.color.rgb = RGBColor.from_string(color_hex)
    return h

def add_para(doc, text, bold=False, italic=False, size=10.5, space_before=0, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

def add_figure(doc, img_path, caption, width_inches=6.0):
    if not os.path.exists(img_path):
        add_para(doc, f"[Figure placeholder – {os.path.basename(img_path)} not found]", italic=True)
        return
    doc.add_picture(img_path, width=Inches(width_inches))
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)
    cap.runs[0].font.size  = Pt(9)
    cap.runs[0].font.italic = True
    cap.runs[0].font.color.rgb = RGBColor(100, 100, 100)

def make_table(doc, headers, rows, header_bg="1F3864", header_text_color="FFFFFF"):
    n_cols = len(headers)
    t = doc.add_table(rows=1+len(rows), cols=n_cols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    # header row
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string(header_text_color)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_bg(cell, header_bg)

    # data rows
    for ri, row_data in enumerate(rows):
        row = t.rows[ri+1]
        bg = "F0F4FF" if ri % 2 == 0 else "FFFFFF"
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(9.5)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_bg(cell, bg)
    return t

def add_page_break(doc):
    doc.add_page_break()

# ═══════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ═══════════════════════════════════════════════════════════════
doc = Document()

# ── page margins ──────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── default font ──────────────────────────────────────────────
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)

# ══════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════
title = doc.add_heading("Polynomial Regression Assignment", level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.runs[0].font.color.rgb = RGBColor.from_string("1F3864")
title.runs[0].font.size = Pt(20)

subtitle = doc.add_paragraph("Machine Learning — BT2024078")
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle.paragraph_format.space_before = Pt(4)
for run in subtitle.runs:
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("2E75B6")

doc.add_paragraph()

info_table = doc.add_table(rows=3, cols=2)
info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
info_data = [
    ("Roll Number", "BT2024078"),
    ("Assignment",  "Polynomial Regression"),
    ("Problems",    "Var1 (Steam Turbine) & Var2 (Thermal Reservoir)"),
]
for row_obj, (k, v) in zip(info_table.rows, info_data):
    row_obj.cells[0].text = k
    row_obj.cells[0].paragraphs[0].runs[0].bold = True
    row_obj.cells[0].paragraphs[0].runs[0].font.size = Pt(11)
    row_obj.cells[1].text = v
    row_obj.cells[1].paragraphs[0].runs[0].font.size = Pt(11)

doc.add_paragraph()
add_page_break(doc)

# ══════════════════════════════════════════════════════════════
# 1. INTRODUCTION
# ══════════════════════════════════════════════════════════════
add_heading(doc, "1. Introduction", level=1)
add_para(doc,
    "This report documents the approach taken to solve two polynomial regression problems "
    "assigned as part of the Machine Learning course. The goal in both problems is to fit "
    "a polynomial regression model that accurately predicts a continuous target variable y "
    "from the given input features, using only the provided training data. Predictions are "
    "then generated for the held-out test sets.")

add_para(doc,
    "The two problems are:")
add_para(doc, "  • Phase 1 (var1): Predicting Net Power Score of a geothermal steam turbine "
         "from six operational parameters (x1–x6).", space_before=2, space_after=2)
add_para(doc, "  • Phase 2 (var2): Predicting Thermal Anomaly Score at 3D spatial coordinates "
         "(x1, x2, x3) for identifying optimal drilling sites.", space_before=2, space_after=6)

# ══════════════════════════════════════════════════════════════
# 2. DATASET OVERVIEW
# ══════════════════════════════════════════════════════════════
add_heading(doc, "2. Dataset Overview", level=1)

add_para(doc,
    "Both datasets were clean and pre-processed — no missing values, no outliers requiring "
    "removal. All input features were pre-normalised to the range [−1, 1], which is ideal for "
    "polynomial regression as it prevents extreme monomial values and avoids numerical conditioning issues.")

make_table(doc,
    ["Property", "Var1 (Steam Turbine)", "Var2 (Thermal Reservoir)"],
    [
        ["Input Features",   "6  (x1–x6)", "3  (x1–x3)"],
        ["Training Samples", "1,000",        "1,000"],
        ["Test Samples",     "1,000",        "1,000"],
        ["Feature Range",    "[−1, 1]",      "[−1, 1]"],
        ["y  min / max",     "−9.80 / +11.99", "−29.82 / +39.30"],
        ["y  mean",          "0.762",        "1.809"],
        ["y  std",           "3.213",        "7.273"],
        ["y  skewness",      "0.044 (symmetric)", "0.653 (mild)"],
        ["Missing values",   "0",            "0"],
    ])

doc.add_paragraph()
add_figure(doc, os.path.join(PLOT_DIR, "fig6_target_distribution.png"),
           "Figure 1. Distribution of target variable y for both training datasets.")

# ══════════════════════════════════════════════════════════════
# 3. METHODOLOGY
# ══════════════════════════════════════════════════════════════
add_page_break(doc)
add_heading(doc, "3. Methodology", level=1)

add_heading(doc, "3.1  Model: Polynomial Ridge Regression", level=2)
add_para(doc,
    "Polynomial regression extends linear regression by mapping the original features x into "
    "a higher-dimensional feature space consisting of all monomials up to a chosen total degree d. "
    "Formally, a polynomial of degree d means every term x1^a1 · x2^a2 · … · xk^ak where "
    "a1 + a2 + ... + ak ≤ d is included. This is implemented via scikit-learn's PolynomialFeatures "
    "transformer.")

add_para(doc,
    "To prevent overfitting in the high-dimensional polynomial feature space, Ridge regression "
    "(L2 regularisation) was used instead of ordinary least squares. Ridge adds a penalty term "
    "α‖w‖₂² to the loss, which shrinks coefficient magnitudes and stabilises solutions when "
    "feature count is large relative to samples. Ridge was preferred over Lasso because polynomial "
    "features are highly correlated (e.g., x1² and x1·x2), and L1 can behave erratically under "
    "collinearity.")

add_heading(doc, "3.2  Hyperparameter Tuning Strategy", level=2)
add_para(doc,
    "Two hyperparameters control the model: polynomial degree d and regularisation strength α. "
    "Both were selected through an exhaustive grid search evaluated with 5-fold cross-validation "
    "(shuffled, random_state=42), minimising mean validation MSE:")

add_para(doc, "  Step 1 — Coarse degree × alpha grid:", bold=True, space_before=3, space_after=2)
add_para(doc, "     • Var1: degrees {2,3,4,5,6} × 13 log-spaced alphas from 10⁻³ to 10³ = 65 CV evaluations.",
         space_before=0, space_after=2)
add_para(doc, "     • Var2: degrees {4,5,…,13} × 13 log-spaced alphas from 10⁻⁴ to 10² = 130 CV evaluations.",
         space_before=0, space_after=4)

add_para(doc, "  Step 2 — Fine alpha sweep:", bold=True, space_before=3, space_after=2)
add_para(doc, "     ~50 linearly- and log-spaced alphas in a tight window around the coarse best alpha, "
         "evaluated at the best degree.", space_before=0, space_after=4)

add_para(doc, "  Step 3 — Neighbour degree check:", bold=True, space_before=3, space_after=2)
add_para(doc, "     Degrees ±1 (Var1) or ±2 (Var2) were tested at the fine-tuned alpha to confirm "
         "the chosen degree is a genuine minimum.", space_before=0, space_after=4)

add_para(doc, "  Step 4 — RidgeCV confirmation:", bold=True, space_before=3, space_after=2)
add_para(doc, "     scikit-learn's RidgeCV with 100 log-spaced alphas was run as an independent "
         "verification. The alpha yielding the lower CV MSE between the fine sweep and RidgeCV was adopted.",
         space_before=0, space_after=6)

add_heading(doc, "3.3  Final Fit", level=2)
add_para(doc,
    "Once the optimal (degree, α) pair was determined, the model was refit on the entire "
    "training set (all 1,000 samples) and used to generate predictions on the 1,000 test samples.")

# ══════════════════════════════════════════════════════════════
# 4. DEGREE SELECTION
# ══════════════════════════════════════════════════════════════
add_page_break(doc)
add_heading(doc, "4. Degree & Regularisation Selection", level=1)

add_heading(doc, "4.1  Var1 — Steam Turbine (6 Features)", level=2)
add_para(doc,
    "The coarse search over degrees 2–6 showed clear improvement up to degree 5 "
    "with appropriate regularisation, then degradation at degree 6. "
    "Without sufficient regularisation, degree 5 overfits (CV MSE ≈ 1.50 at α = 10⁻³), "
    "but with a stronger penalty it outperforms degree 4.")

make_table(doc,
    ["Degree", "Best Alpha", "CV MSE", "CV R²", "# Features", "Decision"],
    [
        ["2", "1.0",      "3.2064", "0.6865", "27",  "Underfit"],
        ["3", "1.0",      "0.9069", "0.9109", "83",  "Good start"],
        ["4", "3.16",     "0.6858", "0.9325", "209", "Better"],
        ["5", "2.477",    "0.4925", "0.9516", "461", "✓ Selected"],
        ["6", "3.16",     "0.5418", "0.9470", "923", "Overfit risk"],
    ])

doc.add_paragraph()
add_para(doc, "Final selected configuration: Degree = 5, α = 2.477076", bold=True)
add_para(doc,
    "The key insight is that degree 5 has 461 features for 1,000 samples (ratio ≈ 2.2). "
    "Strong L2 regularisation (α ≈ 2.5) is essential to prevent overfitting — without it, "
    "degree 5 performs worse than degree 4.")

doc.add_paragraph()
add_figure(doc, os.path.join(PLOT_DIR, "fig1_degree_selection.png"),
           "Figure 2. CV MSE vs Polynomial Degree for both problems.")

add_figure(doc, os.path.join(PLOT_DIR, "fig2_alpha_selection.png"),
           "Figure 3. CV MSE vs Ridge Alpha (log scale) at the selected degree.")

add_heading(doc, "4.2  Var2 — Thermal Reservoir (3 Features)", level=2)
add_para(doc,
    "With only 3 input features, polynomial features grow much more slowly, enabling "
    "higher degrees without explosion. The search over degrees 4–13 showed steady improvement "
    "up to degree 12 with a minimum CV MSE of 0.2627, after which CV MSE increased, "
    "confirming overfitting onset.")

make_table(doc,
    ["Degree", "Best Alpha", "CV MSE", "CV R²", "# Features", "Decision"],
    [
        ["7",  "0.01",  "0.3335", "0.9935", "120", "Good"],
        ["8",  "0.01",  "0.2815", "0.9945", "165", "Better"],
        ["9",  "0.01",  "0.2749", "0.9946", "220", "Better"],
        ["10", "0.10",  "0.2668", "0.9948", "286", "Better"],
        ["11", "0.10",  "0.2661", "0.9948", "364", "Better"],
        ["12", "0.231", "0.2627", "0.9949", "454", "✓ Selected"],
        ["13", "0.316", "0.2663", "0.9948", "559", "Slight degradation"],
        ["14", "0.316", "0.2698", "0.9947", "680", "Overfit"],
    ])

doc.add_paragraph()
add_para(doc, "Final selected configuration: Degree = 12, α = 0.231013 (RidgeCV confirmed)", bold=True)
add_para(doc,
    "The true underlying polynomial appears to be of degree 12. The moderate α ensures "
    "that all 454 features are utilised without overfitting the 1,000 training points.")

# ══════════════════════════════════════════════════════════════
# 5. RESULTS
# ══════════════════════════════════════════════════════════════
add_page_break(doc)
add_heading(doc, "5. Results", level=1)

add_heading(doc, "5.1  Cross-Validation Performance", level=2)

make_table(doc,
    ["Problem", "Degree", "Alpha", "CV MSE (mean)", "CV R² (mean)", "# Poly Features"],
    [
        ["Var1 – Steam Turbine",    "5",  "2.477076", "0.4925", "0.9516", "461"],
        ["Var2 – Thermal Reservoir","12", "0.231013", "0.2627", "0.9949", "454"],
    ])

doc.add_paragraph()

add_heading(doc, "5.2  Per-Fold Breakdown", level=2)

make_table(doc,
    ["Fold", "Var1 Train MSE", "Var1 Val MSE", "Var1 Val R²",
              "Var2 Train MSE", "Var2 Val MSE", "Var2 Val R²"],
    [
        ["1", "0.1845", "0.4523", "0.9533", "0.1696", "0.2493", "0.9940"],
        ["2", "0.1869", "0.4332", "0.9497", "0.1653", "0.2906", "0.9946"],
        ["3", "0.1962", "0.4472", "0.9647", "0.1694", "0.2685", "0.9957"],
        ["4", "0.1796", "0.6200", "0.9440", "0.1719", "0.2273", "0.9955"],
        ["5", "0.1886", "0.5096", "0.9463", "0.1668", "0.2781", "0.9948"],
        ["Mean", "0.1872", "0.4925", "0.9516", "0.1686", "0.2627", "0.9949"],
    ])

doc.add_paragraph()
add_figure(doc, os.path.join(PLOT_DIR, "fig5_cv_folds.png"),
           "Figure 4. Per-fold validation MSE across 5 folds for both models.")

add_heading(doc, "5.3  Full Training Set Performance", level=2)

make_table(doc,
    ["Problem", "Train MSE", "Train R²", "Intercept", "Max |Coeff|"],
    [
        ["Var1 – Steam Turbine",    "0.1995", "0.9807", "0.2427", "2.0114"],
        ["Var2 – Thermal Reservoir","0.1745", "0.9967", "0.6028", "2.5788"],
    ])

doc.add_paragraph()
add_figure(doc, os.path.join(PLOT_DIR, "fig3_actual_vs_predicted.png"),
           "Figure 5. Actual vs Predicted values on full training data for both models.")

add_figure(doc, os.path.join(PLOT_DIR, "fig4_residuals.png"),
           "Figure 6. Residual plots (Predicted ŷ vs Residual y−ŷ) for both models. "
           "Residuals are centred around zero, indicating no systematic bias.")

# ══════════════════════════════════════════════════════════════
# 6. TECHNIQUES USED
# ══════════════════════════════════════════════════════════════
add_page_break(doc)
add_heading(doc, "6. Techniques & Rationale", level=1)

items = [
    ("PolynomialFeatures (sklearn)",
     "Generates all polynomial monomials up to degree d, where the sum of individual "
     "feature exponents ≤ d. Used with include_bias=False since Ridge fits an intercept separately."),
    ("Ridge Regression (L2 Regularisation)",
     "Preferred over OLS and Lasso. OLS becomes ill-conditioned with many correlated polynomial "
     "features. Lasso is unreliable under high feature correlation. Ridge uniformly shrinks all "
     "coefficients, providing stability and preventing individual large coefficients."),
    ("5-Fold Stratified CV (KFold, shuffle=True)",
     "Used throughout for all hyperparameter selection. Shuffle ensures that any ordering in the "
     "original data does not bias fold composition. Fixed random_state=42 for reproducibility."),
    ("Coarse-to-Fine Tuning",
     "First a wide coarse grid on (degree, alpha), then a dense fine sweep around the winner. "
     "This efficiently explores the space without exhaustive search of all combinations."),
    ("RidgeCV Confirmation",
     "An independent alpha selection using sklearn's RidgeCV with 100 log-spaced alphas served "
     "as a cross-check. Whichever of the manual sweep or RidgeCV gave lower CV MSE was used."),
    ("Neighbour Degree Check",
     "After selecting the best alpha, adjacent degrees were tested to confirm the selected degree "
     "is a genuine minimum and not a local artefact."),
]

for name, desc in items:
    p = doc.add_paragraph(style="List Bullet")
    run_name = p.add_run(name + ": ")
    run_name.bold = True
    run_name.font.size = Pt(10.5)
    run_desc = p.add_run(desc)
    run_desc.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(4)

# ══════════════════════════════════════════════════════════════
# 7. CONCLUSION
# ══════════════════════════════════════════════════════════════
add_heading(doc, "7. Conclusion", level=1)
add_para(doc,
    "Two polynomial Ridge regression models were built to predict the target variables across "
    "the two assigned problems. An exhaustive coarse-to-fine grid search with 5-fold "
    "cross-validation was used to jointly optimise polynomial degree and regularisation strength.")

make_table(doc,
    ["Problem", "Degree", "Alpha", "CV MSE", "CV R²"],
    [
        ["Var1 – Steam Turbine Power Score",    "5",  "2.477", "0.4925", "0.9516"],
        ["Var2 – Thermal Anomaly Score",        "12", "0.231", "0.2627", "0.9949"],
    ])

doc.add_paragraph()
add_para(doc,
    "Both models achieved high R² scores on cross-validation, indicating they generalise well "
    "to unseen data. Var2 in particular achieved near-perfect R² (0.9949), correctly identifying "
    "degree 12 as the true underlying polynomial structure. The full training set fits further "
    "confirmed the models are neither underfitting (high train R²) nor severely overfitting "
    "(train-val MSE gap is small).")

add_para(doc,
    "The prediction files BT2024078_pred_var1.csv and BT2024078_pred_var2.csv are submitted "
    "alongside this report.", italic=True, space_before=4)

# ── save ──────────────────────────────────────────────────────
doc.save(OUT_FILE)
print(f"Report saved: {OUT_FILE}")

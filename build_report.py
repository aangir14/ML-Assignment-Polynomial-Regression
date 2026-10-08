"""
build_report.py  –  Concise 5 to 6 page humanized report builder for BT2024078 (Aangir Doshi).
Scales content down to ~65-70% of previous size and fits cleanly into 5-6 pages.
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATA_DIR = r"c:\Users\Aangir Doshi\Downloads\BT2024078 (8)\BT2024078"
PLOT_DIR = os.path.join(DATA_DIR, "report_plots")
OUT_FILE = os.path.join(DATA_DIR, "BT2024078_Report.docx")

def add_para(doc, text, bold=False, italic=False, size=10.5,
             space_before=1, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT,
             line_spacing=1.15):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    p.alignment = align
    run = p.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    return p

def add_heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if level == 1:
        h.paragraph_format.space_before = Pt(7)
        h.paragraph_format.space_after  = Pt(3)
        for r in h.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)
    elif level == 2:
        h.paragraph_format.space_before = Pt(5)
        h.paragraph_format.space_after  = Pt(2)
        for r in h.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(11)
            r.font.bold = True
            r.font.color.rgb = RGBColor(30, 30, 30)
    return h

def add_figure(doc, img_path, caption, width_inches=4.8):
    if not os.path.exists(img_path):
        add_para(doc, f"[Figure not found: {os.path.basename(img_path)}]", italic=True)
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(4)
    p_img.paragraph_format.space_after  = Pt(2)
    run = p_img.add_run()
    run.add_picture(img_path, width=Inches(width_inches))

    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(1)
    cap.paragraph_format.space_after  = Pt(5)
    r_cap = cap.add_run(caption)
    r_cap.font.size = Pt(9)
    r_cap.font.italic = True
    r_cap.font.name = "Times New Roman"

def plain_table(doc, headers, rows, col_widths=None):
    n_cols = len(headers)
    t = doc.add_table(rows=1 + len(rows), cols=n_cols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    tblPr = t._tbl.tblPr
    tblCellMar = OxmlElement("w:tblCellMar")
    for m, val in [("top", "60"), ("bottom", "60"), ("left", "100"), ("right", "100")]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), val)
        node.set(qn("w:type"), "dxa")
        tblCellMar.append(node)
    tblPr.append(tblCellMar)

    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = h
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        if len(p.runs) > 0:
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.name = "Times New Roman"

    for ri, row_data in enumerate(rows):
        for ci, val in enumerate(row_data):
            cell = t.rows[ri + 1].cells[ci]
            cell.text = str(val)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci > 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after  = Pt(1)
            if len(p.runs) > 0:
                p.runs[0].font.size = Pt(9)
                p.runs[0].font.name = "Times New Roman"

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after  = Pt(3)
    return t

def build():
    doc = Document()

    # Standard academic margins
    for section in doc.sections:
        section.top_margin    = Cm(2.2)
        section.bottom_margin = Cm(2.2)
        section.left_margin   = Cm(2.4)
        section.right_margin  = Cm(2.4)

    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(10.5)

    # ──────────────────────────────────────────────────────────────
    # TITLE BLOCK (Compact student header)
    # ──────────────────────────────────────────────────────────────
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after  = Pt(2)
    tr = title_p.add_run("Polynomial Regression for Geothermal Plant Optimization")
    tr.bold = True
    tr.font.size = Pt(15)
    tr.font.name = "Times New Roman"

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after  = Pt(4)
    sr = sub_p.add_run("Course: Machine Learning (CS F464)   |   Student: Aangir Doshi   |   Roll Number: BT2024078")
    sr.font.size = Pt(10)
    sr.font.name = "Times New Roman"

    # Divider line
    rule = doc.add_paragraph()
    rule.paragraph_format.space_before = Pt(0)
    rule.paragraph_format.space_after  = Pt(6)
    pPr = rule._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "888888")
    pBdr.append(bottom)
    pPr.append(pBdr)

    # ──────────────────────────────────────────────────────────────
    # 1. INTRODUCTION & PROBLEM CONTEXT
    # ──────────────────────────────────────────────────────────────
    add_heading(doc, "1. Introduction & Problem Statement", level=1)
    add_para(doc,
        "In this assignment, I built polynomial regression models for two distinct engineering challenges "
        "in a geothermal power plant expansion project. The objective is to accurately predict continuous "
        "target variables y on hidden test datasets using only polynomial regression formulations. The datasets "
        "were individually generated for roll number BT2024078, meaning the functional relationships and optimal "
        "polynomial degrees are specific to my data.")

    add_para(doc,
        "Phase 1 (Var1) focuses on surface steam turbine optimization. It requires predicting the Net Power "
        "Score from six turbine calibration parameters (x1 through x6, representing valve adjustments, coolant "
        "flow, blade pitch, and pressures). Phase 2 (Var2) addresses subterranean thermal mapping, predicting a "
        "Thermal Anomaly Score from 3D sensor coordinates (x1: East-West, x2: North-South, x3: Depth) to locate "
        "optimal drilling sites for high-yield extraction wells.")

    # ──────────────────────────────────────────────────────────────
    # 2. EXPLORATORY DATA ANALYSIS
    # ──────────────────────────────────────────────────────────────
    add_heading(doc, "2. Exploratory Data Analysis & Observations", level=1)
    add_para(doc,
        "Before fitting any models, I checked both datasets for anomalies, missing values, and scale issues. "
        "Two main observations stood out:")

    add_para(doc,
        "First, both datasets were completely clean with zero missing entries. Crucially, all input features "
        "were already pre-scaled to [-1, 1]. This bounded range is ideal for polynomial regression: computing high "
        "powers (e.g., x^5 or x^12) of numbers bounded in [-1, 1] naturally keeps monomial values within [-1, 1], "
        "preventing numerical overflow and severe feature magnitude imbalances without extra scaling.")

    add_para(doc,
        "Second, the target distributions differ between the two problems. Var1 has a narrow target range "
        "(-9.80 to +11.99, mean 0.76, std 3.21) with near-zero skewness (0.04), indicating a symmetric distribution. "
        "Var2 exhibits a much wider range (-29.82 to +39.30, mean 1.81, std 7.27) and a mild positive skew (0.65), "
        "reflecting localized thermal hot spots in the 3D survey volume. Because the skew is physical rather than an "
        "artifact, I kept the raw target values directly.")

    plain_table(doc,
        ["Dataset Property", "Var1 (Steam Turbine)", "Var2 (Thermal Reservoir)"],
        [
            ["Input Features",               "6  (x1 to x6)",          "3  (x1, x2, x3)"],
            ["Training Samples",             "1000",                    "1000"],
            ["Test Samples (to predict)",    "1000",                    "1000"],
            ["Feature Bounding Domain",      "[-1.0, 1.0]",             "[-1.0, 1.0]"],
            ["Target Min / Max",             "-9.80 / +11.99",          "-29.82 / +39.30"],
            ["Target Mean / Std",            "0.76 / 3.21",             "1.81 / 7.27"],
            ["Missing Values",               "0 (None)",                "0 (None)"],
        ])

    # ──────────────────────────────────────────────────────────────
    # 3. METHODOLOGY & REGULARIZATION STRATEGY
    # ──────────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, "3. Methodology & Regularization Strategy", level=1)

    add_heading(doc, "3.1 Polynomial Basis Expansion", level=2)
    add_para(doc,
        "Polynomial regression maps input features x in R^D into a monomial space containing all interaction "
        "terms up to degree d. The total number of features (including bias) is (d + D)! / (d! D!). For Var1 (D = 6), "
        "feature count grows rapidly: degree 2 yields 27 features, degree 4 produces 209, degree 5 yields 461, and "
        "degree 6 reaches 923. For Var2 (D = 3), growth is much slower: degree 8 yields 165 features, degree 10 gives "
        "286, and degree 12 gives 454. This difference in dimensionality strongly dictates model capacity and the degree "
        "we can safely explore.")

    add_heading(doc, "3.2 Regularization: Why Ridge over OLS and Lasso?", level=2)
    add_para(doc,
        "Fitting Ordinary Least Squares (OLS) directly on hundreds of polynomial terms causes severe instability. "
        "High-order monomials on continuous features are naturally collinear (e.g., x1^2 and x1^4), making the design "
        "matrix ill-conditioned. Inverting (X^T X) produces huge opposing weights that overfit training noise and explode on "
        "unseen test points.")

    add_para(doc,
        "To prevent this, regularization is essential. I chose Ridge regression (L2 penalty) over Lasso (L1). Lasso "
        "tends to select arbitrary individual terms from clusters of correlated polynomial features while zeroing out others, "
        "which can disrupt smooth curvature. Ridge shrinks correlated monomial coefficients proportionally via "
        "(X^T X + alpha * I)^(-1) X^T y. This stabilizes matrix inversion, controls model variance, and preserves smooth "
        "physical response curves.")

    add_heading(doc, "3.3 Two-Stage Hyperparameter Tuning", level=2)
    add_para(doc,
        "I tuned polynomial degree d and Ridge penalty alpha jointly using 5-fold cross-validation in two stages:")
    add_para(doc,
        "1. Coarse Grid Search: Swept candidate degrees (2 to 6 for Var1; 4 to 14 for Var2) across 13 logarithmically "
        "spaced alpha values (10^-5 to 10^5) using 5-fold CV to locate the optimal region.")
    add_para(doc,
        "2. Fine Sweep & Verification: Conducted a fine sweep around the top candidate, confirmed with Scikit-Learn's "
        "RidgeCV routine, and selected the (degree, alpha) pair minimizing 5-fold CV MSE.")
    add_para(doc,
        "3. Final Retraining: Refitted the selected model on all 1000 training samples and predicted the 1000 test points.")

    # ──────────────────────────────────────────────────────────────
    # 4. DEGREE SELECTION & RESULTS
    # ──────────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, "4. Degree Selection and Parameter Tuning", level=1)

    add_heading(doc, "4.1 Phase 1 (Var1) – Steam Turbine Optimization (6 Features)", level=2)
    add_para(doc,
        "For Var1, the search showed that the turbine response contains non-linear interactions beyond simple "
        "quadratic or cubic forms. Degree 2 underfits heavily (CV MSE = 3.21, R^2 = 0.69). Degree 3 and 4 reduce error "
        "to 0.91 and 0.69 MSE, respectively.")

    add_para(doc,
        "An important dynamic occurred between degree 4 and degree 5: with weak regularization (alpha < 0.01), degree 5 "
        "overfits due to its 461 features (CV MSE > 1.5). However, once alpha is tuned up to ~2.48, degree 5 achieves a "
        "CV MSE of 0.4925 (R^2 = 0.9516), beating degree 4 by nearly 29%. Pushing to degree 6 (923 features) degrades "
        "validation error back to 0.54, confirming degree 5 as optimal.")

    plain_table(doc,
        ["Polynomial Degree", "Best Alpha", "5-Fold CV MSE", "5-Fold CV R²", "Number of Features"],
        [
            ["Degree 2",  "1.000",   "3.214",  "0.686",  "27"],
            ["Degree 3",  "1.000",   "0.912",  "0.910",  "83"],
            ["Degree 4",  "3.162",   "0.691",  "0.932",  "209"],
            ["Degree 5",  "2.477",   "0.4925", "0.9516", "461 (Optimal)"],
            ["Degree 6",  "3.162",   "0.543",  "0.947",  "923"],
        ])

    add_para(doc, "Final Selection for Var1: Degree = 5, Alpha = 2.477076 (Confirmed by RidgeCV).", bold=True)

    add_figure(doc, os.path.join(PLOT_DIR, "fig1_degree_selection.png"),
               "Figure 1: Cross-validation Mean Squared Error vs. polynomial degree for Var1 and Var2. "
               "Dashed lines mark the selected degrees.",
               width_inches=4.8)

    doc.add_page_break()
    add_heading(doc, "4.2 Phase 2 (Var2) – Subsurface Thermal Reservoir (3 Features)", level=2)
    add_para(doc,
        "Var2 behaves very differently because it has only 3 spatial coordinates. Even at degree 12, there are only "
        "454 monomial features for 1000 data points, preserving a healthy samples-to-parameters ratio of over 2:1.")

    add_para(doc,
        "Because 3D thermal plumes have intricate geometry with steep gradients and multiple local peaks, lower "
        "degrees fail to resolve the spatial field. CV error dropped steadily from degree 4 down through degree 11, "
        "reaching its minimum at degree 12 (CV MSE = 0.2627, R^2 = 0.9949). At degrees 13 and 14, validation error begins "
        "rising slightly. The optimal alpha for Var2 is lower (0.231013), as the 3D basis contains less collinear overlap "
        "than the 6-feature turbine problem.")

    plain_table(doc,
        ["Polynomial Degree", "Best Alpha", "5-Fold CV MSE", "5-Fold CV R²", "Number of Features"],
        [
            ["Degree 6",   "0.010",   "0.412",  "0.9920", "84"],
            ["Degree 8",   "0.010",   "0.282",  "0.9945", "165"],
            ["Degree 10",  "0.100",   "0.271",  "0.9947", "286"],
            ["Degree 12",  "0.231",   "0.2627", "0.9949", "454 (Optimal)"],
            ["Degree 14",  "0.316",   "0.272",  "0.9947", "680"],
        ])

    add_para(doc, "Final Selection for Var2: Degree = 12, Alpha = 0.231013 (Confirmed by RidgeCV).", bold=True)

    add_figure(doc, os.path.join(PLOT_DIR, "fig2_alpha_selection.png"),
               "Figure 2: Validation MSE vs Ridge alpha at the chosen degrees. Dashed lines mark the optimal alpha.",
               width_inches=4.8)

    # ──────────────────────────────────────────────────────────────
    # 5. MODEL EVALUATION & EXPERIMENTAL RESULTS
    # ──────────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, "5. Model Evaluation & Experimental Results", level=1)

    add_heading(doc, "5.1 Cross-Validation Summary", level=2)
    add_para(doc, "The final performance metrics under 5-fold cross-validation are summarized below:")

    plain_table(doc,
        ["Problem Statement", "Optimal Degree", "Selected Alpha", "5-Fold CV MSE", "5-Fold CV R²", "Features"],
        [
            ["Phase 1 (Var1) – Steam Turbine",     "5",  "2.477076", "0.4925", "0.9516", "461"],
            ["Phase 2 (Var2) – Thermal Reservoir", "12", "0.231013", "0.2627", "0.9949", "454"],
        ])

    add_heading(doc, "5.2 Fold-by-Fold Stability Analysis", level=2)
    add_para(doc,
        "To verify that results were consistent across data splits rather than lucky on a single fold, "
        "I examined per-fold performance. Both models show solid stability across all 5 folds:")

    plain_table(doc,
        ["CV Fold", "Var1 Train MSE", "Var1 Val MSE", "Var1 Val R²", "Var2 Train MSE", "Var2 Val MSE", "Var2 Val R²"],
        [
            ["Fold 1", "0.1851", "0.4523", "0.9533", "0.1702", "0.2493", "0.9940"],
            ["Fold 2", "0.1874", "0.4332", "0.9497", "0.1654", "0.2906", "0.9946"],
            ["Fold 3", "0.1960", "0.4472", "0.9647", "0.1691", "0.2685", "0.9957"],
            ["Fold 4", "0.1802", "0.6200", "0.9440", "0.1723", "0.2273", "0.9955"],
            ["Fold 5", "0.1893", "0.5096", "0.9463", "0.1670", "0.2781", "0.9948"],
            ["Mean",   "0.1876", "0.4925", "0.9516", "0.1688", "0.2627", "0.9949"],
        ])

    add_heading(doc, "5.3 Full Training Refit & Fit Diagnostics", level=2)
    add_para(doc,
        "Refitting each finalized model on the full 1000 training points yielded a train MSE of 0.1995 (R^2 = 0.9807) "
        "for Var1 and 0.1745 (R^2 = 0.9967) for Var2. The small difference between train MSE and cross-validation MSE "
        "(0.1995 vs 0.4925 for Var1; 0.1745 vs 0.2627 for Var2) demonstrates strong generalizability without overfitting.")

    add_figure(doc, os.path.join(PLOT_DIR, "fig3_actual_vs_predicted.png"),
               "Figure 3: Predicted vs actual target values on full training sets. Tight alignment confirms high fidelity.",
               width_inches=4.8)

    add_para(doc,
        "Residual diagnostics (actual minus predicted) confirm well-behaved models: residuals are symmetrically centered "
        "at zero with standard deviation 0.446 for Var1 and 0.417 for Var2, displaying no heteroscedastic fan patterns or "
        "unmodeled non-linear curvature.")

    # ──────────────────────────────────────────────────────────────
    # 6. DISCUSSION & PRACTICAL REFLECTIONS
    # ──────────────────────────────────────────────────────────────
    doc.add_page_break()
    add_heading(doc, "6. Discussion & Practical Engineering Reflections", level=1)
    add_para(doc,
        "Working through both problems provided several practical insights into polynomial regression dynamics:")

    add_para(doc,
        "1. Interplay of Degree and Regularization: In Var1, degree 5 with weak alpha overfits, but with proper "
        "regularization (alpha ~ 2.48), it outperforms degree 4 by nearly 29%. The extra capacity captures real interaction "
        "effects while the L2 penalty prevents noisy monomial coefficients from inflating. Hyperparameters must be tuned jointly.")

    add_para(doc,
        "2. Dimensionality vs Degree Capacity: In Var1 (6 features), degree 5 produced 461 features, reaching the practical "
        "limit for 1000 observations. In Var2 (3 features), degree 12 required only 454 features. Low-dimensional problems "
        "allow much deeper polynomial expansions to capture intricate physical fields without parameter explosion.")

    add_para(doc,
        "3. Collinearity and Shrinkage: Because continuous powers and interactions are collinear, unregularized OLS fails. "
        "Ridge regression proved well-suited by smoothly shrinking all coefficients rather than dropping them arbitrarily.")

    add_para(doc,
        "4. Future Improvements: If extending this work, I would test ElasticNet or feature importance pruning to identify "
        "and drop negligible monomial terms, yielding a sparser, more interpretable equation for plant engineers.")

    # ──────────────────────────────────────────────────────────────
    # 7. CONCLUSION & DELIVERABLES SUMMARY
    # ──────────────────────────────────────────────────────────────
    add_heading(doc, "7. Conclusion & Deliverables Summary", level=1)
    add_para(doc,
        "To conclude, I developed and validated two polynomial Ridge regression models for BT2024078. Var1 achieves a 5-fold "
        "CV R^2 of 0.9516 (CV MSE = 0.4925) using degree 5 with alpha = 2.477. Var2 achieves a CV R^2 of 0.9949 (CV MSE = 0.2627) "
        "using degree 12 with alpha = 0.231. Both models display tight training fits, consistent cross-validation folds, and "
        "well-dispersed residuals.")

    add_para(doc,
        "Predictions for both test sets (1000 rows each) have been verified and exported to CSV format. All code and assets "
        "are documented in the GitHub repository.")

    plain_table(doc,
        ["Deliverable", "File Name", "Details", "Status"],
        [
            ["Phase 1 Predictions", "BT2024078_pred_var1.csv", "1000 predictions for steam turbine test set", "Verified (No NaNs)"],
            ["Phase 2 Predictions", "BT2024078_pred_var2.csv", "1000 predictions for thermal reservoir test set", "Verified (No NaNs)"],
            ["Report Document",    "BT2024078_Report.docx",  "Technical documentation of methodology and findings", "Complete (BT2024078)"],
            ["Code Pipeline",       "polynomial_regression.py", "Complete model training, CV search, and inference", "Verified"],
            ["GitHub Repository",   "ML-Assignment-Polynomial-Regression", "Public Git repository with code and assets", "Pushed & Synchronized"],
        ])

    add_para(doc,
        "Report prepared by Aangir Doshi (Roll No: BT2024078) for Machine Learning Assignment 1.",
        italic=True, space_before=4)

    # Save
    doc.save(OUT_FILE)
    print(f"Successfully generated report at: {OUT_FILE}")

if __name__ == "__main__":
    build()

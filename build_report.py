"""
build_report.py  –  Compact version (~5-6 pages) for BT2024078.
Run AFTER generate_report_assets.py.
"""

import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

DATA_DIR = r"c:\Users\Aangir Doshi\Downloads\BT2024078 (8)\BT2024078"
PLOT_DIR = os.path.join(DATA_DIR, "report_plots")
OUT_FILE = os.path.join(DATA_DIR, "BT2024078_Report.docx")

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────

def add_para(doc, text, bold=False, italic=False, size=11,
             space_before=2, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.alignment = align
    run = p.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.font.size = Pt(size)
    return p

def add_heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h.paragraph_format.space_before = Pt(4)
    h.paragraph_format.space_after  = Pt(3)
    return h

def add_figure(doc, img_path, caption, width_inches=5.6):
    if not os.path.exists(img_path):
        add_para(doc, f"[Figure not found: {os.path.basename(img_path)}]", italic=True)
        return
    doc.add_picture(img_path, width=Inches(width_inches))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(1)
    cap.paragraph_format.space_after  = Pt(6)
    cap.runs[0].font.size   = Pt(9)
    cap.runs[0].font.italic = True

def plain_table(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style     = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].bold      = True
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
    for ri, row_data in enumerate(rows):
        for ci, val in enumerate(row_data):
            cell = t.rows[ri + 1].cells[ci]
            cell.text = str(val)
            cell.paragraphs[0].alignment           = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.size   = Pt(9.5)
    t.paragraph_format = None
    return t


# ──────────────────────────────────────────────────────────────
# BUILD DOCUMENT
# ──────────────────────────────────────────────────────────────
doc = Document()

for section in doc.sections:
    section.top_margin    = Cm(2.2)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.8)
    section.right_margin  = Cm(2.8)

doc.styles["Normal"].font.name = "Times New Roman"
doc.styles["Normal"].font.size = Pt(11)


# ══════════════════════════════════════════════════════════════
# TITLE BLOCK  (no full page break — shares page with intro)
# ══════════════════════════════════════════════════════════════

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after  = Pt(4)
tr = title_p.add_run("Polynomial Regression Assignment")
tr.bold = True; tr.font.size = Pt(16); tr.font.name = "Times New Roman"

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_p.paragraph_format.space_after = Pt(6)
sr = sub_p.add_run("Machine Learning  |  Aangir Doshi  |  BT2024078")
sr.font.size = Pt(11); sr.font.name = "Times New Roman"

# thin horizontal rule via bottom border on a blank para
rule = doc.add_paragraph()
rule.paragraph_format.space_after = Pt(6)
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
pPr = rule._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "6")
bottom.set(qn("w:space"), "1");    bottom.set(qn("w:color"), "888888")
pBdr.append(bottom); pPr.append(pBdr)


# ══════════════════════════════════════════════════════════════
# 1. INTRODUCTION  (short)
# ══════════════════════════════════════════════════════════════
add_heading(doc, "1. Introduction", level=1)

add_para(doc,
    "This report describes the polynomial regression models I built for the two assigned "
    "problems. Var1 asks for predictions of the Net Power Score of a geothermal steam "
    "turbine using six operational parameters (x1–x6). Var2 asks for predictions of a "
    "Thermal Anomaly Score at 3D spatial coordinates (x1, x2, x3) to guide drilling "
    "decisions. Both datasets were unique to my roll number.")

add_para(doc,
    "Both datasets were already clean — no missing values, and all features were "
    "pre-normalised to [-1, 1] — so no preprocessing was needed beyond building polynomial "
    "features. The table below gives a quick overview of what the data looks like.")

doc.add_paragraph().paragraph_format.space_after = Pt(2)
plain_table(doc,
    ["", "Var1 (Steam Turbine)", "Var2 (Thermal Reservoir)"],
    [
        ["Features",        "6  (x1–x6)",     "3  (x1, x2, x3)"],
        ["Train / Test",    "1000 / 1000",     "1000 / 1000"],
        ["y range",         "-9.80 to +11.99", "-29.82 to +39.30"],
        ["y std / skew",    "3.21 / 0.04",     "7.27 / 0.65"],
    ])
doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ══════════════════════════════════════════════════════════════
# 2. APPROACH
# ══════════════════════════════════════════════════════════════
add_heading(doc, "2. Approach", level=1)

add_para(doc,
    "I used polynomial Ridge regression throughout. The idea is to expand the original "
    "features into all monomials up to degree d (using sklearn's PolynomialFeatures), "
    "then fit a Ridge-regularised linear model on top. Ridge was chosen over plain OLS "
    "because at higher degrees the number of features grows large and OLS becomes "
    "unstable. Lasso was ruled out because polynomial features are inherently correlated "
    "and L1 handles that poorly.")

add_para(doc,
    "There are two things to tune: the degree d and the Ridge penalty alpha. I ran a "
    "coarse grid search (degrees × 13 log-spaced alpha values, evaluated with 5-fold CV) "
    "to find a promising region, then did a finer alpha sweep around the winner. I also "
    "checked neighbouring degrees at the best alpha, and ran sklearn's RidgeCV as an "
    "independent check. Whichever alpha gave the lower CV MSE was used. The model was "
    "then refitted on all 1000 training samples to generate the final test predictions.")


# ══════════════════════════════════════════════════════════════
# 3. DEGREE & REGULARISATION SELECTION
# ══════════════════════════════════════════════════════════════
add_heading(doc, "3. Degree and Regularisation Selection", level=1)

add_heading(doc, "3.1  Var1 – Steam Turbine (6 features)", level=2)

add_para(doc,
    "The coarse search over degrees 2–6 made the structure pretty clear. Degree 2 and 3 "
    "underfit noticeably. Things improve at degree 4 and peak at degree 5 — but only "
    "with a strong enough penalty. At low alpha, degree 5 overfits and actually does "
    "worse than degree 4. Once alpha is raised to around 2.5, degree 5 beats degree 4 "
    "by about 30% in CV MSE. Degree 6 gets worse again despite having more features.")

doc.add_paragraph().paragraph_format.space_after = Pt(2)
plain_table(doc,
    ["Degree", "Best alpha", "CV MSE", "CV R²", "No. features"],
    [
        ["2", "1.0",   "3.21", "0.69", "27"],
        ["3", "1.0",   "0.91", "0.91", "83"],
        ["4", "3.16",  "0.69", "0.93", "209"],
        ["5", "2.477", "0.49", "0.95", "461"],
        ["6", "3.16",  "0.54", "0.95", "923"],
    ])
doc.add_paragraph().paragraph_format.space_after = Pt(2)

add_para(doc, "Final choice: degree = 5, alpha = 2.477  (RidgeCV confirmed).", bold=True, space_after=6)

add_heading(doc, "3.2  Var2 – Thermal Reservoir (3 features)", level=2)

add_para(doc,
    "With only 3 input features, polynomial expansion is slower — degree 12 still gives "
    "just 454 features, well within what Ridge can handle. The search showed steady "
    "improvement from degree 4 all the way to degree 12, after which CV MSE starts "
    "creeping up again. The optimal alpha here is much smaller (0.23) because the "
    "feature space is less redundant.")

doc.add_paragraph().paragraph_format.space_after = Pt(2)
plain_table(doc,
    ["Degree", "Best alpha", "CV MSE", "CV R²", "No. features"],
    [
        ["8",  "0.01",  "0.28", "0.994", "165"],
        ["9",  "0.01",  "0.27", "0.995", "220"],
        ["10", "0.10",  "0.27", "0.995", "286"],
        ["11", "0.10",  "0.27", "0.995", "364"],
        ["12", "0.231", "0.26", "0.995", "454"],
        ["13", "0.316", "0.27", "0.995", "559"],
    ])
doc.add_paragraph().paragraph_format.space_after = Pt(2)

add_para(doc, "Final choice: degree = 12, alpha = 0.231  (RidgeCV confirmed).", bold=True, space_after=6)

add_figure(doc, os.path.join(PLOT_DIR, "fig1_degree_selection.png"),
           "Figure 1. CV MSE vs polynomial degree for both problems. "
           "Dashed lines mark the selected degrees.")


# ══════════════════════════════════════════════════════════════
# 4. RESULTS
# ══════════════════════════════════════════════════════════════
add_heading(doc, "4. Results", level=1)

add_para(doc, "Cross-validation summary for both final models:")

doc.add_paragraph().paragraph_format.space_after = Pt(2)
plain_table(doc,
    ["Problem", "Degree", "Alpha", "5-fold CV MSE", "5-fold CV R²", "Poly features"],
    [
        ["Var1 – Steam Turbine",     "5",  "2.477", "0.4925", "0.9516", "461"],
        ["Var2 – Thermal Reservoir", "12", "0.231", "0.2627", "0.9949", "454"],
    ])
doc.add_paragraph().paragraph_format.space_after = Pt(4)

add_para(doc,
    "After selecting hyperparameters via CV, I refitted each model on the full 1000 "
    "training samples. Train MSE for var1 was 0.1995 (R² = 0.98) and for var2 was "
    "0.1745 (R² = 0.9967). The small train-to-val gap in both cases suggests the "
    "models generalise well and are not severely overfitting.")

add_figure(doc, os.path.join(PLOT_DIR, "fig3_actual_vs_predicted.png"),
           "Figure 2. Actual vs predicted y on the full training set. "
           "Points close to the diagonal indicate accurate predictions.")

add_para(doc,
    "The per-fold results below confirm the CV scores are consistent across splits "
    "and not driven by one lucky fold:")

doc.add_paragraph().paragraph_format.space_after = Pt(2)
plain_table(doc,
    ["Fold", "Var1 val MSE", "Var1 val R²", "Var2 val MSE", "Var2 val R²"],
    [
        ["1",    "0.452", "0.953", "0.249", "0.994"],
        ["2",    "0.433", "0.950", "0.291", "0.995"],
        ["3",    "0.447", "0.965", "0.269", "0.996"],
        ["4",    "0.620", "0.944", "0.227", "0.996"],
        ["5",    "0.510", "0.946", "0.278", "0.995"],
        ["Mean", "0.493", "0.952", "0.263", "0.995"],
    ])
doc.add_paragraph().paragraph_format.space_after = Pt(4)


# ══════════════════════════════════════════════════════════════
# 5. CONCLUSION
# ══════════════════════════════════════════════════════════════
add_heading(doc, "5. Conclusion", level=1)

add_para(doc,
    "I built two polynomial Ridge regression models using an exhaustive coarse-to-fine "
    "grid search with 5-fold cross-validation to jointly tune degree and regularisation "
    "strength. The final configurations are:")

doc.add_paragraph().paragraph_format.space_after = Pt(2)
plain_table(doc,
    ["Problem", "Degree", "Alpha", "CV MSE", "CV R²"],
    [
        ["Var1 – Net Power Score",  "5",  "2.477", "0.493", "0.952"],
        ["Var2 – Thermal Anomaly",  "12", "0.231", "0.263", "0.995"],
    ])
doc.add_paragraph().paragraph_format.space_after = Pt(4)

add_para(doc,
    "Both models achieved strong cross-validation scores. Var2 in particular reached "
    "near-perfect R² of 0.995, suggesting the degree-12 polynomial captures the "
    "underlying thermal structure very well. The main takeaway from var1 was that "
    "degree and alpha have to be tuned together — degree 5 looked worse than degree 4 "
    "at low regularisation but clearly beat it once a strong enough penalty was applied.")

add_para(doc,
    "Prediction files BT2024078_pred_var1.csv and BT2024078_pred_var2.csv are submitted "
    "alongside this report.", italic=True, space_before=4)

# ── save ──────────────────────────────────────────────────────
doc.save(OUT_FILE)
print(f"Report saved: {OUT_FILE}")

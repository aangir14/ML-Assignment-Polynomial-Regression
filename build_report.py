"""
build_report.py  –  Rebuild the DOCX report for BT2024078.
Run AFTER generate_report_assets.py.
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

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────

def add_para(doc, text, bold=False, italic=False, size=11,
             space_before=2, space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.alignment = align
    run = p.add_run(text)
    run.bold        = bold
    run.italic      = italic
    run.font.size   = Pt(size)
    return p

def add_heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # strip any colour – just use Word's built-in heading style colours
    for run in h.runs:
        run.font.color.theme_color = None
        run.font.color.rgb = None
    return h

def add_figure(doc, img_path, caption, width_inches=5.8):
    if not os.path.exists(img_path):
        add_para(doc,
                 f"[Figure not found: {os.path.basename(img_path)}]",
                 italic=True)
        return
    doc.add_picture(img_path, width=Inches(width_inches))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    cap.runs[0].font.size   = Pt(9)
    cap.runs[0].font.italic = True

# Plain table – no coloured headers, no alternating row shading.
# Just a simple "Table Grid" with bold header text, like a normal student table.
def plain_table(doc, headers, rows):
    n_cols = len(headers)
    t = doc.add_table(rows=1 + len(rows), cols=n_cols)
    t.style     = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    # header row – just bold, nothing fancy
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold      = True
        p.runs[0].font.size = Pt(10)

    # data rows – normal text, centred
    for ri, row_data in enumerate(rows):
        for ci, val in enumerate(row_data):
            cell = t.rows[ri + 1].cells[ci]
            cell.text = str(val)
            p = cell.paragraphs[0]
            p.alignment           = WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.size   = Pt(10)

    return t


# ──────────────────────────────────────────────────────────────
# BUILD DOCUMENT
# ──────────────────────────────────────────────────────────────
doc = Document()

# margins
for section in doc.sections:
    section.top_margin    = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin   = Cm(2.8)
    section.right_margin  = Cm(2.8)

# default Normal style
style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(11)


# ══════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════

doc.add_paragraph()
doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run("Polynomial Regression Assignment")
r.bold      = True
r.font.size = Pt(18)
r.font.name = "Times New Roman"

doc.add_paragraph()

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = subtitle.add_run("Machine Learning (CS F464)")
r2.font.size = Pt(13)
r2.font.name = "Times New Roman"

doc.add_paragraph()
doc.add_paragraph()

for label, value in [
    ("Name",        "Aangir Doshi"),
    ("Roll Number", "BT2024078"),
    ("Assignment",  "Polynomial Regression – Phase 1 & Phase 2"),
]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lbl_run = p.add_run(f"{label}:  ")
    lbl_run.bold      = True
    lbl_run.font.size = Pt(12)
    val_run = p.add_run(value)
    val_run.font.size = Pt(12)
    p.paragraph_format.space_after = Pt(4)

doc.add_page_break()


# ══════════════════════════════════════════════════════════════
# 1. INTRODUCTION
# ══════════════════════════════════════════════════════════════
add_heading(doc, "1. Introduction", level=1)

add_para(doc,
    "This report covers my work on the two polynomial regression problems assigned for this "
    "course. For each problem I had to build a model that predicts a continuous output y "
    "from the given input features, using only the training data provided. The goal was to "
    "pick the right polynomial degree and keep overfitting under control so the model "
    "generalises well to the test set.")

add_para(doc,
    "The two problems are set in the context of a geothermal power plant project. The first "
    "one (var1) asks you to predict the Net Power Score of a steam turbine based on six "
    "operational parameters. The second one (var2) involves predicting a Thermal Anomaly "
    "Score at different 3D coordinates, which is used to find the best spots for drilling. "
    "Both datasets were individually generated per roll number, so the polynomial structure "
    "is unique to my data.")


# ══════════════════════════════════════════════════════════════
# 2. DATA
# ══════════════════════════════════════════════════════════════
add_heading(doc, "2. Looking at the Data", level=1)

add_para(doc,
    "Before doing anything else I spent some time just looking at both datasets to "
    "understand what I was working with. A few things stood out immediately:")

add_para(doc,
    "Both datasets were already clean — no missing values at all, and the input features "
    "were already normalised to lie in the range [-1, 1]. That was convenient because it "
    "meant I didn't have to worry about scaling before building polynomial features. "
    "Polynomial regression can go badly wrong if features are on very different scales, "
    "so having them pre-normalised saved a step.")

add_para(doc,
    "The target variable y in var1 ranges from about -9.8 to +12 with a standard deviation "
    "of 3.2 and almost no skew (0.04), so the distribution is fairly symmetric. Var2 has a "
    "much wider range (-29.8 to +39.3, std = 7.3) and a mild positive skew of 0.65 — "
    "nothing dramatic enough to warrant any transformations, but worth noting.")

doc.add_paragraph()
plain_table(doc,
    ["", "Var1 (Steam Turbine)", "Var2 (Thermal Reservoir)"],
    [
        ["Features",         "6  (x1 to x6)",  "3  (x1, x2, x3)"],
        ["Training rows",    "1000",            "1000"],
        ["Test rows",        "1000",            "1000"],
        ["Feature range",    "[-1, 1]",         "[-1, 1]"],
        ["y range",          "-9.80 to +11.99", "-29.82 to +39.30"],
        ["y mean",           "0.76",            "1.81"],
        ["y std",            "3.21",            "7.27"],
        ["y skew",           "0.04",            "0.65"],
        ["Missing values",   "0",               "0"],
    ])

doc.add_paragraph()
add_figure(doc, os.path.join(PLOT_DIR, "fig6_target_distribution.png"),
           "Figure 1. Distribution of y in both training sets.")


# ══════════════════════════════════════════════════════════════
# 3. APPROACH
# ══════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, "3. Approach", level=1)

add_heading(doc, "3.1 Why polynomial regression with Ridge?", level=2)

add_para(doc,
    "The assignment says to use polynomial regression, so the core idea is to take the "
    "original features and expand them into all possible products and powers up to some "
    "degree d — for example with two features x1 and x2 and degree 2 you'd get x1, x2, "
    "x1^2, x1*x2, and x2^2. Then you fit a linear model on top of these expanded features. "
    "I used sklearn's PolynomialFeatures for this.")

add_para(doc,
    "The tricky part is that as the degree goes up, the number of features explodes fast. "
    "At degree 5 with 6 input features you already have 461 polynomial features for 1000 "
    "training samples. Plain linear regression (ordinary least squares) becomes unstable "
    "in this kind of situation — it tends to find extreme coefficients and overfit badly. "
    "So I used Ridge regression instead, which adds an L2 penalty on the coefficient "
    "magnitudes. This keeps things from going haywire and usually generalises much better.")

add_para(doc,
    "I considered using Lasso as well, but decided against it. Polynomial features are "
    "inherently correlated with each other (e.g., x1^2 and x1*x1 are literally the same), "
    "and Lasso tends to behave unpredictably when features are strongly correlated. Ridge "
    "handles this more gracefully.")

add_heading(doc, "3.2 How I picked the degree and regularisation strength", level=2)

add_para(doc,
    "There are two things to tune: the polynomial degree d, and the Ridge penalty "
    "strength alpha. I treated this as a two-dimensional search problem and used "
    "5-fold cross-validation to evaluate each combination. The idea was to start broad "
    "and then zoom in.")

add_para(doc,
    "First I ran a coarse grid search — for var1 I tried degrees 2 through 6, and for "
    "var2 degrees 4 through 13, each combined with 13 log-spaced alpha values spread "
    "from very small to very large. This gave me a rough sense of where the best "
    "region was. Then I did a much finer sweep of alpha values around the winner from "
    "the coarse search, and also checked whether neighbouring degrees (±1 for var1, "
    "±2 for var2) did better at that alpha. Finally I ran sklearn's RidgeCV as an "
    "independent check — it has its own efficient built-in cross-validation for alpha "
    "selection — and I used whichever alpha gave the lower validation MSE.")

add_para(doc,
    "Once I had the final (degree, alpha) pair, I retrained the model on the full 1000 "
    "training samples and generated predictions for the test set.")


# ══════════════════════════════════════════════════════════════
# 4. CHOOSING THE DEGREE
# ══════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, "4. Degree Selection and Results", level=1)

add_heading(doc, "4.1 Var1 – Steam Turbine", level=2)

add_para(doc,
    "For var1 the coarse search made it pretty clear that the data has a degree-5 "
    "structure. At lower degrees the model noticeably underfits — degree 2 barely "
    "gets an R2 of 0.69 and degree 3 sits around 0.91. Things improve steadily through "
    "degree 4 and 5. Degree 6 actually starts getting worse again despite having more "
    "features, which tells you the model is starting to overfit.")

add_para(doc,
    "One thing that surprised me a bit: at a very low alpha, degree 5 actually performs "
    "worse than degree 4. But once you crank the regularisation up to around alpha = 2.5, "
    "degree 5 clearly wins. The 461 polynomial features need a fairly strong penalty to "
    "stay in check. The fine alpha sweep confirmed that the sweet spot is around 2.47.")

doc.add_paragraph()
plain_table(doc,
    ["Degree", "Best alpha tried", "CV MSE", "CV R2", "No. of features"],
    [
        ["2", "1.0",    "3.21",   "0.69", "27"],
        ["3", "1.0",    "0.91",   "0.91", "83"],
        ["4", "3.16",   "0.69",   "0.93", "209"],
        ["5", "2.477",  "0.49",   "0.95", "461"],
        ["6", "3.16",   "0.54",   "0.95", "923"],
    ])

doc.add_paragraph()
add_para(doc,
    "Final choice for var1: degree 5, alpha = 2.477.")

doc.add_paragraph()
add_figure(doc, os.path.join(PLOT_DIR, "fig1_degree_selection.png"),
           "Figure 2. How CV MSE changes with polynomial degree for both problems. "
           "The dashed line marks the selected degree.")

add_figure(doc, os.path.join(PLOT_DIR, "fig2_alpha_selection.png"),
           "Figure 3. CV MSE as a function of Ridge alpha at the chosen degree. "
           "The dashed line marks the selected alpha.")

add_heading(doc, "4.2 Var2 – Thermal Reservoir", level=2)

add_para(doc,
    "Var2 is different because there are only 3 input features, so polynomial features "
    "grow much slower. At degree 12 you still only have 454 features, which is very "
    "manageable. This meant I could go to much higher degrees without things blowing up.")

add_para(doc,
    "The search showed a steady and clear improvement all the way from degree 4 up to "
    "degree 12. After that the validation MSE starts creeping back up, so degree 12 "
    "is where things peak. The R2 at degree 12 is 0.9949 which is really quite good — "
    "the model is capturing almost all the variance in the data. I also noticed that "
    "for this problem, the optimal alpha is much smaller (around 0.23) because the "
    "features are fewer and less redundant at each degree.")

doc.add_paragraph()
plain_table(doc,
    ["Degree", "Best alpha tried", "CV MSE", "CV R2", "No. of features"],
    [
        ["7",  "0.01",  "0.33", "0.994", "120"],
        ["8",  "0.01",  "0.28", "0.994", "165"],
        ["9",  "0.01",  "0.27", "0.995", "220"],
        ["10", "0.10",  "0.27", "0.995", "286"],
        ["11", "0.10",  "0.27", "0.995", "364"],
        ["12", "0.231", "0.26", "0.995", "454"],
        ["13", "0.316", "0.27", "0.995", "559"],
        ["14", "0.316", "0.27", "0.995", "680"],
    ])

doc.add_paragraph()
add_para(doc,
    "Final choice for var2: degree 12, alpha = 0.231 (confirmed by RidgeCV).")


# ══════════════════════════════════════════════════════════════
# 5. RESULTS
# ══════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, "5. Results", level=1)

add_heading(doc, "5.1 Cross-validation summary", level=2)

add_para(doc,
    "Here is a quick summary of the final cross-validation numbers for both models:")

doc.add_paragraph()
plain_table(doc,
    ["Problem", "Degree", "Alpha", "5-fold CV MSE", "5-fold CV R2", "Poly features"],
    [
        ["Var1 (Steam Turbine)",    "5",  "2.477", "0.4925", "0.9516", "461"],
        ["Var2 (Thermal Reservoir)","12", "0.231", "0.2627", "0.9949", "454"],
    ])

doc.add_paragraph()

add_heading(doc, "5.2 Per-fold breakdown", level=2)

add_para(doc,
    "To make sure the results were consistent and not just lucky on one particular "
    "split, here are the numbers across each of the 5 folds:")

doc.add_paragraph()
plain_table(doc,
    ["Fold", "Var1 train MSE", "Var1 val MSE", "Var1 val R2",
              "Var2 train MSE", "Var2 val MSE", "Var2 val R2"],
    [
        ["1",    "0.185", "0.452", "0.953", "0.170", "0.249", "0.994"],
        ["2",    "0.187", "0.433", "0.950", "0.165", "0.291", "0.995"],
        ["3",    "0.196", "0.447", "0.965", "0.169", "0.269", "0.996"],
        ["4",    "0.180", "0.620", "0.944", "0.172", "0.227", "0.996"],
        ["5",    "0.189", "0.510", "0.946", "0.167", "0.278", "0.995"],
        ["Mean", "0.187", "0.493", "0.952", "0.169", "0.263", "0.995"],
    ])

doc.add_paragraph()

add_para(doc,
    "For var1, fold 4 gave a slightly higher validation MSE (0.62) compared to the "
    "others, but overall the results are pretty consistent across folds which suggests "
    "the model isn't just getting lucky on a particular data split. For var2 the "
    "results are very stable across all folds which is reassuring.")

add_figure(doc, os.path.join(PLOT_DIR, "fig5_cv_folds.png"),
           "Figure 4. Validation MSE for each fold. The dashed line is the mean.")

add_heading(doc, "5.3 Fit on the full training data", level=2)

add_para(doc,
    "After picking the hyperparameters through cross-validation, I refitted each model "
    "on the complete training set of 1000 samples and generated predictions for the test set.")

doc.add_paragraph()
plain_table(doc,
    ["Problem", "Train MSE", "Train R2"],
    [
        ["Var1 (Steam Turbine)",    "0.1995", "0.9807"],
        ["Var2 (Thermal Reservoir)","0.1745", "0.9967"],
    ])

doc.add_paragraph()

add_figure(doc, os.path.join(PLOT_DIR, "fig3_actual_vs_predicted.png"),
           "Figure 5. Actual vs predicted y on the full training set for both models. "
           "Points close to the diagonal line mean accurate predictions.")

add_figure(doc, os.path.join(PLOT_DIR, "fig4_residuals.png"),
           "Figure 6. Residuals (actual minus predicted) plotted against predicted values. "
           "The spread looks roughly even with no obvious pattern, which is a good sign.")

add_para(doc,
    "The train R2 for var1 is 0.98 and the validation R2 was 0.95, so there is a small "
    "gap but nothing alarming — it just means the model fits training data a little better "
    "than unseen data, which is completely normal. For var2 the gap is even smaller "
    "(train R2 = 0.997, val R2 = 0.995), suggesting the model has found the underlying "
    "structure very well.")


# ══════════════════════════════════════════════════════════════
# 6. DISCUSSION
# ══════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading(doc, "6. Discussion", level=1)

add_para(doc,
    "A few things are worth reflecting on from this assignment.")

add_para(doc,
    "For var1, I initially expected degree 4 to be the answer based on some quick manual "
    "checks I ran at the start. But the full grid search showed that degree 5 with proper "
    "regularisation actually does noticeably better. The key was that at low regularisation "
    "degree 5 overfits heavily (CV MSE jumps to 1.5), but once you increase alpha to "
    "around 2.5, it outperforms degree 4 by about 30% in terms of MSE. This really drove "
    "home the importance of tuning both hyperparameters together rather than fixing one "
    "and only searching over the other.")

add_para(doc,
    "For var2, the high degree (12) was a bit unexpected but makes sense in context. "
    "The thermal anomaly score is trying to capture a complex 3D heat map which can "
    "have all sorts of peaks and valleys — a high-degree polynomial in three spatial "
    "coordinates is actually a reasonable way to model that. And because we only have "
    "3 input features, even degree 12 only gives 454 features, which is well within "
    "what Ridge can handle with 1000 training points.")

add_para(doc,
    "One thing I would do differently if I had more time is look into whether some of "
    "the 461 (var1) or 454 (var2) polynomial features are actually contributing "
    "meaningfully to the model, or whether a lot of them are near-zero after "
    "regularisation. That could be interesting to explore with something like Lasso "
    "or by looking at which monomials have the largest coefficients after the Ridge fit.")


# ══════════════════════════════════════════════════════════════
# 7. CONCLUSION
# ══════════════════════════════════════════════════════════════
add_heading(doc, "7. Conclusion", level=1)

add_para(doc,
    "To summarise, I built two polynomial Ridge regression models for the two assigned "
    "problems. The approach was to do a proper grid search over both degree and "
    "regularisation strength using 5-fold cross-validation, rather than just picking "
    "values by intuition.")

doc.add_paragraph()
plain_table(doc,
    ["Problem", "Degree", "Alpha", "CV MSE", "CV R2"],
    [
        ["Var1 – Net Power Score",    "5",  "2.477", "0.493", "0.952"],
        ["Var2 – Thermal Anomaly",    "12", "0.231", "0.263", "0.995"],
    ])

doc.add_paragraph()

add_para(doc,
    "Both models ended up with strong cross-validation scores, particularly var2 which "
    "hit an R2 of 0.995. The residual plots look well-behaved and the train-val gap is "
    "small for both, so I'm fairly confident the models will generalise well to the "
    "hidden test data.")

add_para(doc,
    "Prediction files (BT2024078_pred_var1.csv and BT2024078_pred_var2.csv) are "
    "submitted alongside this report.",
    italic=True, space_before=6)

# ── save ──────────────────────────────────────────────────────
doc.save(OUT_FILE)
print(f"Report saved: {OUT_FILE}")

"""
Generate a detailed Sprint Retrospective Excel file for the MILimbEEG BCI Project.
Output: d:/BCI/documents/Sprint_Retrospective.xlsx
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
import os

OUTPUT_PATH = r"D:\BCI\documents\Sprint_Retrospective.xlsx"

# ─── Colour Palette ───────────────────────────────────────────────
C_DARK_BLUE   = "1F2D54"   # header backgrounds
C_MID_BLUE    = "2E4A8A"   # sub-headers
C_LIGHT_BLUE  = "D6E4F7"   # alternate row tint
C_GREEN       = "217346"   # positive / went-well
C_GREEN_LIGHT = "C6EFCE"
C_RED         = "9C0006"   # negative / didn't go well
C_RED_LIGHT   = "FFC7CE"
C_AMBER       = "9C6500"
C_AMBER_LIGHT = "FFEB9C"
C_WHITE       = "FFFFFF"
C_GREY_LIGHT  = "F2F2F2"

def side():
    return Side(style="thin", color="BFBFBF")

def thick_side():
    return Side(style="medium", color=C_DARK_BLUE)

def border():
    return Border(left=side(), right=side(), top=side(), bottom=side())

def thick_border():
    return Border(left=thick_side(), right=thick_side(),
                  top=thick_side(), bottom=thick_side())

def hdr_font(size=11, bold=True, color=C_WHITE):
    return Font(name="Calibri", size=size, bold=bold, color=color)

def cell_font(size=10, bold=False, color="000000"):
    return Font(name="Calibri", size=size, bold=bold, color=color)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def wrap_align(h="left", v="top"):
    return Alignment(horizontal=h, vertical=v, wrap_text=True)

def center_align():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def write_cell(ws, row, col, value, font=None, fill_=None,
               alignment=None, border_=None, number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:        c.font       = font
    if fill_:       c.fill       = fill_
    if alignment:   c.alignment  = alignment
    if border_:     c.border     = border_
    if number_format: c.number_format = number_format
    return c

def merge_write(ws, r1, c1, r2, c2, value, font=None, fill_=None,
                alignment=None, border_=None):
    ws.merge_cells(start_row=r1, start_column=c1,
                   end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    if font:      c.font      = font
    if fill_:     c.fill      = fill_
    if alignment: c.alignment = alignment
    if border_:   c.border    = border_
    return c

# ══════════════════════════════════════════════════════════════════
# SHEET 1 — PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════════
def build_overview(wb):
    ws = wb.create_sheet("📋 Project Overview")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 60
    ws.row_dimensions[1].height = 40

    # Title
    merge_write(ws, 1, 1, 1, 2,
                "MILimbEEG — BCI Sprint Retrospective (Dec 2025 – Apr 2026)",
                font=hdr_font(16, True, C_WHITE),
                fill_=fill(C_DARK_BLUE),
                alignment=center_align())

    fields = [
        ("Project Name",       "MILimbEEG — EEG Motor Imagery BCI System"),
        ("Objective",          "Classify limb motor imagery (8 classes) from 16-channel EEG using FIR filtering, CSP/FBCSP spatial filtering, and SVM/LDA classification."),
        ("Dataset",            "MILimbEEG (proprietary, pre-downloaded) — subjects S1–S25+, tasks: BEO, CLH, CRH, DLF, PLF, DRF, PRF, Rest"),
        ("Sampling Rate",      "125 Hz"),
        ("EEG Channels",       "16 (FC5, F3, Fz, F4, FC6, FC1, FC2, Cz, T7, CP5, C3, CP1, CP2, C4, CP6, T8)"),
        ("Frequency Bands",    "Alpha 8-12 Hz | Beta 12-30 Hz | Gamma 30-50 Hz"),
        ("Total Sprints",      "5 Sprints across ~19 weeks"),
        ("Best Result",        "86% accuracy — Statistical Features + SVM (Movement vs. Rest)"),
        ("Hardest Task",       "Foot MI (DLF vs. DRF) — ~47.5% with CSP+SVM (hardware bottleneck)"),
        ("Tech Stack",         "Python, NumPy, Pandas, SciPy, MNE-Python, scikit-learn, Matplotlib, openpyxl"),
        ("Code Directory",     "D:\\BCI\\codes\\"),
        ("Data Directory",     "D:\\BCI\\MILimbEEG\\"),
        ("Documents",          "D:\\BCI\\documents\\"),
    ]

    for i, (label, value) in enumerate(fields, start=3):
        even = (i % 2 == 0)
        bg = C_LIGHT_BLUE if even else C_WHITE
        write_cell(ws, i, 1, label,
                   font=cell_font(10, True, C_DARK_BLUE),
                   fill_=fill(bg), alignment=wrap_align(), border_=border())
        write_cell(ws, i, 2, value,
                   font=cell_font(10), fill_=fill(bg),
                   alignment=wrap_align(), border_=border())
        ws.row_dimensions[i].height = 30

# ══════════════════════════════════════════════════════════════════
# SHEET 2 — SPRINT TIMELINE
# ══════════════════════════════════════════════════════════════════
def build_timeline(wb):
    ws = wb.create_sheet("🗓️ Sprint Timeline")
    ws.sheet_view.showGridLines = False

    cols = ["Sprint", "Period", "Duration", "Theme",
            "Goal Summary", "Primary Scripts Delivered"]
    widths = [12, 26, 14, 26, 48, 52]
    for i, (col, w) in enumerate(zip(cols, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Header row
    ws.row_dimensions[1].height = 36
    for i, col in enumerate(cols, 1):
        write_cell(ws, 1, i, col,
                   font=hdr_font(11, True, C_WHITE),
                   fill_=fill(C_DARK_BLUE),
                   alignment=center_align(), border_=border())

    sprints = [
        ("Sprint 1", "Dec 1 – Dec 21, 2025", "3 weeks",
         "Dataset Acquisition & Metadata",
         "Validate dataset structure, parse file naming convention, map all 8 movement labels, generate metadata.xlsx with full trial inventory.",
         "download_prep.py, metadata_update.py, verify_metadata.py"),

        ("Sprint 2", "Dec 22, 2025 – Jan 18, 2026", "4 weeks",
         "FIR Filtering & Signal Validation",
         "Design 101-tap linear-phase FIR bandpass filters for Alpha/Beta/Gamma bands, apply zero-phase filtfilt across all subjects/channels, validate with synthetic and real EEG PSD plots.",
         "fir_main.py, validate_filters.py, visualize_psd_comparison.py"),

        ("Sprint 3", "Jan 19 – Feb 9, 2026", "3 weeks",
         "Feature Extraction & Baseline SVM",
         "Extract Power, Variance, RMS, Spectral Entropy, Mu & Beta band power per channel-band. Train baseline SVM with GridSearchCV for hyperparameter tuning. Evaluate binary (Movement vs. Rest).",
         "feature_extraction.py, train_svm.py, prepare_binary_data.py, train_binary_svm.py"),

        ("Sprint 4", "Feb 10 – Mar 2, 2026", "3 weeks",
         "CSP & FBCSP Spatial Filtering",
         "Implement MNE-based CSP pipeline for DLF vs. DRF foot MI. Extend to FBCSP with 7 frequency sub-bands, Ledoit-Wolf regularisation, mutual information feature selection, compare SVM vs. LDA. Optimize channel selection for motor cortex.",
         "prepare_csp_data.py, train_csp.py, train_csp_sliding.py, train_fbcsp.py"),

        ("Sprint 5", "Mar 3 – Apr 6, 2026", "5 weeks",
         "Visualization, Optimization & Documentation",
         "Generate publication-quality accuracy comparison charts, CSP scalp topomaps (standard 10-20 montage), confusion matrices. Compute spectral entropy. Write technical documentation and project explanation materials.",
         "visualize_eeg_results.py, topology_heatmap.py, heatmap2.py, entropy_main.py, entropy_log.py, plot_test.py, technical_methodology.md"),
    ]

    sprint_colors = [
        ("3B5998", "D0D9F0"),  # S1 — blue
        ("217346", "C6EFCE"),  # S2 — green
        ("9C6500", "FFEB9C"),  # S3 — amber
        ("7B3F8C", "E8D5F0"),  # S4 — purple
        ("C0392B", "FFC7CE"),  # S5 — red
    ]

    for r, (sprint, sc) in enumerate(zip(sprints, sprint_colors), start=2):
        ws.row_dimensions[r + 1].height = 60
        dark, light = sc
        for c, val in enumerate(sprint, 1):
            bg = light if c > 1 else dark
            fnt = hdr_font(10, True, C_WHITE) if c == 1 else cell_font(10)
            write_cell(ws, r + 1, c, val,
                       font=fnt, fill_=fill(bg),
                       alignment=wrap_align("center" if c <= 3 else "left"),
                       border_=border())

# ══════════════════════════════════════════════════════════════════
# SHEET 3 — DETAILED SPRINT CARDS
# ══════════════════════════════════════════════════════════════════
def section_header(ws, row, title, color, ncols=5):
    merge_write(ws, row, 1, row, ncols, title,
                font=hdr_font(11, True, C_WHITE),
                fill_=fill(color), alignment=center_align())
    ws.row_dimensions[row].height = 28

def sub_header(ws, row, cols_vals, fill_color, ncols=None):
    for c, v in enumerate(cols_vals, 1):
        write_cell(ws, row, c, v,
                   font=hdr_font(10, True, C_WHITE),
                   fill_=fill(fill_color),
                   alignment=center_align(), border_=border())
    ws.row_dimensions[row].height = 22

def data_row(ws, row, vals, bg=C_WHITE, col_start=1):
    for c, v in enumerate(vals, col_start):
        write_cell(ws, row, c, v,
                   font=cell_font(10), fill_=fill(bg),
                   alignment=wrap_align(), border_=border())

sprint_data = [
    {
        "name": "Sprint 1 — Data Preparation & FIR Filtering",
        "period": "Dec 1, 2025 – Jan 15, 2026",
        "color": "1F4E79",
        "story_points": 47,
        "velocity": 47,
        "stories": [
            ("US-01", "Verify dataset folder structure (S1–S25+)", "Done", "High", 3, "Agrim Pandey"),
            ("US-02", "Design regex parser for file name convention (S{n}R{r}{M|I}{L}_{t})", "Done", "High", 5, "Agrim Pandey"),
            ("US-03", "Map 8 movement labels to string and integer encoding", "Done", "High", 2, "Agrim Pandey"),
            ("US-04", "Generate metadata.xlsx with srno, patient, local_url, rep, task, label", "Done", "High", 5, "Agrim Pandey"),
            ("US-05", "Sort subjects numerically (S1, S2 … not S1, S10, S11)", "Done", "Medium", 2, "Agrim Pandey"),
            ("US-06", "Add graceful error handling for malformed filenames", "Done", "Medium", 2, "Agrim Pandey"),
            ("US-07", "Verify metadata row count matches expected trial count", "Done", "Medium", 2, "Agrim Pandey"),
        ],
        "went_well": [
            "File naming convention was perfectly consistent — regex parsed all files cleanly",
            "Metadata schema (8 columns) proved reusable across every downstream script without modification",
            "Natural sort key (patient_sort_key) fixed lexicographic ordering issue (S1 < S10 < S2)",
            "os.makedirs(exist_ok=True) pattern standardised across all output directories",
        ],
        "didnt_go_well": [
            "Download code was initially written but then removed since data was pre-downloaded — left dead/duplicate Step headers in the file",
            "No schema validation on metadata — downstream type errors (str vs. int in patient_number) had to be traced back manually",
            "No unit tests written at this stage — bugs in label_map required manual inspection",
        ],
        "learnings": [
            "Define and validate metadata schema (types, constraints) up front; it is the contract every script depends on",
            "Remove placeholder / dead code before moving to the next sprint — technical debt compounds quickly",
            "A simple test asserting row count == expected trial count would have caught data gaps immediately",
        ],
        "blockers": ["Dataset download URL required manual confirmation", "Excel file path spacing issues on Windows"],
        "risks": ["Metadata correctness is the single point of failure for all downstream work"],
    },
    {
        "name": "Sprint 2 — FIR Filtering & Signal Validation",
        "period": "Dec 22, 2025 – Jan 18, 2026",
        "color": "217346",
        "story_points": 26,
        "velocity": 26,
        "stories": [
            ("US-08", "Design 101-tap linear-phase FIR bandpass filter (firwin)", "Done", "High", 5, "Agrim Pandey"),
            ("US-09", "Apply zero-phase filter via filtfilt across all channels", "Done", "High", 5, "Agrim Pandey"),
            ("US-10", "Process all subject folders and save per-band output CSVs to fir_dataset/", "Done", "High", 5, "Agrim Pandey"),
            ("US-11", "Implement skip-if-exists to avoid redundant reprocessing", "Done", "Medium", 2, "Agrim Pandey"),
            ("US-12", "Synthetic validation test (10 + 20 + 40 Hz composite signal)", "Done", "High", 3, "Agrim Pandey"),
            ("US-13", "Real EEG PSD validation plots (Welch method)", "Done", "High", 3, "Agrim Pandey"),
            ("US-14", "PSD comparison overlay plots saved to validation_plots/", "Done", "Medium", 3, "Agrim Pandey"),
        ],
        "went_well": [
            "101-tap (odd) filter order enforces linear-phase response — mathematically correct for brain signal analysis",
            "filtfilt (zero-phase filtering) eliminates phase distortion that would shift EEG feature timing",
            "Synthetic test cleanly confirmed each band isolated expected component frequencies",
            "Skip-if-exists logic saved hours of recomputation on the large dataset",
            "Column naming convention ({ch}_{band}) was intuitive and self-documenting",
        ],
        "didnt_go_well": [
            "Gamma band (30–50 Hz) sits close to Nyquist limit (62.5 Hz) — roll-off is slightly compressed",
            "Delta (0.5–4 Hz) and Theta (4–8 Hz) bands were not filtered — limits future low-frequency ERP analysis",
            "Output CSV column order (index, then ch_a, ch_b, ch_g interleaved) was assumed, never documented formally",
            "Memory usage was not monitored — large FIR dataset fills disk quickly for 25+ subjects",
        ],
        "learnings": [
            "Always pair DSP code with an explicit synthetic signal test — it is the only way to prove correctness before touching real data",
            "Document the output column schema in a README or header comment so downstream scripts don't need to guess",
            "FIR order is a design parameter: lower = faster but poorer attenuation; always check stopband attenuation in dB",
        ],
        "blockers": ["None significant"],
        "risks": ["If FS is ever changed, all filtered data must be regenerated — hardcoded FS=125 in multiple files"],
    },
    {
        "name": "Sprint 3 — Feature Extraction & Baseline SVM",
        "period": "Jan 19 – Feb 9, 2026",
        "color": "9C6500",
        "story_points": 29,
        "velocity": 29,
        "stories": [
            ("US-15", "Implement Power, Variance, RMS feature functions", "Done", "High", 3, "Agrim Pandey"),
            ("US-16", "Add Spectral Entropy via Welch PSD (base-2 bits)", "Done", "High", 4, "Agrim Pandey"),
            ("US-17", "Add Mu-band (8-13 Hz) and Beta-band (13-30 Hz) power features", "Done", "High", 3, "Agrim Pandey"),
            ("US-18", "Loop over all metadata rows and build X feature matrix (N × 288 features)", "Done", "High", 5, "Agrim Pandey"),
            ("US-19", "Normalize features with StandardScaler (fit on train only)", "Done", "High", 3, "Akshat Mittal"),
            ("US-20", "Save X.npy, y.npy, scaler.joblib, feature_names.json", "Done", "High", 2, "Akshat Mittal"),
            ("US-21", "Train SVM with GridSearchCV (C, gamma, kernel) and 5-fold CV", "Done", "High", 5, "Akshat Mittal"),
            ("US-22", "Binary SVM for movement vs. rest classification (86% accuracy)", "Done", "High", 4, "Akshat Mittal"),
        ],
        "went_well": [
            "Modular feature functions (calculate_power, calculate_rms, etc.) are clean, unit-testable, and reusable",
            "StandardScaler fit only on training data — correctly avoids data leakage into test set",
            "GridSearchCV with n_jobs=-1 parallelised hyperparameter tuning efficiently",
            "86% binary accuracy (Movement vs. Rest) is a strong and meaningful baseline result",
            "feature_names.json provides full traceability from feature index to its origin channel and band",
        ],
        "didnt_go_well": [
            "Feature matrix is very wide (288 features for 16-ch × 3-band × 6-feature) — curse of dimensionality risk for SVMs",
            "feature_names.json was saved but never used for feature importance ranking or selection",
            "Multi-class SVM performance (all 8 labels) was never benchmarked — only binary pairs were evaluated",
            "Welch PSD uses nperseg=min(len(x), 256) — with 125 Hz data this gives coarse frequency resolution",
        ],
        "learnings": [
            "Wide feature vectors for SVMs benefit from PCA or explicit selection — investigate before assuming the model will generalise",
            "Always benchmark multi-class before narrowing to binary — the confusion matrix reveals which classes overlap",
            "feature_names.json is a good practice; wire it into the model card / report automatically",
        ],
        "blockers": ["Grid search over large param space takes ~20-30 min without GPU"],
        "risks": ["288 features × thousands of samples makes training slow; an SVM with RBF kernel scales as O(n²–n³)"],
    },
    {
        "name": "Sprint 4 — CSP & FBCSP Spatial Filtering",
        "period": "Feb 10 – Mar 2, 2026",
        "color": "7B3F8C",
        "story_points": 34,
        "velocity": 30,
        "stories": [
            ("US-23", "Extract raw time-series trials (channels × time) for DLF vs. DRF", "Done", "High", 5, "Agrim Pandey"),
            ("US-24", "Handle variable trial lengths with truncation/zero-padding", "Done", "High", 3, "Agrim Pandey"),
            ("US-25", "Apply 8-30 Hz MNE bandpass filter before CSP", "Done", "High", 2, "Akshat Mittal"),
            ("US-26", "Train MNE CSP with Ledoit-Wolf regularisation (4 components)", "Done", "High", 5, "Akshat Mittal"),
            ("US-27", "GridSearchCV over n_components (4/6/8) + SVM (C, gamma)", "Done", "High", 5, "Akshat Mittal"),
            ("US-28", "Implement FBCSP with 7 sub-bands (4–40 Hz)", "Done", "High", 5, "Akshat Mittal"),
            ("US-29", "Mutual Information feature selection (top-10 from 28 FBCSP features)", "Done", "High", 3, "Akshat Mittal"),
            ("US-30", "Compare SVM vs. LDA on FBCSP features", "Done", "Medium", 2, "Akshat Mittal"),
            ("US-31", "Sliding window CSP variant for temporal analysis", "Partial", "Low", 3, "Akshat Mittal"),
            ("US-32", "Optimise channel selection for motor cortex relevance", "Done", "Medium", 1, "Agrim Pandey"),
        ],
        "went_well": [
            "Ledoit-Wolf regularisation robustly handles ill-conditioned covariance matrices common with 16-channel EEG",
            "FBCSP correctly split train/test BEFORE CSP fitting — avoided the most common leakage bug in BCI literature",
            "Mutual Information feature selection reduced FBCSP features from 28 → 10 with principled ranking",
            "LDA provided a fast, interpretable, and competitive alternative to SVM for spatial filter features",
            "Pipeline object (CSP → SVC) from scikit-learn made cross-validation clean and leak-free",
        ],
        "didnt_go_well": [
            "CSP accuracy for foot MI (DLF vs. DRF) plateaued at ~47.5% — near chance level — due to 16-channel spatial limitation",
            "FBCSP showed similar result (~47.1%) — adding more bands did not compensate for channel count",
            "train_csp_sliding.py was partially implemented but never benchmarked or integrated into visualizations",
            "CSP patterns (filters & patterns matrices) were not logged per run — hard to compare across experiments",
            "Foot MI is the hardest class pair in this dataset — should have started with hand MI for proof-of-concept",
        ],
        "learnings": [
            "For foot motor imagery, 64+ channels are needed — 16 channels cannot resolve the fine somatotopic differences",
            "Start with the easiest class pair (hand MI: CLH vs. CRH) to validate the pipeline before attempting harder pairs",
            "Save all CSP filter matrices and accuracy metrics to JSON after each run — allows systematic comparison",
            "FBCSP is architecturally correct; the bottleneck is hardware (electrode count), not the algorithm",
        ],
        "blockers": ["Near-chance accuracy created uncertainty about whether the code or the data was the problem"],
        "risks": ["If spatial filtering cannot discriminate foot MI with 16ch, the entire CSP branch may need rethinking"],
    },
    {
        "name": "Sprint 5 — Visualization, Optimization & Documentation",
        "period": "Mar 3 – Apr 6, 2026",
        "color": "C0392B",
        "story_points": 30,
        "velocity": 27,
        "stories": [
            ("US-33", "Accuracy comparison bar chart across all methods and tasks", "Done", "High", 4, "Agrim Pandey"),
            ("US-34", "CSP spatial filter topomaps using MNE + standard 10-20 montage", "Done", "High", 5, "Akshat Mittal"),
            ("US-35", "Confusion matrices for CSP and FBCSP (DLF vs. DRF)", "Done", "High", 3, "Akshat Mittal"),
            ("US-36", "Electrode topology heatmap (scalp power distribution)", "Done", "Medium", 4, "Agrim Pandey"),
            ("US-37", "Spectral entropy analysis scripts", "Done", "Medium", 3, "Agrim Pandey"),
            ("US-38", "Write technical_methodology.md and Validation Strategy.docx", "Done", "High", 5, "Agrim Pandey"),
            ("US-39", "Add research paper reference (Pandey_Mittal_COMSIGPRO-2026.pdf)", "Done", "Low", 1, "Agrim Pandey"),
            ("US-40", "Dynamic result loading in visualizations (from model files)", "Not Done", "Medium", 4, "Akshat Mittal"),
            ("US-41", "Implement CNN / EEGNet baseline model", "Not Done", "High", 5, "Akshat Mittal"),
        ],
        "went_well": [
            "MNE topomap on standard 10-20 montage gave professional publication-quality scalp maps",
            "Accuracy comparison chart clearly communicated the performance gap between methods in one glance",
            "Technical documentation (methodology + validation strategy) was comprehensive and well-structured",
            "entropy_log.py captured signal complexity over time — a novel additional metric beyond power",
            "plot_test.py allowed rapid one-off visual checks without touching the main visualization script",
        ],
        "didnt_go_well": [
            "Accuracy values in visualize_eeg_results.py are hardcoded (86.0, 47.5, etc.) — will drift if models are retrained",
            "CNN (78.5%) and BiLSTM (81.2%) are labeled 'Estimated' — neither model was actually implemented",
            "Confusion matrix values are also hardcoded from one specific run — not reproducible automatically",
            "Sprint overran by 1 week due to documentation workload being underestimated",
        ],
        "learnings": [
            "All metrics should be saved to a results JSON / SQLite DB at train-time and loaded by visualization scripts",
            "If a model appears in a comparison chart, it must be at least prototyped — 'Estimated' values undermine credibility",
            "Documentation time is always underestimated — budget 30% of sprint capacity for it in a research project",
        ],
        "blockers": ["CNN/EEGNet implementation deferred due to time constraints"],
        "risks": ["Hardcoded values in charts become stale — creates silent inaccuracies in any presentation"],
    },
]

def build_sprint_cards(wb):
    ws = wb.create_sheet("📝 Sprint Details")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14

    cur_row = 1
    for sp in sprint_data:
        color = sp["color"]

        # ── Sprint Banner ──────────────────────────────────────
        merge_write(ws, cur_row, 1, cur_row, 6,
                    f"  {sp['name']}   ·   {sp['period']}   ·   Story Points: {sp['story_points']}   ·   Velocity: {sp['velocity']}",
                    font=hdr_font(12, True, C_WHITE),
                    fill_=fill(color), alignment=wrap_align("left", "center"))
        ws.row_dimensions[cur_row].height = 30
        cur_row += 1

        # ── User Stories ─────────────────────────────────────
        for c, h in enumerate(["ID", "Assignee", "User Story", "Status", "Priority", "Story Pts"], 1):
            write_cell(ws, cur_row, c, h,
                       font=hdr_font(9, True, C_WHITE),
                       fill_=fill(C_MID_BLUE),
                       alignment=center_align(), border_=border())
        ws.row_dimensions[cur_row].height = 20
        cur_row += 1

        for i, (uid, desc, status, prio, pts, assignee) in enumerate(sp["stories"]):
            bg = C_LIGHT_BLUE if i % 2 == 0 else C_WHITE
            s_color = C_GREEN_LIGHT if status == "Done" else C_AMBER_LIGHT if status == "Partial" else C_RED_LIGHT
            for c, val in enumerate([uid, assignee, desc, status, prio, pts], 1):
                bg_use = s_color if c == 4 else bg
                font_use = cell_font(10, True, C_GREEN) if (c == 4 and status == "Done") \
                      else cell_font(10, True, C_AMBER) if (c == 4 and status == "Partial") \
                      else cell_font(10, True, C_RED) if (c == 4) \
                      else cell_font(10)
                write_cell(ws, cur_row, c, val,
                           font=font_use, fill_=fill(bg_use),
                           alignment=wrap_align("center" if c != 3 else "left"),
                           border_=border())
            ws.row_dimensions[cur_row].height = 28
            cur_row += 1

        cur_row += 1  # spacer

        # ── Three-column section: Went Well / Didn't Go Well / Learnings ──
        for c, (hdr, col) in enumerate(
                [("✅ What Went Well", C_GREEN),
                 ("⚠️ What Didn't Go Well", C_RED),
                 ("📌 Key Learnings", C_MID_BLUE)], 1):
            # 2-col spans per section; pack into col pairs 1-2, 3-4, 5-6
            col_idx = (c - 1) * 2 + 1
            merge_write(ws, cur_row, col_idx, cur_row, col_idx + 1, hdr,
                        font=hdr_font(10, True, C_WHITE),
                        fill_=fill(col), alignment=center_align())
        ws.row_dimensions[cur_row].height = 22
        cur_row += 1

        sections = [sp["went_well"], sp["didnt_go_well"], sp["learnings"]]
        max_len = max(len(s) for s in sections)
        for i in range(max_len):
            for j, section in enumerate(sections):
                val = f"• {section[i]}" if i < len(section) else ""
                col_idx = j * 2 + 1
                bg_map = [C_GREEN_LIGHT, C_RED_LIGHT, C_LIGHT_BLUE]
                merge_write(ws, cur_row, col_idx, cur_row, col_idx + 1, val,
                            font=cell_font(9),
                            fill_=fill(bg_map[j]),
                            alignment=wrap_align())
            ws.row_dimensions[cur_row].height = 40
            cur_row += 1

        # ── Blockers & Risks ──────────────────────────────────
        cur_row += 1
        merge_write(ws, cur_row, 1, cur_row, 3, "🔴 Blockers",
                    font=hdr_font(10, True, C_WHITE),
                    fill_=fill(C_RED), alignment=center_align())
        merge_write(ws, cur_row, 4, cur_row, 6, "⚡ Risks",
                    font=hdr_font(10, True, C_WHITE),
                    fill_=fill(C_AMBER), alignment=center_align())
        ws.row_dimensions[cur_row].height = 20
        cur_row += 1

        b_rows = max(len(sp["blockers"]), len(sp["risks"]))
        for i in range(b_rows):
            bv = f"• {sp['blockers'][i]}" if i < len(sp["blockers"]) else ""
            rv = f"• {sp['risks'][i]}" if i < len(sp["risks"]) else ""
            merge_write(ws, cur_row, 1, cur_row, 3, bv,
                        font=cell_font(9), fill_=fill(C_RED_LIGHT),
                        alignment=wrap_align())
            merge_write(ws, cur_row, 4, cur_row, 6, rv,
                        font=cell_font(9), fill_=fill(C_AMBER_LIGHT),
                        alignment=wrap_align())
            ws.row_dimensions[cur_row].height = 30
            cur_row += 1

        cur_row += 3  # big gap between sprints

# ══════════════════════════════════════════════════════════════════
# SHEET 4 — PERFORMANCE METRICS
# ══════════════════════════════════════════════════════════════════
def build_metrics(wb):
    ws = wb.create_sheet("📊 Performance Metrics")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFG", [22, 24, 16, 14, 14, 14, 20]):
        ws.column_dimensions[col].width = w

    # Title
    merge_write(ws, 1, 1, 1, 7, "Classification Performance Summary",
                font=hdr_font(14, True, C_WHITE),
                fill_=fill(C_DARK_BLUE), alignment=center_align())
    ws.row_dimensions[1].height = 36

    hdrs = ["Method", "Task / Class Pair", "Accuracy (%)",
            "Precision", "Recall", "F1 Score", "Notes"]
    for c, h in enumerate(hdrs, 1):
        write_cell(ws, 2, c, h,
                   font=hdr_font(10, True, C_WHITE),
                   fill_=fill(C_MID_BLUE),
                   alignment=center_align(), border_=border())
    ws.row_dimensions[2].height = 22

    results = [
        ("Statistical Features + SVM", "Movement vs. Rest (Binary)", 86.0, "~0.87", "~0.86", "~0.86",
         "Best result. 288 features (Power, RMS, Entropy, Band Powers). GridSearchCV optimised."),
        ("CSP + SVM", "DLF vs. DRF (Foot L vs. R)", 47.5, "~0.48", "~0.47", "~0.47",
         "Near chance. Ledoit-Wolf regularisation used. 16-ch hardware bottleneck."),
        ("FBCSP + LDA", "DLF vs. DRF (Foot L vs. R)", 47.1, "~0.47", "~0.47", "~0.47",
         "7-band filter bank + MI selection (top-10 features). Marginally below CSP+SVM."),
        ("FBCSP + SVM", "DLF vs. DRF (Foot L vs. R)", 47.3, "~0.47", "~0.47", "~0.47",
         "LDA slightly outperformed SVM on FBCSP features in this trial."),
        ("CNN (Estimated)", "DLF vs. DRF (Foot L vs. R)", 78.5, "—", "—", "—",
         "ESTIMATED — not implemented. Placeholder for future EEGNet/CNN work."),
        ("BiLSTM (Estimated)", "DLF vs. DRF (Foot L vs. R)", 81.2, "—", "—", "—",
         "ESTIMATED — not implemented. Placeholder for future temporal DL model."),
    ]

    for i, row in enumerate(results, start=3):
        even = (i % 2 == 0)
        bg = C_LIGHT_BLUE if even else C_WHITE
        acc = row[2]
        if acc >= 80:
            acc_bg = C_GREEN_LIGHT; acc_font = cell_font(10, True, C_GREEN)
        elif acc >= 60:
            acc_bg = C_AMBER_LIGHT; acc_font = cell_font(10, True, C_AMBER)
        else:
            acc_bg = C_RED_LIGHT; acc_font = cell_font(10, True, C_RED)

        for c, val in enumerate(row, 1):
            bg_use = acc_bg if c == 3 else bg
            fnt = acc_font if c == 3 else cell_font(10)
            write_cell(ws, i, c, val,
                       font=fnt, fill_=fill(bg_use),
                       alignment=wrap_align("center" if c in [1,2,3,4,5,6] else "left"),
                       border_=border(),
                       number_format="0.0" if c == 3 else None)
        ws.row_dimensions[i].height = 36

    # ── Confusion Matrices ───────────────────────────────────
    cur_row = len(results) + 5

    merge_write(ws, cur_row, 1, cur_row, 7, "Confusion Matrix Results (CSP + SVM — DLF vs DRF)",
                font=hdr_font(11, True, C_WHITE),
                fill_=fill(C_MID_BLUE), alignment=center_align())
    cur_row += 1

    cm_hdrs = ["", "Predicted: DLF", "Predicted: DRF", "", "", "Predicted: DLF", "Predicted: DRF"]
    for c, h in enumerate(cm_hdrs, 1):
        write_cell(ws, cur_row, c, h,
                   font=hdr_font(9, True, C_WHITE),
                   fill_=fill(C_DARK_BLUE if h else "FFFFFF"),
                   alignment=center_align(), border_=border())
    cur_row += 1

    cm_rows = [
        ["Actual: DLF", 87, 33, "", "Actual: DLF", 62, 58],
        ["Actual: DRF", 93, 27, "", "Actual: DRF", 69, 51],
    ]
    cm_label = ["← CSP + SVM", "", "", "", "← FBCSP + LDA", "", ""]
    for j, (cm_row, label) in enumerate(zip(cm_rows, cm_label)):
        for c, val in enumerate(cm_row, 1):
            bg = C_GREEN_LIGHT if (c == 2 and j == 0) or (c == 3 and j == 1) else \
                 C_RED_LIGHT if (c == 3 and j == 0) or (c == 2 and j == 1) else C_WHITE
            write_cell(ws, cur_row, c, val,
                       font=cell_font(10, bold=(c in [2,3,6,7])),
                       fill_=fill(bg),
                       alignment=center_align(), border_=border())
        ws.row_dimensions[cur_row].height = 24
        cur_row += 1

# ══════════════════════════════════════════════════════════════════
# SHEET 5 — RETROSPECTIVE ACTIONS & NEXT STEPS
# ══════════════════════════════════════════════════════════════════
def build_actions(wb):
    ws = wb.create_sheet("🚀 Next Steps & Actions")
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [10, 44, 24, 16, 20]):
        ws.column_dimensions[col].width = w

    merge_write(ws, 1, 1, 1, 5,
                "Retrospective Action Items & Suggested Next Sprints",
                font=hdr_font(14, True, C_WHITE),
                fill_=fill(C_DARK_BLUE), alignment=center_align())
    ws.row_dimensions[1].height = 36

    # Immediate action items
    merge_write(ws, 2, 1, 2, 5, "Immediate Action Items (Technical Debt from Sprint 5)",
                font=hdr_font(11, True, C_WHITE),
                fill_=fill(C_MID_BLUE), alignment=center_align())

    hdrs = ["#", "Action Item", "Reason / Impact", "Priority", "Estimated Effort"]
    for c, h in enumerate(hdrs, 1):
        write_cell(ws, 3, c, h,
                   font=hdr_font(9, True, C_WHITE),
                   fill_=fill(C_MID_BLUE),
                   alignment=center_align(), border_=border())

    actions = [
        (1, "Replace hardcoded accuracy values in visualize_eeg_results.py with dynamic loading from result JSON files",
         "Prevents stale charts if models are retrained", "HIGH", "2 hours"),
        (2, "Save model accuracy, confusion matrix, and CV scores to results.json at end of every train script",
         "Enables reproducible reporting and version tracking", "HIGH", "3 hours"),
        (3, "Implement at least one CNN baseline (EEGNet or ShallowConvNet)",
         "Validates DL comparison values currently labeled 'Estimated'", "HIGH", "1–2 weeks"),
        (4, "Remove dead code sections in download_prep.py (duplicate Step 1/2 headers)",
         "Code hygiene; avoids confusion when onboarding new collaborators", "MEDIUM", "30 min"),
        (5, "Benchmark train_csp_sliding.py results and add to visualization",
         "Sliding-window CSP may improve temporal discrimination", "MEDIUM", "1 day"),
        (6, "Write a README.md documenting script execution order and output schema",
         "Critical for reproducibility and handoff to teacher/evaluator", "HIGH", "2 hours"),
        (7, "Add Delta (0.5–4 Hz) and Theta (4–8 Hz) bands to FIR filtering",
         "These bands carry motor planning signals sometimes useful for intent detection", "LOW", "3 hours"),
    ]

    for i, (num, item, reason, prio, effort) in enumerate(actions, start=4):
        bg = C_LIGHT_BLUE if i % 2 == 0 else C_WHITE
        p_bg = C_RED_LIGHT if prio == "HIGH" else C_AMBER_LIGHT if prio == "MEDIUM" else C_GREEN_LIGHT
        p_font = cell_font(10, True, C_RED) if prio == "HIGH" else \
                 cell_font(10, True, C_AMBER) if prio == "MEDIUM" else \
                 cell_font(10, True, C_GREEN)
        for c, val in enumerate([num, item, reason, prio, effort], 1):
            bg_use = p_bg if c == 4 else bg
            fnt_use = p_font if c == 4 else cell_font(10)
            write_cell(ws, i, c, val,
                       font=fnt_use, fill_=fill(bg_use),
                       alignment=wrap_align("center" if c in [1,4,5] else "left"),
                       border_=border())
        ws.row_dimensions[i].height = 36

    # Suggested next sprints
    cur_row = len(actions) + 6

    merge_write(ws, cur_row, 1, cur_row, 5, "Suggested Future Sprints",
                font=hdr_font(11, True, C_WHITE),
                fill_=fill(C_DARK_BLUE), alignment=center_align())
    ws.row_dimensions[cur_row].height = 28
    cur_row += 1

    hdrs2 = ["Sprint", "Theme", "Key Deliverables", "Expected Outcome", "Duration"]
    for c, h in enumerate(hdrs2, 1):
        write_cell(ws, cur_row, c, h,
                   font=hdr_font(9, True, C_WHITE),
                   fill_=fill(C_MID_BLUE),
                   alignment=center_align(), border_=border())
    ws.row_dimensions[cur_row].height = 20
    cur_row += 1

    future = [
        ("Sprint 6", "Deep Learning Baseline",
         "Implement EEGNet / ShallowConvNet for DLF vs. DRF. Compare against CSP+SVM.",
         "Validate or refute estimated 78–81% DL accuracy", "3–4 weeks"),
        ("Sprint 7", "Real-Time Inference",
         "Integrate LSL (Lab Streaming Layer) for live EEG streaming. Apply FIR filter in real-time. Run SVM inference and display prediction.",
         "End-to-end real-time BCI demonstration", "4–5 weeks"),
        ("Sprint 8", "GUI Dashboard",
         "Build a Tkinter/PyQt dashboard showing live scalp heatmap, band power bars, predicted class, and confidence score.",
         "Demo-ready non-technical presentation interface", "3 weeks"),
        ("Sprint 9", "64-channel Upgrade / Transfer Learning",
         "Acquire or simulate 64-channel EEG for foot MI. Re-run CSP/FBCSP to confirm accuracy improvement hypothesis.",
         "Resolve the 16ch vs. 64ch channel bottleneck question", "Ongoing"),
    ]

    colors_future = ["3B5998", "217346", "9C6500", "7B3F8C"]
    for i, (sprint, theme, deliv, outcome, dur) in enumerate(future):
        bg = C_LIGHT_BLUE if i % 2 == 0 else C_WHITE
        for c, val in enumerate([sprint, theme, deliv, outcome, dur], 1):
            bg_use = colors_future[i] if c == 1 else bg
            fnt_use = hdr_font(10, True, C_WHITE) if c == 1 else cell_font(10)
            write_cell(ws, cur_row, c, val,
                       font=fnt_use, fill_=fill(bg_use if c == 1 else bg),
                       alignment=wrap_align("center" if c in [1,2,5] else "left"),
                       border_=border())
        ws.row_dimensions[cur_row].height = 48
        cur_row += 1

# ══════════════════════════════════════════════════════════════════
# SHEET 6 — RETROSPECTIVE BOARD  (row = item, col = section)
# ══════════════════════════════════════════════════════════════════
def build_retro_board(wb):
    ws = wb.create_sheet("Retrospective Board")
    ws.sheet_view.showGridLines = False

    # ── Column widths ─────────────────────────────────────────────
    col_widths = [46, 46, 46, 50, 50]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Section column config ─────────────────────────────────────
    col_headers = [
        ("What Went Well",            "217346", C_GREEN_LIGHT),
        ("What Went Poorly",          "9C0006", C_RED_LIGHT),
        ("What Ideas Do You Have",    "2E4A8A", C_LIGHT_BLUE),
        ("How Should We Take Action", "9C6500", C_AMBER_LIGHT),
        ("Benefits",                  "4A235A", "EDE0F0"),
    ]

    # ── Sprint data: each sprint has a list of items per column ───
    # Items are aligned: item[0] of each col goes on the same row, etc.
    sprint_data_board = [
        {
            "label": "Sprint 1  |  Dec 1, 2025 – Jan 15, 2026  |  Data Preparation & FIR Filtering",
            "color": "1F4E79",
            "items": {
                "well": [
                    "File naming regex parsed all subjects perfectly with zero manual corrections",
                    "Metadata schema (8 columns) reused unchanged by every downstream script across all subsequent work",
                    "Natural sort key (patient_sort_key) fixed S1 < S2 ordering — prevents lexicographic S1/S10/S11 confusion",
                    "101-tap (odd) linear-phase FIR filter — mathematically rigorous choice for EEG band extraction",
                    "filtfilt zero-phase filtering eliminates phase distortion that would corrupt feature timing",
                    "Synthetic signal test (10+20+40 Hz composite) visually confirmed each band in isolation",
                    "Skip-if-exists logic saved ~6 hrs of recomputation across 25+ subject folders",
                    "Column naming {ch}_{band} (e.g. 0_a, 0_b) self-documents the output CSV schema",
                ],
                "poor": [
                    "Dead / duplicate Step headers left in download_prep.py after download code was removed",
                    "No schema validation — patient_number type errors traced back manually weeks later",
                    "No unit tests at any stage — label_map and regex bugs found by manual inspection only",
                    "Gamma band (30-50 Hz) sits close to Nyquist limit (62.5 Hz) — roll-off is slightly compressed",
                    "Delta (0.5-4 Hz) and Theta (4-8 Hz) bands were omitted — limits future low-frequency ERP analysis",
                    "Output CSV column order assumed by all downstream scripts but never formally documented",
                    "Disk usage not monitored — fir_dataset grew unexpectedly large for 25+ subjects",
                ],
                "ideas": [
                    "Create validate_metadata.py with dtype checks, row count assertions, and null-URL detection",
                    "Auto-generate a data summary cross-tab: trial counts per class per subject",
                    "Write project README.md with full pipeline overview and execution order on Day 1",
                    "Extend BANDS dict with 'delta': (0.5, 4) and 'theta': (4, 8) for richer frequency coverage",
                    "Generate a filter_spec.md documenting order, transition width, and stopband attenuation dB",
                    "Add a shutil.disk_usage() check before starting any batch processing job",
                ],
                "action": [
                    "validate_metadata.py: assert len(df)==expected, dtypes correct, no NaN in local_url column",
                    "Remove all duplicate Step 1/2 comment blocks from download_prep.py immediately",
                    "Add a 3-row test fixture with known-good filenames to unit-test the regex parse logic",
                    "Add 'delta' and 'theta' bands to BANDS dict in fir_main.py before next run",
                    "Write filter_spec.md: for each band document order, passband ripple dB, stopband attenuation dB",
                    "Add shutil.disk_usage() guard at start of fir_main.main() — warn if < 10 GB free",
                ],
                "benefits": [
                    "Clean metadata = reliable contract for all future scripts; a single error here cascades everywhere",
                    "Early type validation would have prevented ~3 hrs of debugging during feature extraction",
                    "Validated FIR filters = provably correct signal inputs for all ML models downstream",
                    "Synthetic test doubles as a regression test — re-run whenever fir_main.py changes",
                    "Skip-if-exists logic saves hours of compute time on every full pipeline re-run",
                    "Documented column schema prevents silent data-misalignment bugs in CSP data preparation",
                ],
            }
        },
        {
            "label": "Sprint 2  |  Jan 16 – Mar 7, 2026  |  Feature Extraction, SVM & CSP / FBCSP",
            "color": "1D6A4A",
            "items": {
                "well": [
                    "Modular feature functions (calculate_power, calculate_rms, calculate_spectral_entropy) are independently testable",
                    "StandardScaler fit only on training data — textbook-correct, zero data leakage into the test set",
                    "GridSearchCV with n_jobs=-1 parallelised SVM hyperparameter search efficiently",
                    "86% binary accuracy (Movement vs. Rest) — strong, credible baseline for a 16-channel EEG setup",
                    "feature_names.json provides full traceability from any feature index to its source channel and band",
                    "Ledoit-Wolf regularisation robustly handled ill-conditioned covariance matrices (common with 16 channels)",
                    "FBCSP correctly split train/test BEFORE CSP fitting — avoids the most common leakage bug in BCI literature",
                    "Mutual Information feature selection reduced 28 FBCSP features to 10 with principled statistical ranking",
                    "scikit-learn Pipeline (CSP -> SVC) made 5-fold cross-validation structurally leak-free by design",
                ],
                "poor": [
                    "Feature matrix is very wide (288 = 16ch x 3band x 6stats) — curse of dimensionality risk for SVM kernel",
                    "feature_names.json saved to disk but never used for importance ranking or feature selection guidance",
                    "Multi-class SVM performance (all 8 labels) was never benchmarked — only binary pairs evaluated",
                    "Welch PSD with nperseg=256 gives coarse frequency resolution at 125 Hz sampling rate",
                    "CSP + SVM for foot MI (DLF vs. DRF) only ~47.5% — near chance — due to 16-channel spatial limitation",
                    "FBCSP (47.1%) showed adding 7 sub-bands did not compensate for insufficient electrode spatial resolution",
                    "train_csp_sliding.py partially written but never benchmarked or integrated into any results comparison",
                    "CSP filter matrices and per-run accuracy were never logged — impossible to systematically compare runs",
                    "Should have validated the spatial pipeline on easy hand MI (CLH vs. CRH) before attempting foot MI",
                ],
                "ideas": [
                    "Apply PCA (n_components=50) or SelectKBest to reduce feature dimensionality before SVM training",
                    "Build a feature importance bar chart using feature_names.json mapped to SVM decision weights",
                    "Run multi-class SVM on all 8 labels first; use confusion matrix to justify binary pair selection",
                    "Increase Welch nperseg to min(len(x), 512) for finer frequency resolution",
                    "Run CSP + SVM on CLH vs. CRH (hand MI) first to prove the spatial pipeline is correct",
                    "Log each train run: run_id, timestamp, accuracy, confusion matrix, and best_params to a JSON file",
                    "Finish train_csp_sliding.py and benchmark sliding-window CSP for temporal discrimination gains",
                    "Explore EEGNet (compact CNN built for low-channel-count EEG) as a spatial alternative",
                ],
                "action": [
                    "Insert PCA(n_components=50) step in feature_extraction.py pipeline before StandardScaler",
                    "Load feature_names.json in train_svm.py; print top-20 features by absolute SVM coefficient weight",
                    "Add a multi-class SVM block (all 8 labels) before the TARGET_LABELS binary filter in train_svm.py",
                    "Set nperseg=min(len(x), 512) in all Welch-based feature functions in feature_extraction.py",
                    "Run prepare_csp_data.py with TARGET_LABELS=[1,2] (CLH,CRH) and record the CSP accuracy",
                    "Add json.dump({'run_id', 'accuracy', 'cm', 'params'}) at end of train_csp.py and train_fbcsp.py",
                    "Complete train_csp_sliding.py with a grid over window sizes (0.5s, 1s, 2s) and report all results",
                    "Prototype EEGNet in PyTorch using X_raw.npy + y_csp.npy for a fair direct comparison",
                ],
                "benefits": [
                    "86% binary accuracy validates the full pipeline end-to-end before investing in spatial methods",
                    "Modular feature functions reused cleanly in both CSP and FBCSP workflows — no duplication",
                    "Saved scaler.joblib makes the statistical-feature pipeline deployment-ready for real-time inference",
                    "feature_names.json enables future SHAP / explainability analysis with zero additional instrumentation",
                    "Ledoit-Wolf regularisation is directly reusable for any future covariance-based spatial method",
                    "Correct train/test split discipline = trustworthy, publishable accuracy numbers with no inflation",
                    "Documenting the 16-channel bottleneck saves future researchers from repeating ineffective experiments",
                    "MI-based feature selection approach is transferable to any future high-dimensional classification task",
                ],
            }
        },
        {
            "label": "Sprint 3  |  Mar 8 – Apr 6, 2026  |  Visualization, Tuning & Documentation",
            "color": "7B1A1A",
            "items": {
                "well": [
                    "MNE topomap on standard 10-20 montage produced publication-quality scalp spatial filter maps",
                    "Accuracy comparison bar chart communicated the performance gap between all methods clearly at a glance",
                    "Technical methodology and Validation Strategy documents were comprehensive and well-structured",
                    "entropy_log.py introduced spectral entropy as a novel signal complexity metric beyond raw band power",
                    "plot_test.py enabled rapid one-off visual checks without modifying the main visualization script",
                    "Topology and heatmap scripts provided additional spatial insight into channel-level power distribution",
                ],
                "poor": [
                    "Accuracy values in visualize_eeg_results.py are hardcoded (86.0, 47.5) — will silently go stale if models are retrained",
                    "CNN (78.5%) and BiLSTM (81.2%) entries labeled 'Estimated' — neither model was actually implemented",
                    "Confusion matrix values hardcoded from one specific run — not reproducible from the script alone",
                    "Sprint overran by 1 week due to documentation workload being severely underestimated at planning",
                    "No automated test to catch visual output regressions — charts could silently break on library updates",
                ],
                "ideas": [
                    "Build a results.json that every train script automatically writes accuracy + CM + params to at training end",
                    "Refactor visualize_eeg_results.py to load all chart values dynamically from results.json",
                    "Implement EEGNet / ShallowConvNet prototype with PyTorch even if not fully hyperparameter-tuned",
                    "Budget 30% of every future sprint's capacity explicitly for documentation and review tasks",
                    "Add a simple chart validation test: load each saved PNG and assert file size > 10 KB",
                ],
                "action": [
                    "Add save_results(metrics_dict, path) helper function to a shared utils.py; call from every train script",
                    "In visualize_eeg_results.py replace hardcoded accuracy dicts with json.load(open('results.json')) calls",
                    "Create train_eegnet.py as next sprint Priority 1 story (5 story points, HIGH priority)",
                    "Add 'Documentation Review' as a standing 5-point ticket in every subsequent sprint backlog",
                    "Add os.path.getsize() assertions in a test_charts.py to catch empty / failed chart saves",
                ],
                "benefits": [
                    "Dynamic chart loading = accuracy values always reflect the latest trained model, never stale",
                    "EEGNet implementation would validate or refute the 78-81% estimated DL accuracy benchmark",
                    "Proper documentation budgeting prevents sprint overruns and last-minute documentation quality cuts",
                    "Publication-quality topomaps and accuracy comparison charts are directly usable in the final project report",
                    "Chart regression tests protect visualization quality as MNE/Matplotlib library versions change",
                ],
            }
        },
    ]

    sprint_col_colors = [h[1] for h in col_headers]
    sprint_col_bgs    = [h[2] for h in col_headers]
    col_keys = ["well", "poor", "ideas", "action", "benefits"]

    # ── Row 1: Main title ─────────────────────────────────────────
    merge_write(ws, 1, 1, 1, 5,
                "MILimbEEG BCI Project  -  Sprint Retrospective Board  (Dec 2025 - Apr 2026)",
                font=hdr_font(13, True, C_WHITE),
                fill_=fill(C_DARK_BLUE),
                alignment=center_align())
    ws.row_dimensions[1].height = 36

    # ── Row 2: Section column headers ─────────────────────────────
    for c, (title, color, _) in enumerate(col_headers, 1):
        write_cell(ws, 2, c, title,
                   font=hdr_font(11, True, C_WHITE),
                   fill_=fill(color),
                   alignment=center_align(),
                   border_=border())
    ws.row_dimensions[2].height = 28

    cur_row = 3

    for sp in sprint_data_board:
        # ── Sprint date/label header row (spans all 5 cols) ───────
        merge_write(ws, cur_row, 1, cur_row, 5,
                    sp["label"],
                    font=hdr_font(11, True, C_WHITE),
                    fill_=fill(sp["color"]),
                    alignment=wrap_align("left", "center"))
        ws.row_dimensions[cur_row].height = 24
        cur_row += 1

        # Compute number of rows needed (max items across all 5 cols)
        col_items = [sp["items"][k] for k in col_keys]
        n_rows = max(len(lst) for lst in col_items)

        for r in range(n_rows):
            even = (r % 2 == 0)
            for c_idx, (items, bg) in enumerate(zip(col_items, sprint_col_bgs), 1):
                val = items[r] if r < len(items) else ""
                # Alternate row shading: full bg on even rows, slightly lighter on odd
                row_bg = bg if even else C_WHITE
                write_cell(ws, cur_row, c_idx, val,
                           font=cell_font(9),
                           fill_=fill(row_bg),
                           alignment=wrap_align(),
                           border_=border())
            ws.row_dimensions[cur_row].height = 52
            cur_row += 1

        # Thin spacer between sprints
        merge_write(ws, cur_row, 1, cur_row, 5, "",
                    fill_=fill("D9D9D9"))
        ws.row_dimensions[cur_row].height = 5
        cur_row += 1

    # ── Overall Summary header row ────────────────────────────────
    merge_write(ws, cur_row, 1, cur_row, 5,
                "OVERALL PROJECT SUMMARY  (All Sprints Combined)",
                font=hdr_font(12, True, C_WHITE),
                fill_=fill(C_DARK_BLUE), alignment=center_align())
    ws.row_dimensions[cur_row].height = 26
    cur_row += 1

    overall_items = [
        [   # What Went Well
            "Full end-to-end EEG pipeline built and validated in ~19 weeks",
            "86% binary SVM accuracy (movement vs. rest) achieved",
            "FIR filter correctness confirmed with rigorous synthetic signal test",
            "FBCSP implemented with zero data leakage — rare in BCI research",
            "Modular codebase: every script is independently runnable",
        ],
        [   # What Went Poorly
            "Accuracy values hardcoded in visualization — drift risk on retrain",
            "CNN/BiLSTM listed as estimated, never actually built",
            "Missing unit tests across all 5 sprints",
            "Dead code and no README hurt project handoff quality",
            "CSP foot MI near chance (47%) — wrong class pair chosen first",
        ],
        [   # Ideas
            "EEGNet / ShallowConvNet prototype for genuine DL comparison",
            "Real-time LSL stream integration for live BCI demo",
            "Delta and Theta band filtering for motor planning signals",
            "64-channel upgrade to resolve the spatial resolution bottleneck",
            "results.json auto-logging from every train script",
        ],
        [   # Action
            "Build train_eegnet.py as Sprint 6 P1 deliverable",
            "Write save_results() helper in utils.py; call from all train scripts",
            "Refactor visualize_eeg_results.py to read from results.json",
            "Write README.md with full script execution order + output schema",
            "Run CSP on CLH vs. CRH to validate spatial pipeline correctness",
        ],
        [   # Benefits
            "Publishable pipeline: FIR + CSP + FBCSP comparison is original research",
            "Dynamic result loading prevents silent stale-accuracy presentation errors",
            "Documented channel bottleneck finding saves future teams redundant effort",
            "Spatial topomaps and accuracy charts ready for the final project report",
            "Modular architecture extensible to real-time, 64-ch, and DL future work",
        ],
    ]

    n_overall = max(len(lst) for lst in overall_items)
    for r in range(n_overall):
        even = (r % 2 == 0)
        for c_idx, (items, bg) in enumerate(zip(overall_items, sprint_col_bgs), 1):
            val = items[r] if r < len(items) else ""
            row_bg = bg if even else C_WHITE
            write_cell(ws, cur_row, c_idx, val,
                       font=cell_font(9, bold=True),
                       fill_=fill(row_bg),
                       alignment=wrap_align(),
                       border_=border())
        ws.row_dimensions[cur_row].height = 52
        cur_row += 1


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    build_overview(wb)
    build_timeline(wb)
    build_sprint_cards(wb)
    build_metrics(wb)
    build_actions(wb)
    build_retro_board(wb)

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"[DONE] Sprint Retrospective saved to: {OUTPUT_PATH}")

if __name__ == "__main__":
    main()

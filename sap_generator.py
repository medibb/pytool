#!/usr/bin/env python3
"""
sap_generator.py — Statistical Analysis Plan (SAP) 자동 생성기

TransCelerate Common SAP Template 구조 + ICH E9(R1) estimand 프레임워크 기반으로
config dict 하나에서 SAP 문서(docx) + sample size 계산 + power curve 생성.

=== 핵심 기능 ===
1. SAP docx 생성: TransCelerate 표준 섹션 구조
2. Sample size 계산: statsmodels 기반 (G*Power 대체)
3. Power curve: matplotlib → PNG → docx 삽입
4. Sensitivity table: 다양한 effect size / SD 조합
5. FSHD 프로젝트 검증 완료

=== 사용법 ===

from sap_generator import generate_sap

config = {
    "study": {
        "title": "...",
        "protocol_number": "PNUH-2026-XXX",
        "phase": "Investigator-initiated",
        "sponsor": "부산대학교병원 재활의학과",
        "pi": "이재현",
        "sap_version": "1.0",
        "sap_date": "2026-04-01",
    },
    "design": {
        "type": "rct_parallel",  # rct_parallel, rct_crossover, cohort, case_control
        "arms": [
            {"name": "Anterior approach", "ratio": 1},
            {"name": "Posterior approach", "ratio": 1},
        ],
        "blinding": "single_blind_assessor",
        "randomization": {
            "method": "block",
            "block_sizes": [4, 6],
            "stratification": ["sex"],
        },
    },
    "objectives": {
        "primary": {
            "description": "Compare SPADI total score change at 3 months",
            "hypothesis": "Anterior approach is non-inferior to posterior approach",
        },
        "secondary": [
            "Compare shoulder ROM (ER, IR, ABD, flexion) change at 3 months",
            "Compare pain NRS change at 3 months",
        ],
        "exploratory": [
            "Ultrasound findings (SST integrity, biceps effusion) at 3 months",
        ],
    },
    "primary_outcome": {
        "name": "SPADI total score change at 3 months",
        "type": "continuous",   # continuous, binary, time_to_event
        "test": "independent_t",  # independent_t, paired_t, anova, chi_square, ...
        "mcid": 13,
        "sd": 20,
        "alpha": 0.05,
        "power": 0.80,
        "sides": 2,
        "dropout_rate": 0.20,
    },
    "secondary_outcomes": [
        {"name": "Shoulder ER ROM (degrees)", "type": "continuous"},
        {"name": "Pain NRS (0-10)", "type": "continuous"},
        {"name": "SST score (0-12)", "type": "continuous"},
    ],
    "populations": {
        "itt": "All randomized subjects",
        "pp": "Subjects completing 3-month follow-up without major protocol deviation",
        "safety": "All subjects who received injection",
    },
    "analysis": {
        "primary_method": "Independent samples t-test (or ANCOVA adjusting for baseline)",
        "missing_data": "Multiple imputation (MICE)",
        "multiplicity": "Bonferroni correction for secondary outcomes",
        "sensitivity": [
            "Per-protocol analysis",
            "Worst-case imputation",
            "ANCOVA with baseline as covariate",
        ],
        "subgroup": ["sex", "diabetes_status", "dominant_hand_affected"],
        "software": "Python (scipy, statsmodels), R (as needed)",
        "significance_level": "Two-sided α = 0.05",
    },
}

generate_sap(config, output_path="./SAP.docx")

CLI:
    python3 sap_generator.py config.json [--output SAP.docx]
    python3 sap_generator.py --example > example_config.json
    python3 sap_generator.py --example --run   # 예시 config로 바로 생성

의존성: python-docx, statsmodels, matplotlib, numpy, scipy
"""

import os
import sys
import json
import math
import argparse
import tempfile
from datetime import datetime

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from statsmodels.stats.power import (
    TTestIndPower,
    TTestPower,
    FTestAnovaPower,
    GofChisquarePower,
)


# ============================================================
# Sample Size Calculation Engine
# ============================================================

def calculate_sample_size(outcome_config):
    """
    Calculate sample size from outcome config dict.

    Parameters
    ----------
    outcome_config : dict
        Must contain: type, test, mcid, sd, alpha, power, sides
        Optional: dropout_rate, ratio, k_groups, proportion1, proportion2

    Returns
    -------
    dict with keys:
        n_per_group, n_total, n_adjusted (dropout-adjusted),
        effect_size, formula_used, assumptions
    """
    oc = outcome_config
    test = oc.get("test", "independent_t")
    alpha = oc.get("alpha", 0.05)
    power = oc.get("power", 0.80)
    sides = oc.get("sides", 2)
    dropout = oc.get("dropout_rate", 0.0)
    ratio = oc.get("ratio", 1.0)
    alt = "two-sided" if sides == 2 else "larger"

    result = {"formula_used": "", "assumptions": {}}

    if test in ("independent_t", "ancova"):
        mcid = oc["mcid"]
        sd = oc["sd"]
        effect_size = mcid / sd
        analysis = TTestIndPower()
        n1 = analysis.solve_power(
            effect_size=effect_size, alpha=alpha, power=power,
            ratio=ratio, alternative=alt
        )
        n_per_group = int(math.ceil(n1))
        n_arms = 2
        result["formula_used"] = (
            f"Two-sample t-test: n = 2 × ((z_α/2 + z_β) × SD / MCID)²"
        )
        result["assumptions"] = {
            "MCID": mcid, "SD": sd, "Effect size (d)": round(effect_size, 3),
            "Alpha": alpha, "Power": power, "Sides": sides,
        }

    elif test == "paired_t":
        mcid = oc["mcid"]
        sd = oc["sd"]
        effect_size = mcid / sd
        analysis = TTestPower()
        n = analysis.solve_power(
            effect_size=effect_size, alpha=alpha, power=power,
            alternative=alt
        )
        n_per_group = int(math.ceil(n))
        n_arms = 1
        result["formula_used"] = "Paired t-test: n = ((z_α/2 + z_β) × SD_diff / MCID)²"
        result["assumptions"] = {
            "MCID": mcid, "SD of differences": sd,
            "Effect size (d)": round(effect_size, 3),
            "Alpha": alpha, "Power": power, "Sides": sides,
        }

    elif test == "anova":
        mcid = oc["mcid"]
        sd = oc["sd"]
        k = oc.get("k_groups", 3)
        effect_size_f = (mcid / sd) / math.sqrt(2)
        analysis = FTestAnovaPower()
        n_total_raw = analysis.solve_power(
            effect_size=effect_size_f, alpha=alpha, power=power,
            k_groups=k
        )
        n_per_group = int(math.ceil(n_total_raw / k))
        n_arms = k
        result["formula_used"] = f"One-way ANOVA F-test (k={k} groups)"
        result["assumptions"] = {
            "MCID": mcid, "SD": sd, "Cohen's f": round(effect_size_f, 3),
            "Groups": k, "Alpha": alpha, "Power": power,
        }

    elif test == "chi_square":
        p1 = oc["proportion1"]
        p2 = oc["proportion2"]
        from statsmodels.stats.proportion import proportion_effectsize
        from statsmodels.stats.power import NormalIndPower
        es = proportion_effectsize(p1, p2)
        analysis = NormalIndPower()
        n1 = analysis.solve_power(
            effect_size=abs(es), alpha=alpha, power=power,
            ratio=ratio, alternative=alt
        )
        n_per_group = int(math.ceil(n1))
        n_arms = 2
        effect_size = abs(es)
        result["formula_used"] = "Two-proportion z-test (arcsine transformation)"
        result["assumptions"] = {
            "Proportion 1": p1, "Proportion 2": p2,
            "Effect size (h)": round(effect_size, 3),
            "Alpha": alpha, "Power": power,
        }

    else:
        raise ValueError(f"Unsupported test type: {test}")

    n_total = n_per_group * n_arms if test != "paired_t" else n_per_group
    n_adjusted = int(math.ceil(n_total / (1 - dropout))) if dropout > 0 else n_total
    n_per_group_adj = int(math.ceil(n_per_group / (1 - dropout))) if dropout > 0 else n_per_group

    result.update({
        "n_per_group": n_per_group,
        "n_per_group_adjusted": n_per_group_adj,
        "n_total": n_total,
        "n_total_adjusted": n_adjusted,
        "dropout_rate": dropout,
    })
    return result


def generate_sensitivity_table(outcome_config, mcid_range=None, sd_range=None):
    """
    Generate sensitivity table: n_per_group for various MCID × SD combinations.

    Returns list of dicts for table rendering.
    """
    oc = outcome_config.copy()
    base_mcid = oc["mcid"]
    base_sd = oc["sd"]

    if mcid_range is None:
        mcid_range = [
            round(base_mcid * 0.7, 1),
            round(base_mcid * 0.85, 1),
            base_mcid,
            round(base_mcid * 1.15, 1),
            round(base_mcid * 1.3, 1),
        ]
    if sd_range is None:
        sd_range = [
            round(base_sd * 0.8, 1),
            base_sd,
            round(base_sd * 1.2, 1),
            round(base_sd * 1.4, 1),
        ]

    rows = []
    for mcid in mcid_range:
        row = {"MCID": mcid}
        for sd in sd_range:
            oc_temp = oc.copy()
            oc_temp["mcid"] = mcid
            oc_temp["sd"] = sd
            try:
                res = calculate_sample_size(oc_temp)
                row[f"SD={sd}"] = res["n_per_group_adjusted"]
            except Exception:
                row[f"SD={sd}"] = "—"
        rows.append(row)
    return rows, mcid_range, sd_range


def generate_power_curve(outcome_config, output_path):
    """
    Generate power curve plot and save as PNG.

    X-axis: sample size per group
    Curves: multiple effect sizes around the specified MCID/SD
    """
    oc = outcome_config
    test = oc.get("test", "independent_t")
    alpha = oc.get("alpha", 0.05)
    sd = oc["sd"]
    mcid = oc["mcid"]
    sides = oc.get("sides", 2)
    alt = "two-sided" if sides == 2 else "larger"

    base_d = mcid / sd
    effect_sizes = [
        round(base_d * 0.6, 3),
        round(base_d * 0.8, 3),
        round(base_d, 3),
        round(base_d * 1.2, 3),
    ]

    n_range = np.arange(5, 101, 1)

    fig, ax = plt.subplots(figsize=(7, 4.5))

    if test in ("independent_t", "ancova"):
        analysis = TTestIndPower()
    elif test == "paired_t":
        analysis = TTestPower()
    else:
        analysis = TTestIndPower()

    for d in effect_sizes:
        powers = []
        for n in n_range:
            if test == "paired_t":
                p = analysis.power(effect_size=d, nobs=n, alpha=alpha, alternative=alt)
            else:
                p = analysis.power(effect_size=d, nobs1=n, alpha=alpha,
                                   ratio=1.0, alternative=alt)
            powers.append(p)

        mcid_equiv = round(d * sd, 1)
        ax.plot(n_range, powers, label=f"d={d:.2f} (MCID={mcid_equiv})", linewidth=1.5)

    ax.axhline(y=0.80, color="red", linestyle="--", alpha=0.6, label="Power = 0.80")
    ax.axhline(y=0.90, color="orange", linestyle="--", alpha=0.4, label="Power = 0.90")

    ax.set_xlabel("Sample Size per Group (n)", fontsize=11)
    ax.set_ylabel("Statistical Power (1 − β)", fontsize=11)
    ax.set_title("Power Curve — Sample Size vs Statistical Power", fontsize=12)
    ax.legend(fontsize=8, loc="lower right")
    ax.set_ylim(0, 1.05)
    ax.set_xlim(5, 100)
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches="tight")
    plt.close()
    return output_path


# ============================================================
# DOCX Generation — TransCelerate SAP Structure
# ============================================================

def _add_heading(doc, text, level=1):
    """Add heading with consistent styling."""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
    return h


def _add_table(doc, headers, rows, col_widths=None):
    """Add a formatted table to the document."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = str(h)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = str(val)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)

    return table


def generate_sap(config, output_path="SAP.docx"):
    """
    Generate a SAP document from config dict.

    Parameters
    ----------
    config : dict
        Study configuration (see module docstring for schema).
    output_path : str
        Path for output .docx file.

    Returns
    -------
    dict with keys: output_path, sample_size_result, power_curve_path
    """
    doc = Document()

    study = config["study"]
    design = config["design"]
    objectives = config["objectives"]
    primary = config["primary_outcome"]
    analysis = config["analysis"]
    populations = config.get("populations", {})
    secondary_outcomes = config.get("secondary_outcomes", [])

    # --- Style defaults ---
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    # ========================================
    # COVER PAGE
    # ========================================
    doc.add_paragraph("")
    doc.add_paragraph("")
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("STATISTICAL ANALYSIS PLAN")
    run.bold = True
    run.font.size = Pt(20)

    doc.add_paragraph("")
    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle_p.add_run(study["title"])
    run.bold = True
    run.font.size = Pt(14)

    doc.add_paragraph("")
    doc.add_paragraph("")

    info_items = [
        ("Protocol Number", study.get("protocol_number", "—")),
        ("Study Phase", study.get("phase", "—")),
        ("Sponsor", study.get("sponsor", "—")),
        ("Principal Investigator", study.get("pi", "—")),
        ("SAP Version", study.get("sap_version", "1.0")),
        ("SAP Date", study.get("sap_date", datetime.now().strftime("%Y-%m-%d"))),
    ]
    for label, value in info_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{label}: ")
        run.bold = True
        run.font.size = Pt(11)
        run = p.add_run(value)
        run.font.size = Pt(11)

    doc.add_page_break()

    # ========================================
    # SAP APPROVAL SIGNATURE PAGE
    # ========================================
    _add_heading(doc, "SAP Approval", level=1)
    doc.add_paragraph(
        "This Statistical Analysis Plan has been reviewed and approved by "
        "the following individuals:"
    )
    _add_table(doc,
        ["Role", "Name", "Signature", "Date"],
        [
            ["Principal Investigator", study.get("pi", ""), "", ""],
            ["Biostatistician", "", "", ""],
            ["Co-Investigator", "", "", ""],
        ]
    )
    doc.add_paragraph("")

    # ========================================
    # REVISION HISTORY
    # ========================================
    _add_heading(doc, "SAP Revision History", level=1)
    _add_table(doc,
        ["Version", "Date", "Description of Changes", "Author"],
        [
            [study.get("sap_version", "1.0"),
             study.get("sap_date", ""),
             "Initial version", study.get("pi", "")],
        ]
    )
    doc.add_page_break()

    # ========================================
    # TABLE OF CONTENTS placeholder
    # ========================================
    _add_heading(doc, "Table of Contents", level=1)
    doc.add_paragraph("[Auto-generated — Insert TOC in Word after document review]")
    doc.add_page_break()

    # ========================================
    # SECTION 1: INTRODUCTION
    # ========================================
    _add_heading(doc, "1. Introduction", level=1)
    doc.add_paragraph(
        f"This document describes the statistical analysis plan for the study "
        f'"{study["title"]}". '
        f"This SAP was developed in accordance with ICH E9 (Statistical Principles "
        f"for Clinical Trials), ICH E9(R1) (Estimands and Sensitivity Analysis), "
        f"and the TransCelerate Common SAP Template structure."
    )
    doc.add_paragraph(
        "The SAP is intended to be finalized and signed before database lock. "
        "Any deviations from this plan will be documented in the clinical study report."
    )

    # ========================================
    # SECTION 2: STUDY OBJECTIVES AND ENDPOINTS
    # ========================================
    _add_heading(doc, "2. Study Objectives and Endpoints", level=1)

    _add_heading(doc, "2.1 Primary Objective", level=2)
    doc.add_paragraph(objectives["primary"]["description"])
    if "hypothesis" in objectives["primary"]:
        p = doc.add_paragraph()
        run = p.add_run("Hypothesis: ")
        run.bold = True
        p.add_run(objectives["primary"]["hypothesis"])

    _add_heading(doc, "2.2 Primary Endpoint", level=2)
    doc.add_paragraph(
        f"{primary['name']} ({primary['type']} outcome)"
    )

    if objectives.get("secondary"):
        _add_heading(doc, "2.3 Secondary Objectives and Endpoints", level=2)
        for i, obj in enumerate(objectives["secondary"], 1):
            doc.add_paragraph(f"{i}. {obj}", style="List Number")

    if secondary_outcomes:
        _add_heading(doc, "2.4 Secondary Endpoints", level=2)
        for so in secondary_outcomes:
            doc.add_paragraph(f"• {so['name']} ({so['type']})", style="List Bullet")

    if objectives.get("exploratory"):
        _add_heading(doc, "2.5 Exploratory Objectives", level=2)
        for obj in objectives["exploratory"]:
            doc.add_paragraph(f"• {obj}", style="List Bullet")

    # ========================================
    # SECTION 3: STUDY DESIGN
    # ========================================
    _add_heading(doc, "3. Study Design", level=1)

    design_map = {
        "rct_parallel": "Randomized, controlled, parallel-group trial",
        "rct_crossover": "Randomized, controlled, crossover trial",
        "cohort": "Prospective cohort study",
        "case_control": "Case-control study",
    }
    design_desc = design_map.get(design["type"], design["type"])
    blinding_map = {
        "open_label": "Open-label (no blinding)",
        "single_blind_assessor": "Single-blind (outcome assessor blinded)",
        "double_blind": "Double-blind (participant and assessor blinded)",
    }
    blinding_desc = blinding_map.get(design.get("blinding", ""), design.get("blinding", "—"))

    _add_table(doc,
        ["Parameter", "Description"],
        [
            ["Study Design", design_desc],
            ["Treatment Arms", ", ".join(a["name"] for a in design["arms"])],
            ["Allocation Ratio", ":".join(str(a.get("ratio", 1)) for a in design["arms"])],
            ["Blinding", blinding_desc],
        ]
    )

    # ========================================
    # SECTION 4: SAMPLE SIZE JUSTIFICATION
    # ========================================
    _add_heading(doc, "4. Sample Size Justification", level=1)

    ss = calculate_sample_size(primary)

    doc.add_paragraph(
        f"The sample size calculation is based on the primary endpoint: "
        f"{primary['name']}."
    )

    _add_heading(doc, "4.1 Assumptions and Calculation", level=2)
    _add_table(doc,
        ["Parameter", "Value"],
        [[k, str(v)] for k, v in ss["assumptions"].items()]
    )

    doc.add_paragraph("")
    p = doc.add_paragraph()
    run = p.add_run("Formula: ")
    run.bold = True
    p.add_run(ss["formula_used"])

    doc.add_paragraph("")
    result_items = [
        ["Sample size per group (unadjusted)", str(ss["n_per_group"])],
        ["Total sample size (unadjusted)", str(ss["n_total"])],
    ]
    if ss["dropout_rate"] > 0:
        result_items.extend([
            ["Anticipated dropout rate", f"{ss['dropout_rate']*100:.0f}%"],
            ["Sample size per group (adjusted)", str(ss["n_per_group_adjusted"])],
            ["Total sample size (adjusted)", str(ss["n_total_adjusted"])],
        ])
    _add_table(doc, ["Result", "Value"], result_items)

    # Power curve
    _add_heading(doc, "4.2 Power Curve", level=2)
    doc.add_paragraph(
        "The following figure shows statistical power as a function of sample size "
        "per group for various effect sizes around the assumed MCID."
    )
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        power_curve_path = tmp.name
    generate_power_curve(primary, power_curve_path)
    doc.add_picture(power_curve_path, width=Inches(5.5))
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Sensitivity table
    _add_heading(doc, "4.3 Sensitivity Analysis — Sample Size", level=2)
    doc.add_paragraph(
        "The table below shows the required sample size per group (adjusted for "
        f"{ss['dropout_rate']*100:.0f}% dropout) under various MCID and SD assumptions."
    )
    sens_rows, mcid_range, sd_range = generate_sensitivity_table(primary)
    headers = ["MCID"] + [f"SD={sd}" for sd in sd_range]
    table_rows = []
    for row in sens_rows:
        table_rows.append([str(row["MCID"])] + [str(row[f"SD={sd}"]) for sd in sd_range])
    _add_table(doc, headers, table_rows)

    base_mcid = primary["mcid"]
    base_sd = primary["sd"]
    doc.add_paragraph(
        f"Note: The primary analysis assumes MCID = {base_mcid} and SD = {base_sd} "
        f"(highlighted cell).",
        style="List Bullet"
    )

    # ========================================
    # SECTION 5: RANDOMIZATION AND BLINDING
    # ========================================
    _add_heading(doc, "5. Randomization, Stratification, and Blinding", level=1)

    rand = design.get("randomization", {})
    rand_method_map = {
        "block": "Permuted block randomization",
        "simple": "Simple randomization",
        "stratified_block": "Stratified block randomization",
        "minimization": "Minimization",
    }
    rand_desc = rand_method_map.get(rand.get("method", ""), rand.get("method", "—"))

    items = [["Method", rand_desc]]
    if rand.get("block_sizes"):
        items.append(["Block sizes", str(rand["block_sizes"])])
    if rand.get("stratification"):
        items.append(["Stratification factors", ", ".join(rand["stratification"])])
    items.append(["Blinding", blinding_desc])

    _add_table(doc, ["Parameter", "Description"], items)

    # ========================================
    # SECTION 6: ANALYSIS POPULATIONS
    # ========================================
    _add_heading(doc, "6. Analysis Populations", level=1)

    pop_items = []
    pop_map = {
        "itt": "Intent-to-Treat (ITT) / Full Analysis Set",
        "mitt": "Modified ITT (mITT)",
        "pp": "Per-Protocol (PP)",
        "safety": "Safety Population",
    }
    for key, label in pop_map.items():
        if key in populations:
            pop_items.append([label, populations[key]])

    if pop_items:
        _add_table(doc, ["Population", "Definition"], pop_items)
    else:
        doc.add_paragraph("[Define analysis populations]")

    # ========================================
    # SECTION 7: STATISTICAL ANALYSES
    # ========================================
    _add_heading(doc, "7. Statistical Analyses", level=1)

    # 7.1 General
    _add_heading(doc, "7.1 General Statistical Considerations", level=2)
    general_items = [
        ["Statistical software", analysis.get("software", "Python / R")],
        ["Significance level", analysis.get("significance_level", "Two-sided α = 0.05")],
        ["Confidence intervals", "95% CI for all point estimates"],
        ["Missing data", analysis.get("missing_data", "—")],
        ["Multiplicity adjustment", analysis.get("multiplicity", "—")],
    ]
    _add_table(doc, ["Parameter", "Description"], general_items)

    # 7.2 Demographics
    _add_heading(doc, "7.2 Demographics and Baseline Characteristics", level=2)
    doc.add_paragraph(
        "Baseline demographics and clinical characteristics will be summarized by "
        "treatment group using descriptive statistics:"
    )
    doc.add_paragraph(
        "• Continuous variables: mean ± SD (or median [IQR] if non-normal)",
        style="List Bullet"
    )
    doc.add_paragraph(
        "• Categorical variables: frequency and percentage, n (%)",
        style="List Bullet"
    )
    doc.add_paragraph(
        "Baseline comparisons between groups will be presented using independent "
        "t-tests (or Mann-Whitney U) for continuous variables and chi-square "
        "(or Fisher's exact) for categorical variables. "
        "Results will be presented as Table 1."
    )

    # 7.3 Primary Efficacy
    _add_heading(doc, "7.3 Primary Efficacy Analysis", level=2)
    p = doc.add_paragraph()
    run = p.add_run("Primary endpoint: ")
    run.bold = True
    p.add_run(primary["name"])

    doc.add_paragraph("")
    p = doc.add_paragraph()
    run = p.add_run("Analysis method: ")
    run.bold = True
    p.add_run(analysis.get("primary_method", "—"))

    doc.add_paragraph("")
    doc.add_paragraph(
        "The primary analysis will be performed on the ITT population. "
        "Results will be reported as the between-group difference with "
        "95% confidence interval and p-value."
    )

    # 7.4 Secondary
    _add_heading(doc, "7.4 Secondary Efficacy Analyses", level=2)
    if secondary_outcomes:
        for so in secondary_outcomes:
            doc.add_paragraph(f"• {so['name']} ({so['type']})", style="List Bullet")
        doc.add_paragraph("")
        doc.add_paragraph(
            "Secondary endpoints will be analyzed using the same methods as the "
            "primary analysis. P-values will be adjusted for multiplicity using "
            f"{analysis.get('multiplicity', 'appropriate methods')}."
        )
    else:
        doc.add_paragraph("[Define secondary analyses]")

    # 7.5 Sensitivity
    _add_heading(doc, "7.5 Sensitivity Analyses", level=2)
    if analysis.get("sensitivity"):
        for i, s in enumerate(analysis["sensitivity"], 1):
            doc.add_paragraph(f"{i}. {s}", style="List Number")
    else:
        doc.add_paragraph("[Define sensitivity analyses]")

    # 7.6 Subgroup
    _add_heading(doc, "7.6 Subgroup Analyses", level=2)
    if analysis.get("subgroup"):
        doc.add_paragraph(
            "Pre-specified subgroup analyses will be performed for the primary "
            "endpoint stratified by the following variables:"
        )
        for sg in analysis["subgroup"]:
            doc.add_paragraph(f"• {sg}", style="List Bullet")
        doc.add_paragraph("")
        doc.add_paragraph(
            "Subgroup effects will be assessed using interaction tests within "
            "the primary analysis model. Forest plots will be generated for "
            "visual presentation."
        )
    else:
        doc.add_paragraph("No pre-specified subgroup analyses planned.")

    # 7.7 Safety
    _add_heading(doc, "7.7 Safety Analyses", level=2)
    doc.add_paragraph(
        "Safety analyses will be performed on the safety population. "
        "Adverse events will be summarized by treatment group using "
        "frequency counts and percentages. Serious adverse events (SAEs) "
        "will be listed individually."
    )

    # ========================================
    # SECTION 8: CHANGES FROM PROTOCOL
    # ========================================
    _add_heading(doc, "8. Changes from Protocol-Specified Analyses", level=1)
    doc.add_paragraph("No changes from the protocol-specified analyses at this time.")

    # ========================================
    # SECTION 9: CONVENTIONS
    # ========================================
    _add_heading(doc, "9. Conventions", level=1)
    conventions = [
        ["Rounding", "Continuous variables: 1 decimal place. P-values: 3 decimal places (report p < 0.001 if applicable)"],
        ["Date imputation", "Partial dates will be imputed using the 15th of the month if day is missing"],
        ["Age calculation", "Age = (Informed consent date − Date of birth) / 365.25"],
        ["Study day", "Study Day 1 = date of intervention; negative days for pre-intervention visits"],
        ["BMI", "BMI = weight(kg) / height(m)²"],
    ]
    _add_table(doc, ["Convention", "Rule"], conventions)

    # ========================================
    # APPENDIX A: TFL SHELLS
    # ========================================
    doc.add_page_break()
    _add_heading(doc, "Appendix A: Tables, Figures, and Listings (TFL) Shells", level=1)
    doc.add_paragraph(
        "The following mock-up tables will be populated after database lock:"
    )
    doc.add_paragraph("• Table 1: Baseline demographics and clinical characteristics", style="List Bullet")
    doc.add_paragraph("• Table 2: Primary endpoint results by treatment group", style="List Bullet")
    doc.add_paragraph("• Table 3: Secondary endpoint results", style="List Bullet")
    doc.add_paragraph("• Table 4: Adverse events summary", style="List Bullet")
    doc.add_paragraph("• Figure 1: CONSORT flow diagram", style="List Bullet")
    doc.add_paragraph("• Figure 2: Primary endpoint — between-group comparison", style="List Bullet")
    doc.add_paragraph("• Figure 3: Subgroup analysis forest plot", style="List Bullet")

    # ========================================
    # SAVE
    # ========================================
    doc.save(output_path)

    # Cleanup temp power curve
    try:
        os.unlink(power_curve_path)
    except Exception:
        pass

    return {
        "output_path": output_path,
        "sample_size": ss,
        "power_curve_generated": True,
    }


# ============================================================
# Example Config (FSHD Frozen Shoulder Pilot)
# ============================================================

EXAMPLE_CONFIG = {
    "study": {
        "title": "Comparison of Anterior vs Posterior Approach Ultrasound-guided "
                 "Hydrodilatation for Frozen Shoulder: A Randomized Controlled Trial",
        "protocol_number": "PNUH-2026-IRB-XXX",
        "phase": "Investigator-initiated clinical trial",
        "sponsor": "Department of Rehabilitation Medicine, Pusan National University Hospital",
        "pi": "Jaehyun Lee, MD",
        "sap_version": "1.0",
        "sap_date": "2026-04-01",
    },
    "design": {
        "type": "rct_parallel",
        "arms": [
            {"name": "Anterior approach hydrodilatation", "ratio": 1},
            {"name": "Posterior approach hydrodilatation", "ratio": 1},
        ],
        "blinding": "single_blind_assessor",
        "randomization": {
            "method": "block",
            "block_sizes": [4, 6],
            "stratification": ["sex"],
        },
    },
    "objectives": {
        "primary": {
            "description": "To compare the effectiveness of anterior vs posterior approach "
                           "ultrasound-guided hydrodilatation in patients with frozen shoulder, "
                           "as measured by SPADI total score change from baseline to 3 months.",
            "hypothesis": "The anterior approach is non-inferior to the posterior approach "
                          "in improving SPADI total score at 3 months post-intervention.",
        },
        "secondary": [
            "Compare shoulder ROM (ER, IR, ABD, flexion) change at 3 months",
            "Compare pain NRS (at rest and during activity) change at 3 months",
            "Compare SST score change at 3 months",
            "Compare PGA (patient global assessment) at 3 months",
        ],
        "exploratory": [
            "Ultrasound findings (SST integrity, biceps effusion, capsular thickness) at 3 months",
            "Reinjection rate within 3 months",
            "Time to meaningful clinical improvement (SPADI MCID = 13)",
        ],
    },
    "primary_outcome": {
        "name": "SPADI total score change from baseline to 3 months",
        "type": "continuous",
        "test": "independent_t",
        "mcid": 13,
        "sd": 20,
        "alpha": 0.05,
        "power": 0.80,
        "sides": 2,
        "dropout_rate": 0.20,
    },
    "secondary_outcomes": [
        {"name": "Shoulder ER ROM (degrees)", "type": "continuous"},
        {"name": "Shoulder IR ROM (degrees)", "type": "continuous"},
        {"name": "Shoulder ABD ROM (degrees)", "type": "continuous"},
        {"name": "Shoulder Flexion ROM (degrees)", "type": "continuous"},
        {"name": "Pain NRS at rest (0-10)", "type": "continuous"},
        {"name": "Pain NRS during activity (0-10)", "type": "continuous"},
        {"name": "Night pain NRS (0-10)", "type": "continuous"},
        {"name": "SST score (0-12)", "type": "continuous"},
        {"name": "PGA (patient global assessment, 0-10)", "type": "continuous"},
    ],
    "populations": {
        "itt": "All randomized subjects, analyzed according to assigned treatment group, "
               "regardless of actual treatment received or protocol deviations.",
        "pp": "All randomized subjects who completed the 3-month follow-up visit "
              "without major protocol deviations (e.g., wrong approach, reinjection "
              "before 3 months, missing primary endpoint).",
        "safety": "All subjects who received at least one injection procedure.",
    },
    "analysis": {
        "primary_method": "Independent samples t-test for the primary comparison. "
                          "ANCOVA adjusting for baseline SPADI score as a pre-specified "
                          "confirmatory analysis.",
        "missing_data": "Multiple imputation using chained equations (MICE) under "
                        "missing-at-random (MAR) assumption. Complete-case analysis "
                        "as sensitivity.",
        "multiplicity": "Bonferroni correction applied to secondary endpoints. "
                        "Exploratory analyses reported with nominal p-values.",
        "sensitivity": [
            "Per-protocol analysis (PP population)",
            "ANCOVA with baseline SPADI as covariate (ITT population)",
            "Complete-case analysis (without imputation)",
            "Worst-case imputation (last observation carried backward for dropouts)",
        ],
        "subgroup": [
            "Sex (male vs female)",
            "Diabetes mellitus (yes vs no)",
            "Dominant hand affected (yes vs no)",
            "Symptom duration (3-6 months vs >6 months)",
        ],
        "software": "Python 3.x (scipy, statsmodels, pandas), R 4.x (as needed for specific analyses)",
        "significance_level": "Two-sided α = 0.05 for the primary endpoint",
    },
}


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="SAP Generator — Statistical Analysis Plan from config"
    )
    parser.add_argument(
        "config_file", nargs="?",
        help="Path to config JSON file"
    )
    parser.add_argument(
        "--output", "-o", default="SAP.docx",
        help="Output docx path (default: SAP.docx)"
    )
    parser.add_argument(
        "--example", action="store_true",
        help="Print example config JSON to stdout"
    )
    parser.add_argument(
        "--run", action="store_true",
        help="Run with example config (use with --example)"
    )
    args = parser.parse_args()

    if args.example:
        if args.run:
            result = generate_sap(EXAMPLE_CONFIG, output_path=args.output)
            print(f"SAP generated: {result['output_path']}")
            ss = result["sample_size"]
            print(f"Sample size: {ss['n_per_group']} per group "
                  f"(adjusted for {ss['dropout_rate']*100:.0f}% dropout: "
                  f"{ss['n_per_group_adjusted']} per group, "
                  f"{ss['n_total_adjusted']} total)")
        else:
            print(json.dumps(EXAMPLE_CONFIG, indent=2, ensure_ascii=False))
        return

    if not args.config_file:
        parser.print_help()
        return

    with open(args.config_file) as f:
        if args.config_file.endswith((".yaml", ".yml")):
            import yaml
            config = yaml.safe_load(f)
        else:
            config = json.load(f)

    result = generate_sap(config, output_path=args.output)
    print(f"SAP generated: {result['output_path']}")
    ss = result["sample_size"]
    print(f"Sample size: {ss['n_per_group']} per group "
          f"(adjusted for {ss['dropout_rate']*100:.0f}% dropout: "
          f"{ss['n_per_group_adjusted']} per group, "
          f"{ss['n_total_adjusted']} total)")


if __name__ == "__main__":
    main()

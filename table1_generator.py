#!/usr/bin/env python3
"""
table1_generator.py — Table 1 (Baseline Characteristics) + Analysis Script Skeleton

Phase 5 of the clinical research workflow:
  1. Table 1 Generator: CSV/Excel → baseline characteristics table with group
     comparisons (mean±SD, n(%), p-values)
  2. Analysis Script Skeleton: SAP config → numbered Python analysis script files

=== Usage ===

Python:
    from table1_generator import (
        generate_table1, generate_analysis_skeleton,
        EXAMPLE_TABLE1_CONFIG
    )

    # Table 1 from CSV data
    result = generate_table1("data.csv", EXAMPLE_TABLE1_CONFIG)

    # Analysis skeleton from SAP config
    generate_analysis_skeleton(sap_config, output_dir="./analysis")

CLI:
    python3 table1_generator.py table1 data.csv --config table1_config.json -o table1.docx
    python3 table1_generator.py skeleton --sap-config sap.json -o ./analysis
    python3 table1_generator.py --example

Dependencies: pandas, scipy, python-docx
"""

import os
import sys
import json
import argparse
import math
from datetime import datetime
from collections import OrderedDict

try:
    import pandas as pd
    import numpy as np
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from scipy import stats
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ============================================================
# Table 1 Generator
# ============================================================

def generate_table1(data_path, config, output_path=None):
    """Generate Table 1 (Baseline Characteristics) from CSV/Excel data.

    Args:
        data_path: path to CSV or Excel file
        config: dict with:
            group_col: column name for group variable
            variables: list of {name, type, label}
                type: 'continuous' (mean±SD), 'continuous_skewed' (median[IQR]),
                      'categorical' (n(%)), 'binary' (n(%))
            title: optional table title
        output_path: if ends with .docx → Word, .md → Markdown, .csv → CSV

    Returns:
        dict: {table: list of rows, stats: list of test results, markdown: str}
    """
    if not HAS_PANDAS:
        raise ImportError("pandas required. Install: pip install pandas")
    if not HAS_SCIPY:
        raise ImportError("scipy required. Install: pip install scipy")

    # Load data
    if data_path.endswith(".xlsx") or data_path.endswith(".xls"):
        df = pd.read_excel(data_path)
    else:
        df = pd.read_csv(data_path)

    group_col = config["group_col"]
    variables = config["variables"]

    # Identify groups
    groups = sorted(df[group_col].dropna().unique())
    n_groups = len(groups)

    # Build table rows
    rows = []  # Each row: (label, overall, *group_values, p_value, test_name)

    # Header row
    header = ["Variable", f"Overall (N={len(df)})"]
    for g in groups:
        n_g = len(df[df[group_col] == g])
        header.append(f"{g} (n={n_g})")
    header.extend(["p-value", "Test"])
    rows.append(header)

    for var in variables:
        col_name = var["name"]
        var_type = var.get("type", "continuous")
        label = var.get("label", col_name)

        if col_name not in df.columns:
            rows.append([label, "—"] + ["—"] * n_groups + ["—", "—"])
            continue

        if var_type == "continuous":
            row = _continuous_row(df, col_name, group_col, groups, label, use_median=False)
        elif var_type == "continuous_skewed":
            row = _continuous_row(df, col_name, group_col, groups, label, use_median=True)
        elif var_type in ("categorical", "binary"):
            cat_rows = _categorical_rows(df, col_name, group_col, groups, label)
            rows.extend(cat_rows)
            continue
        else:
            row = _continuous_row(df, col_name, group_col, groups, label, use_median=False)

        rows.append(row)

    # Generate markdown
    markdown = _table_to_markdown(rows, config.get("title", "Table 1. Baseline Characteristics"))

    # Save outputs
    if output_path:
        if output_path.endswith(".docx"):
            _table_to_docx(rows, config, output_path)
        elif output_path.endswith(".md"):
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(markdown)
        elif output_path.endswith(".csv"):
            _table_to_csv(rows, output_path)

    return {
        "table": rows,
        "markdown": markdown,
        "n_variables": len(variables),
        "n_groups": n_groups,
        "groups": list(groups),
    }


def _continuous_row(df, col, group_col, groups, label, use_median=False):
    """Generate a row for continuous variable."""
    series = df[col].dropna()

    if use_median:
        overall = f"{series.median():.1f} [{series.quantile(0.25):.1f}–{series.quantile(0.75):.1f}]"
    else:
        overall = f"{series.mean():.1f} ± {series.std():.1f}"

    group_vals = []
    group_data = []
    for g in groups:
        gs = df.loc[df[group_col] == g, col].dropna()
        group_data.append(gs)
        if use_median:
            group_vals.append(f"{gs.median():.1f} [{gs.quantile(0.25):.1f}–{gs.quantile(0.75):.1f}]")
        else:
            group_vals.append(f"{gs.mean():.1f} ± {gs.std():.1f}")

    # Statistical test
    if len(groups) == 2:
        if use_median:
            stat, p = stats.mannwhitneyu(group_data[0], group_data[1], alternative="two-sided")
            test = "Mann-Whitney"
        else:
            stat, p = stats.ttest_ind(group_data[0], group_data[1])
            test = "t-test"
    elif len(groups) > 2:
        if use_median:
            stat, p = stats.kruskal(*group_data)
            test = "Kruskal-Wallis"
        else:
            stat, p = stats.f_oneway(*group_data)
            test = "ANOVA"
    else:
        p = None
        test = "—"

    p_str = f"{p:.3f}" if p is not None else "—"
    if p is not None and p < 0.001:
        p_str = "<0.001"

    return [label, overall] + group_vals + [p_str, test]


def _categorical_rows(df, col, group_col, groups, label):
    """Generate rows for categorical variable (header + categories)."""
    rows = []
    categories = sorted(df[col].dropna().unique())

    # Header row
    p_val, test = _categorical_test(df, col, group_col)
    p_str = f"{p_val:.3f}" if p_val is not None else "—"
    if p_val is not None and p_val < 0.001:
        p_str = "<0.001"

    header_row = [f"{label}, n (%)", "", ""] + [""] * (len(groups) - 1) + [p_str, test]
    rows.append(header_row)

    # Category rows
    for cat in categories:
        overall_n = (df[col] == cat).sum()
        overall_pct = overall_n / len(df) * 100
        overall_str = f"{overall_n} ({overall_pct:.1f}%)"

        group_vals = []
        for g in groups:
            g_df = df[df[group_col] == g]
            n = (g_df[col] == cat).sum()
            pct = n / len(g_df) * 100 if len(g_df) > 0 else 0
            group_vals.append(f"{n} ({pct:.1f}%)")

        rows.append([f"  {cat}", overall_str] + group_vals + ["", ""])

    return rows


def _categorical_test(df, col, group_col):
    """Chi-square or Fisher's exact test for categorical variable."""
    try:
        ct = pd.crosstab(df[col], df[group_col])
        if ct.min().min() < 5:
            # Fisher's exact for 2x2, else chi-square
            if ct.shape == (2, 2):
                _, p = stats.fisher_exact(ct)
                return p, "Fisher"
        chi2, p, dof, expected = stats.chi2_contingency(ct)
        return p, "Chi-square"
    except Exception:
        return None, "—"


# ============================================================
# Table output formats
# ============================================================

def _table_to_markdown(rows, title="Table 1"):
    """Convert table rows to markdown."""
    lines = [f"### {title}", ""]
    if not rows:
        return "\n".join(lines)

    # Calculate column widths
    n_cols = len(rows[0])
    widths = [max(len(str(row[i])) if i < len(row) else 0 for row in rows) for i in range(n_cols)]

    # Header
    header = rows[0]
    lines.append("| " + " | ".join(str(h).ljust(w) for h, w in zip(header, widths)) + " |")
    lines.append("| " + " | ".join("-" * w for w in widths) + " |")

    # Data rows
    for row in rows[1:]:
        padded = [str(row[i]).ljust(widths[i]) if i < len(row) else " " * widths[i]
                  for i in range(n_cols)]
        lines.append("| " + " | ".join(padded) + " |")

    lines.append("")
    return "\n".join(lines)


def _table_to_docx(rows, config, output_path):
    """Export Table 1 to docx."""
    if not HAS_DOCX:
        raise ImportError("python-docx required.")

    doc = Document()
    title = config.get("title", "Table 1. Baseline Characteristics")
    doc.add_heading(title, level=2)

    if not rows:
        doc.save(output_path)
        return

    n_cols = len(rows[0])
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = "Table Grid"

    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            cell = table.rows[i].cells[j]
            run = cell.paragraphs[0].add_run(str(cell_text))
            run.font.size = Pt(8)
            if i == 0:
                run.bold = True

    doc.save(output_path)


def _table_to_csv(rows, output_path):
    """Export Table 1 to CSV."""
    import csv
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(row)


# ============================================================
# Analysis Script Skeleton Generator
# ============================================================

def generate_analysis_skeleton(sap_config, output_dir="./analysis"):
    """Generate numbered Python analysis script skeleton from SAP config.

    Creates a directory with analysis scripts that follow the SAP structure:
      01_data_import.py
      02_data_cleaning.py
      03_descriptive.py (Table 1)
      04_primary_analysis.py
      05_secondary_analysis.py
      06_sensitivity_analysis.py
      07_figures.py
      08_export_results.py

    Returns list of created file paths.
    """
    os.makedirs(output_dir, exist_ok=True)
    study = sap_config.get("study", {})
    study_title = study.get("title", "Clinical Study")

    scripts = []

    # 01: Data Import
    scripts.append(("01_data_import.py", f'''#!/usr/bin/env python3
"""01_data_import.py — Data Import
Study: {study_title}
Generated: {datetime.now().strftime("%Y-%m-%d")} by table1_generator.py
"""
import pandas as pd

# --- Configuration ---
DATA_PATH = "raw_data.csv"  # Or .xlsx
ENCODING = "utf-8"  # Or "cp949" for Korean Excel

def load_data(path=DATA_PATH):
    """Load raw data from file."""
    if path.endswith(".xlsx") or path.endswith(".xls"):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path, encoding=ENCODING)
    print(f"Loaded: {{df.shape[0]}} rows x {{df.shape[1]}} columns")
    print(f"Columns: {{list(df.columns)}}")
    return df

if __name__ == "__main__":
    df = load_data()
    print(df.head())
    print(df.describe())
'''))

    # 02: Data Cleaning
    po = sap_config.get("primary_outcome", {})
    dropout = po.get("dropout_rate", 0.2)
    scripts.append(("02_data_cleaning.py", f'''#!/usr/bin/env python3
"""02_data_cleaning.py — Data Cleaning & Validation
Study: {study_title}
"""
import pandas as pd
import numpy as np

def clean_data(df):
    """Clean and validate raw data."""
    n_original = len(df)

    # 1. Remove exact duplicates
    df = df.drop_duplicates()
    print(f"Duplicates removed: {{n_original - len(df)}}")

    # 2. Check for expected columns
    # TODO: Define required columns based on CRF
    required_cols = []  # e.g., ["subject_id", "group", "spadi_baseline", "spadi_3mo"]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        print(f"WARNING: Missing columns: {{missing_cols}}")

    # 3. Type validation
    # TODO: Validate data types (numeric, categorical, dates)

    # 4. Range checks
    # TODO: Add range checks based on CRF expected values

    # 5. Missing data summary
    missing = df.isnull().sum()
    missing_pct = (missing / len(df) * 100).round(1)
    missing_report = pd.DataFrame({{"n_missing": missing, "pct": missing_pct}})
    missing_report = missing_report[missing_report["n_missing"] > 0]
    if len(missing_report) > 0:
        print("\\nMissing data:")
        print(missing_report)
    else:
        print("No missing data found.")

    # 6. Expected dropout rate: {dropout*100:.0f}%
    # Verify actual vs expected

    return df

if __name__ == "__main__":
    from data_01_import import load_data
    df = load_data()
    df_clean = clean_data(df)
    df_clean.to_csv("cleaned_data.csv", index=False)
    print(f"\\nCleaned data: {{df_clean.shape}}")
'''))

    # 03: Descriptive (Table 1)
    scripts.append(("03_descriptive.py", f'''#!/usr/bin/env python3
"""03_descriptive.py — Descriptive Statistics (Table 1)
Study: {study_title}
"""
import pandas as pd
import sys
sys.path.insert(0, "{os.path.abspath(os.path.dirname(__file__) + '/..')}")

# Import from pytool if available
try:
    from table1_generator import generate_table1
    HAS_TABLE1 = True
except ImportError:
    HAS_TABLE1 = False

# --- Table 1 Configuration ---
TABLE1_CONFIG = {{
    "group_col": "group",  # Column name for treatment group
    "title": "Table 1. Baseline Characteristics",
    "variables": [
        # TODO: Update based on your CRF variables
        {{"name": "age", "type": "continuous", "label": "Age (years)"}},
        {{"name": "sex", "type": "categorical", "label": "Sex"}},
        {{"name": "bmi", "type": "continuous", "label": "BMI (kg/m²)"}},
        # Primary outcome baseline
        {{"name": "spadi_baseline", "type": "continuous", "label": "SPADI total (%)"}},
        # Add more variables...
    ],
}}

def run_table1(data_path="cleaned_data.csv", output="table1.docx"):
    if HAS_TABLE1:
        result = generate_table1(data_path, TABLE1_CONFIG, output)
        print(result["markdown"])
        return result
    else:
        print("table1_generator not available. Using basic pandas describe().")
        df = pd.read_csv(data_path)
        print(df.groupby(TABLE1_CONFIG["group_col"]).describe())

if __name__ == "__main__":
    run_table1()
'''))

    # 04: Primary Analysis
    primary = sap_config.get("primary_outcome", {})
    method = sap_config.get("analysis", {}).get("primary_method", "t-test")
    scripts.append(("04_primary_analysis.py", f'''#!/usr/bin/env python3
"""04_primary_analysis.py — Primary Outcome Analysis
Study: {study_title}
Primary outcome: {primary.get("name", "TBD")}
Method: {method}
"""
import pandas as pd
import numpy as np
from scipy import stats

def primary_analysis(df, group_col="group", outcome_col="primary_change"):
    """Primary analysis: {method}"""
    groups = sorted(df[group_col].unique())
    assert len(groups) == 2, f"Expected 2 groups, got {{len(groups)}}"

    g1 = df.loc[df[group_col] == groups[0], outcome_col].dropna()
    g2 = df.loc[df[group_col] == groups[1], outcome_col].dropna()

    print(f"Group {{groups[0]}}: n={{len(g1)}}, mean={{g1.mean():.2f}} ± {{g1.std():.2f}}")
    print(f"Group {{groups[1]}}: n={{len(g2)}}, mean={{g2.mean():.2f}} ± {{g2.std():.2f}}")

    # Independent t-test
    t_stat, p_value = stats.ttest_ind(g1, g2)
    mean_diff = g1.mean() - g2.mean()
    se_diff = np.sqrt(g1.var()/len(g1) + g2.var()/len(g2))
    ci_lower = mean_diff - 1.96 * se_diff
    ci_upper = mean_diff + 1.96 * se_diff

    print(f"\\nMean difference: {{mean_diff:.2f}} (95% CI: {{ci_lower:.2f}} to {{ci_upper:.2f}})")
    print(f"t-statistic: {{t_stat:.3f}}, p-value: {{p_value:.4f}}")

    # MCID = {primary.get("mcid", "TBD")}
    # Clinically meaningful?

    return {{
        "mean_diff": mean_diff,
        "ci": (ci_lower, ci_upper),
        "t_stat": t_stat,
        "p_value": p_value,
    }}

if __name__ == "__main__":
    df = pd.read_csv("cleaned_data.csv")
    result = primary_analysis(df)
'''))

    # 05: Secondary Analysis
    secondary = sap_config.get("secondary_outcomes", [])
    sec_names = [s.get("name", s) if isinstance(s, dict) else s for s in secondary]
    scripts.append(("05_secondary_analysis.py", f'''#!/usr/bin/env python3
"""05_secondary_analysis.py — Secondary Outcome Analyses
Study: {study_title}
Secondary outcomes: {sec_names}
Multiplicity: {sap_config.get("analysis", {}).get("multiplicity", "Bonferroni")}
"""
import pandas as pd
from scipy import stats

SECONDARY_OUTCOMES = {json.dumps(sec_names, indent=4, ensure_ascii=False)}

def secondary_analyses(df, group_col="group"):
    """Analyze secondary outcomes with multiplicity correction."""
    n_tests = len(SECONDARY_OUTCOMES)
    alpha_adj = 0.05 / n_tests  # Bonferroni

    results = []
    for outcome in SECONDARY_OUTCOMES:
        col = outcome.lower().replace(" ", "_")
        if col not in df.columns:
            print(f"WARNING: Column '{{col}}' not found, skipping {{outcome}}")
            continue
        # TODO: Implement per-outcome analysis
        print(f"Analyzing: {{outcome}} (adjusted α = {{alpha_adj:.4f}})")

    return results

if __name__ == "__main__":
    df = pd.read_csv("cleaned_data.csv")
    secondary_analyses(df)
'''))

    # 06: Sensitivity Analysis
    sensitivity = sap_config.get("analysis", {}).get("sensitivity", [])
    subgroups = sap_config.get("analysis", {}).get("subgroup", [])
    scripts.append(("06_sensitivity_analysis.py", f'''#!/usr/bin/env python3
"""06_sensitivity_analysis.py — Sensitivity & Subgroup Analyses
Study: {study_title}
Sensitivity: {sensitivity}
Subgroups: {subgroups}
"""
import pandas as pd
from scipy import stats

def per_protocol_analysis(df, group_col="group", pp_col="per_protocol"):
    """Per-protocol analysis (PP population)."""
    df_pp = df[df[pp_col] == True]
    print(f"Per-protocol population: {{len(df_pp)}}/{{len(df)}}")
    # TODO: Repeat primary analysis on PP population

def subgroup_analysis(df, group_col="group", subgroup_col=None, outcome_col="primary_change"):
    """Subgroup analysis with interaction test."""
    if subgroup_col is None or subgroup_col not in df.columns:
        return
    subgroups = df[subgroup_col].unique()
    for sg in subgroups:
        df_sub = df[df[subgroup_col] == sg]
        print(f"\\nSubgroup: {{subgroup_col}} = {{sg}} (n={{len(df_sub)}})")
        # TODO: Repeat primary analysis per subgroup

    # Interaction test
    # TODO: ANCOVA with group*subgroup interaction term

if __name__ == "__main__":
    df = pd.read_csv("cleaned_data.csv")
    # per_protocol_analysis(df)
    # subgroup_analysis(df, subgroup_col="sex")
'''))

    # 07: Figures
    scripts.append(("07_figures.py", f'''#!/usr/bin/env python3
"""07_figures.py — Publication-quality Figures
Study: {study_title}
"""
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

def figure_primary_outcome(df, group_col="group", outcome_col="primary_change", output="fig_primary.png"):
    """Box plot of primary outcome by group."""
    fig, ax = plt.subplots(figsize=(6, 5))
    groups = sorted(df[group_col].unique())
    data = [df.loc[df[group_col] == g, outcome_col].dropna() for g in groups]
    bp = ax.boxplot(data, labels=groups, patch_artist=True)
    colors = ["#4E79A7", "#E15759"]
    for patch, color in zip(bp["boxes"], colors[:len(groups)]):
        patch.set_facecolor(color)
        patch.set_alpha(0.7)
    ax.set_ylabel("Change from baseline")
    ax.set_title("Primary Outcome")
    plt.tight_layout()
    plt.savefig(output, dpi=300)
    plt.close()
    print(f"Saved: {{output}}")

def figure_consort_flow(counts, output="fig_consort.png"):
    """CONSORT flow diagram placeholder."""
    # TODO: Implement CONSORT flow diagram
    pass

if __name__ == "__main__":
    df = pd.read_csv("cleaned_data.csv")
    # figure_primary_outcome(df)
'''))

    # 08: Export Results
    scripts.append(("08_export_results.py", f'''#!/usr/bin/env python3
"""08_export_results.py — Export Results to docx/xlsx
Study: {study_title}
"""
import pandas as pd
import json

def export_all_results(results_dir="./results", output="study_results.xlsx"):
    """Compile all analysis results into a single Excel file."""
    # TODO: Collect results from each analysis script
    # Each script should save its results as JSON

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Table 1
        # Primary analysis
        # Secondary analysis
        # Sensitivity analysis
        pass

    print(f"Results exported: {{output}}")

if __name__ == "__main__":
    export_all_results()
'''))

    # Write all scripts
    created = []
    for filename, content in scripts:
        path = os.path.join(output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        created.append(path)

    # Write README
    readme_path = os.path.join(output_dir, "README.md")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(f"# Analysis Scripts: {study_title}\n\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d')} by table1_generator.py\n\n")
        f.write("## Execution Order\n\n")
        f.write("```bash\n")
        for filename, _ in scripts:
            f.write(f"python3 {filename}\n")
        f.write("```\n\n")
        f.write("## Script Descriptions\n\n")
        for filename, content in scripts:
            # Extract docstring first line
            desc = content.split('"""')[1].split("\n")[0] if '"""' in content else filename
            f.write(f"- `{filename}` — {desc}\n")
    created.append(readme_path)

    return created


# ============================================================
# Example Config
# ============================================================

EXAMPLE_TABLE1_CONFIG = {
    "group_col": "group",
    "title": "Table 1. Baseline Characteristics of Study Participants",
    "variables": [
        {"name": "age", "type": "continuous", "label": "Age (years)"},
        {"name": "sex", "type": "categorical", "label": "Sex"},
        {"name": "bmi", "type": "continuous", "label": "BMI (kg/m²)"},
        {"name": "affected_side", "type": "categorical", "label": "Affected side"},
        {"name": "dominant_hand", "type": "categorical", "label": "Dominant hand"},
        {"name": "symptom_duration", "type": "continuous_skewed", "label": "Symptom duration (months)"},
        {"name": "diabetes", "type": "binary", "label": "Diabetes mellitus"},
        {"name": "phase", "type": "categorical", "label": "Frozen shoulder phase"},
        {"name": "previous_injection", "type": "continuous_skewed", "label": "Previous injections (n)"},
        {"name": "nrs_rest", "type": "continuous", "label": "Pain NRS, rest (/10)"},
        {"name": "nrs_activity", "type": "continuous", "label": "Pain NRS, activity (/10)"},
        {"name": "nrs_night", "type": "continuous", "label": "Pain NRS, night (/10)"},
        {"name": "rom_flexion", "type": "continuous", "label": "ROM flexion (°)"},
        {"name": "rom_abd", "type": "continuous", "label": "ROM abduction (°)"},
        {"name": "rom_er", "type": "continuous", "label": "ROM external rotation (°)"},
        {"name": "spadi_total", "type": "continuous", "label": "SPADI total (%)"},
        {"name": "pga", "type": "continuous", "label": "PGA (/10)"},
    ],
}


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Table 1 Generator + Analysis Script Skeleton"
    )
    subparsers = parser.add_subparsers(dest="command")

    # table1 command
    t1 = subparsers.add_parser("table1", help="Generate Table 1 from data")
    t1.add_argument("data", help="CSV or Excel data file")
    t1.add_argument("--config", "-c", help="Table 1 config JSON")
    t1.add_argument("--output", "-o", default="table1.docx", help="Output path")

    # skeleton command
    sk = subparsers.add_parser("skeleton", help="Generate analysis script skeleton")
    sk.add_argument("--sap-config", "-s", help="SAP config JSON")
    sk.add_argument("--output", "-o", default="./analysis", help="Output directory")

    # example command
    parser.add_argument("--example", action="store_true", help="Print example config")

    args = parser.parse_args()

    if args.example:
        print(json.dumps({
            "table1_config": EXAMPLE_TABLE1_CONFIG,
        }, indent=2, ensure_ascii=False))
        return

    if args.command == "table1":
        if args.config:
            with open(args.config) as f:
                config = json.load(f)
        else:
            config = EXAMPLE_TABLE1_CONFIG
        result = generate_table1(args.data, config, args.output)
        print(result["markdown"])

    elif args.command == "skeleton":
        if args.sap_config:
            with open(args.sap_config) as f:
                sap_cfg = json.load(f)
        else:
            # Try importing from sap_generator
            try:
                from sap_generator import EXAMPLE_CONFIG
                sap_cfg = EXAMPLE_CONFIG
            except ImportError:
                print("ERROR: Provide --sap-config or install sap_generator.py")
                sys.exit(1)
        files = generate_analysis_skeleton(sap_cfg, args.output)
        print(f"✅ Created {len(files)} analysis scripts in {args.output}/")
        for f in files:
            print(f"  📄 {os.path.basename(f)}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

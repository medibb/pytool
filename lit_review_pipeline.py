#!/usr/bin/env python3
"""
lit_review_pipeline.py — 체계적 문헌검토 파이프라인

Phase 3 of the clinical research workflow:
  1. PICO → PubMed search query generator
  2. PubMed E-utilities API: search + fetch abstracts
  3. Data extraction template generator
  4. PRISMA 2020 flow diagram (text-based + matplotlib)

=== Usage ===

Python:
    from lit_review_pipeline import (
        build_search_query, search_pubmed, fetch_articles,
        generate_extraction_template, generate_prisma_flow,
        run_pipeline, EXAMPLE_PICO
    )

    # Quick pipeline
    results = run_pipeline(EXAMPLE_PICO, email="user@example.com")

CLI:
    python3 lit_review_pipeline.py --example --run --email user@example.com
    python3 lit_review_pipeline.py pico.json --email user@example.com -o ./output

Dependencies: biopython, python-docx, matplotlib (optional for PRISMA diagram)
"""

import os
import sys
import json
import argparse
import csv
from datetime import datetime
from collections import OrderedDict

try:
    from Bio import Entrez, Medline
    HAS_BIO = True
except ImportError:
    HAS_BIO = False

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    HAS_MPL = True
except ImportError:
    HAS_MPL = False


# ============================================================
# PICO → Search Query
# ============================================================

def build_search_query(pico, max_year=None, min_year=None):
    """Convert PICO dict to PubMed search query string.

    Args:
        pico: dict with keys: population, intervention, comparison, outcome,
              and optional: study_types, additional_terms, exclude_terms
        max_year: upper year limit (e.g. 2026)
        min_year: lower year limit (e.g. 2020)

    Returns:
        query: PubMed-compatible search string
    """
    parts = []

    # Population
    pop_terms = pico.get("population", [])
    if pop_terms:
        pop_str = " OR ".join(f'"{t}"[MeSH Terms] OR "{t}"[Title/Abstract]' for t in pop_terms)
        parts.append(f"({pop_str})")

    # Intervention
    int_terms = pico.get("intervention", [])
    if int_terms:
        int_str = " OR ".join(f'"{t}"[MeSH Terms] OR "{t}"[Title/Abstract]' for t in int_terms)
        parts.append(f"({int_str})")

    # Comparison (optional — often merged with intervention)
    comp_terms = pico.get("comparison", [])
    if comp_terms:
        comp_str = " OR ".join(f'"{t}"[Title/Abstract]' for t in comp_terms)
        parts.append(f"({comp_str})")

    # Outcome
    out_terms = pico.get("outcome", [])
    if out_terms:
        out_str = " OR ".join(f'"{t}"[Title/Abstract]' for t in out_terms)
        parts.append(f"({out_str})")

    query = " AND ".join(parts)

    # Study type filter
    study_types = pico.get("study_types", [])
    if study_types:
        st_str = " OR ".join(f'"{t}"[Publication Type]' for t in study_types)
        query += f" AND ({st_str})"

    # Additional terms
    for term in pico.get("additional_terms", []):
        query += f' AND "{term}"[Title/Abstract]'

    # Exclude terms
    for term in pico.get("exclude_terms", []):
        query += f' NOT "{term}"[Title/Abstract]'

    # Date range
    if min_year or max_year:
        date_min = f"{min_year}/01/01" if min_year else "1900/01/01"
        date_max = f"{max_year}/12/31" if max_year else "2099/12/31"
        query += f' AND ("{date_min}"[Date - Publication] : "{date_max}"[Date - Publication])'

    # Language filter
    lang = pico.get("language", "English")
    if lang:
        query += f' AND "{lang}"[Language]'

    return query


# ============================================================
# PubMed Search & Fetch
# ============================================================

def search_pubmed(query, email, max_results=200, sort="relevance"):
    """Search PubMed and return list of PMIDs.

    Args:
        query: PubMed search string
        email: required by NCBI E-utilities
        max_results: max number of results to retrieve
        sort: 'relevance' or 'date'

    Returns:
        dict: {pmids: list, count: int, query: str}
    """
    if not HAS_BIO:
        raise ImportError("biopython is required. Install: pip install biopython")

    Entrez.email = email
    handle = Entrez.esearch(
        db="pubmed",
        term=query,
        retmax=max_results,
        sort=sort,
        usehistory="y",
    )
    results = Entrez.read(handle)
    handle.close()

    return {
        "pmids": results.get("IdList", []),
        "count": int(results.get("Count", 0)),
        "query": query,
        "webenv": results.get("WebEnv", ""),
        "query_key": results.get("QueryKey", ""),
    }


def fetch_articles(pmids=None, email=None, webenv=None, query_key=None, batch_size=50):
    """Fetch article details from PubMed for given PMIDs.

    Args:
        pmids: list of PMID strings (use if no webenv)
        email: required by NCBI
        webenv, query_key: from search_pubmed (for large results)
        batch_size: fetch in batches

    Returns:
        list of article dicts with keys:
            pmid, title, authors, journal, year, abstract, doi, pub_type
    """
    if not HAS_BIO:
        raise ImportError("biopython is required.")

    Entrez.email = email
    articles = []

    if webenv and query_key:
        # Use history server for large result sets
        total = len(pmids) if pmids else 0
        for start in range(0, max(total, 1), batch_size):
            handle = Entrez.efetch(
                db="pubmed",
                rettype="medline",
                retmode="text",
                retstart=start,
                retmax=batch_size,
                webenv=webenv,
                query_key=query_key,
            )
            records = Medline.parse(handle)
            for record in records:
                articles.append(_parse_medline_record(record))
            handle.close()
    elif pmids:
        for i in range(0, len(pmids), batch_size):
            batch = pmids[i:i + batch_size]
            handle = Entrez.efetch(
                db="pubmed",
                id=",".join(batch),
                rettype="medline",
                retmode="text",
            )
            records = Medline.parse(handle)
            for record in records:
                articles.append(_parse_medline_record(record))
            handle.close()

    return articles


def _parse_medline_record(record):
    """Parse a Medline record into a clean dict."""
    authors = record.get("AU", [])
    first_author = authors[0] if authors else "Unknown"
    author_str = f"{first_author} et al." if len(authors) > 1 else first_author

    # Extract DOI from Article Identifier
    doi = ""
    for aid in record.get("AID", []):
        if "[doi]" in aid:
            doi = aid.replace(" [doi]", "")

    return {
        "pmid": record.get("PMID", ""),
        "title": record.get("TI", ""),
        "authors": author_str,
        "authors_full": "; ".join(authors),
        "journal": record.get("JT", record.get("TA", "")),
        "journal_abbr": record.get("TA", ""),
        "year": record.get("DP", "")[:4],
        "abstract": record.get("AB", ""),
        "doi": doi,
        "pub_type": "; ".join(record.get("PT", [])),
        "mesh": record.get("MH", []),
    }


# ============================================================
# Data Extraction Template
# ============================================================

def generate_extraction_template(pico, output_path="extraction_template.docx"):
    """Generate a systematic review data extraction template as docx.

    Creates a table-based template with standard fields for extracting
    data from included studies.
    """
    if not HAS_DOCX:
        raise ImportError("python-docx required.")

    doc = Document()
    title = pico.get("review_title", "Systematic Review Data Extraction")
    doc.add_heading(f"Data Extraction Form: {title}", level=0)
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d')}")
    doc.add_paragraph(f"Tool: lit_review_pipeline.py (pytool Phase 3)")
    doc.add_paragraph("")

    # Standard extraction fields
    sections = OrderedDict([
        ("Study Identification", [
            ("Study ID", ""),
            ("First Author", ""),
            ("Year", ""),
            ("Journal", ""),
            ("Country", ""),
            ("Funding Source", ""),
        ]),
        ("Study Design", [
            ("Design", "RCT / Cohort / Case-control / Cross-sectional / Case series"),
            ("Setting", "Hospital / Community / Mixed"),
            ("Duration", ""),
            ("Follow-up Period", ""),
            ("Risk of Bias", "Low / Some concerns / High"),
        ]),
        ("Participants", [
            ("Total N", ""),
            ("Intervention N", ""),
            ("Control N", ""),
            ("Age (mean ± SD)", ""),
            ("Sex (% female)", ""),
            ("Inclusion Criteria", ""),
            ("Exclusion Criteria", ""),
        ]),
        ("Intervention", [
            ("Intervention Description", ""),
            ("Comparator Description", ""),
            ("Co-interventions", ""),
            ("Adherence", ""),
        ]),
        ("Outcomes", []),  # Will be filled from PICO
        ("Results", [
            ("Primary Outcome Result", ""),
            ("Effect Size (95% CI)", ""),
            ("p-value", ""),
            ("Secondary Outcomes", ""),
            ("Adverse Events", ""),
        ]),
        ("Notes", [
            ("Reviewer Notes", ""),
            ("Quality Concerns", ""),
        ]),
    ])

    # Add outcome fields from PICO
    outcomes = pico.get("outcome", [])
    outcome_fields = []
    for o in outcomes:
        outcome_fields.append((f"Outcome: {o}", "Definition / Measurement / Time point"))
    if not outcome_fields:
        outcome_fields.append(("Primary Outcome", ""))
        outcome_fields.append(("Secondary Outcomes", ""))
    sections["Outcomes"] = outcome_fields

    for section_name, fields in sections.items():
        doc.add_heading(section_name, level=2)
        table = doc.add_table(rows=len(fields), cols=2)
        table.style = "Table Grid"
        for i, (label, hint) in enumerate(fields):
            row = table.rows[i]
            run = row.cells[0].paragraphs[0].add_run(label)
            run.bold = True
            run.font.size = Pt(10)
            if hint:
                run2 = row.cells[1].paragraphs[0].add_run(hint)
                run2.font.size = Pt(9)
                run2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.save(output_path)
    return output_path


# ============================================================
# PRISMA 2020 Flow Diagram
# ============================================================

def generate_prisma_flow(counts, output_path="prisma_flow.png"):
    """Generate PRISMA 2020 flow diagram.

    Args:
        counts: dict with keys:
            identified: total records from databases
            duplicates_removed: duplicates
            screened: title/abstract screened
            excluded_screening: excluded at screening
            sought_retrieval: sought for full-text
            not_retrieved: could not retrieve
            assessed_eligibility: full-text assessed
            excluded_reasons: dict {reason: n}
            included_studies: final included
            included_reports: final included reports (can equal studies)
        output_path: PNG output path

    Returns:
        output_path
    """
    if not HAS_MPL:
        # Fallback to text-based PRISMA
        return _prisma_text(counts)

    fig, ax = plt.subplots(figsize=(10, 14))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 14)
    ax.axis("off")

    box_props = dict(boxstyle="round,pad=0.4", facecolor="#E8F0FE", edgecolor="#2F5496", linewidth=1.5)
    excl_props = dict(boxstyle="round,pad=0.4", facecolor="#FFF3E0", edgecolor="#E65100", linewidth=1.5)
    final_props = dict(boxstyle="round,pad=0.4", facecolor="#E8F5E9", edgecolor="#2E7D32", linewidth=1.5)
    arrow_props = dict(arrowstyle="-|>", color="#333333", lw=1.5)

    c = counts  # shorthand

    # Box positions (x, y)
    boxes = [
        (5, 13, f"Identification\n\nRecords identified (n={c.get('identified', 0)})", box_props),
        (5, 11, f"Duplicates removed (n={c.get('duplicates_removed', 0)})\n\nRecords screened\n(n={c.get('screened', 0)})", box_props),
        (8.5, 11, f"Records excluded\n(n={c.get('excluded_screening', 0)})", excl_props),
        (5, 8.5, f"Reports sought for retrieval\n(n={c.get('sought_retrieval', 0)})", box_props),
        (8.5, 8.5, f"Reports not retrieved\n(n={c.get('not_retrieved', 0)})", excl_props),
        (5, 6, f"Reports assessed for eligibility\n(n={c.get('assessed_eligibility', 0)})", box_props),
    ]

    # Excluded reasons
    excl_reasons = c.get("excluded_reasons", {})
    excl_text = "Reports excluded:\n" + "\n".join(
        f"  {reason} (n={n})" for reason, n in excl_reasons.items()
    ) if excl_reasons else f"Reports excluded (n={c.get('assessed_eligibility', 0) - c.get('included_studies', 0)})"
    boxes.append((8.5, 6, excl_text, excl_props))

    # Final included
    inc_text = f"Studies included\n(n={c.get('included_studies', 0)})"
    boxes.append((5, 3.5, inc_text, final_props))

    # Draw boxes
    for (x, y, text, props) in boxes:
        ax.text(x, y, text, ha="center", va="center", fontsize=9,
                bbox=props, family="sans-serif")

    # Arrows (from_xy, to_xy)
    arrows = [
        ((5, 12.4), (5, 11.7)),      # Identified → Screened
        ((6.5, 11), (7.5, 11)),       # Screened → Excluded screening
        ((5, 10.3), (5, 9.2)),        # Screened → Sought
        ((6.5, 8.5), (7.5, 8.5)),     # Sought → Not retrieved
        ((5, 7.8), (5, 6.7)),         # Sought → Eligibility
        ((6.5, 6), (7.5, 6)),         # Eligibility → Excluded
        ((5, 5.3), (5, 4.2)),         # Eligibility → Included
    ]

    for (start, end) in arrows:
        ax.annotate("", xy=end, xytext=start, arrowprops=arrow_props)

    # Section labels
    ax.text(0.5, 13, "IDENTIFICATION", fontsize=8, fontweight="bold", color="#666",
            rotation=90, va="center")
    ax.text(0.5, 11, "SCREENING", fontsize=8, fontweight="bold", color="#666",
            rotation=90, va="center")
    ax.text(0.5, 7.5, "ELIGIBILITY", fontsize=8, fontweight="bold", color="#666",
            rotation=90, va="center")
    ax.text(0.5, 3.5, "INCLUDED", fontsize=8, fontweight="bold", color="#666",
            rotation=90, va="center")

    # Title
    ax.text(5, 14.2, "PRISMA 2020 Flow Diagram", ha="center", fontsize=14,
            fontweight="bold", color="#2F5496")

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches="tight")
    plt.close()

    return output_path


def _prisma_text(counts):
    """Generate text-based PRISMA flow diagram."""
    c = counts
    lines = [
        "PRISMA 2020 Flow Diagram",
        "=" * 50,
        f"IDENTIFICATION",
        f"  Records identified: {c.get('identified', 0)}",
        f"  Duplicates removed: {c.get('duplicates_removed', 0)}",
        "",
        f"SCREENING",
        f"  Records screened: {c.get('screened', 0)}",
        f"  Records excluded: {c.get('excluded_screening', 0)}",
        "",
        f"ELIGIBILITY",
        f"  Sought for retrieval: {c.get('sought_retrieval', 0)}",
        f"  Not retrieved: {c.get('not_retrieved', 0)}",
        f"  Assessed for eligibility: {c.get('assessed_eligibility', 0)}",
    ]
    for reason, n in c.get("excluded_reasons", {}).items():
        lines.append(f"    Excluded — {reason}: {n}")
    lines.extend([
        "",
        f"INCLUDED",
        f"  Studies included: {c.get('included_studies', 0)}",
    ])
    return "\n".join(lines)


# ============================================================
# Articles → CSV/Markdown export
# ============================================================

def export_articles_csv(articles, output_path="search_results.csv"):
    """Export articles to CSV for screening."""
    fields = ["pmid", "year", "authors", "title", "journal_abbr", "doi", "pub_type"]
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields + ["include", "reason"],
                                extrasaction="ignore")
        writer.writeheader()
        for a in articles:
            row = {k: a.get(k, "") for k in fields}
            row["include"] = ""  # Blank for reviewer
            row["reason"] = ""
            writer.writerow(row)
    return output_path


def export_articles_markdown(articles, output_path="search_results.md"):
    """Export articles to Markdown table."""
    lines = [
        "# PubMed Search Results",
        "",
        f"Total: {len(articles)} articles",
        "",
        "| # | Year | Authors | Title | Journal | PMID |",
        "|---|------|---------|-------|---------|------|",
    ]
    for i, a in enumerate(articles, 1):
        title_short = a["title"][:60] + "..." if len(a["title"]) > 60 else a["title"]
        lines.append(
            f"| {i} | {a['year']} | {a['authors']} | {title_short} | "
            f"{a['journal_abbr']} | {a['pmid']} |"
        )
    lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return output_path


# ============================================================
# Full Pipeline
# ============================================================

def run_pipeline(pico, email, output_dir=".", max_results=100):
    """Run complete literature review pipeline.

    1. Build search query from PICO
    2. Search PubMed
    3. Fetch article details
    4. Export to CSV + Markdown
    5. Generate extraction template
    6. Generate PRISMA flow (placeholder counts)

    Returns dict with all results and file paths.
    """
    os.makedirs(output_dir, exist_ok=True)

    # 1. Build query
    query = build_search_query(
        pico,
        min_year=pico.get("min_year"),
        max_year=pico.get("max_year"),
    )
    print(f"📋 Search query:\n{query}\n")

    # 2. Search
    search_result = search_pubmed(query, email, max_results=max_results)
    pmids = search_result["pmids"]
    total_count = search_result["count"]
    print(f"🔍 Found {total_count} results, fetching top {len(pmids)}...")

    # 3. Fetch
    articles = fetch_articles(
        pmids=pmids, email=email,
        webenv=search_result.get("webenv"),
        query_key=search_result.get("query_key"),
    )
    print(f"📄 Fetched {len(articles)} article details")

    # 4. Export
    csv_path = export_articles_csv(articles, os.path.join(output_dir, "search_results.csv"))
    md_path = export_articles_markdown(articles, os.path.join(output_dir, "search_results.md"))
    print(f"💾 Exported: {csv_path}, {md_path}")

    # 5. Extraction template
    template_path = generate_extraction_template(
        pico, os.path.join(output_dir, "extraction_template.docx")
    )
    print(f"📋 Extraction template: {template_path}")

    # 6. PRISMA flow (initial counts — user fills in after screening)
    prisma_counts = {
        "identified": total_count,
        "duplicates_removed": 0,
        "screened": len(articles),
        "excluded_screening": 0,
        "sought_retrieval": len(articles),
        "not_retrieved": 0,
        "assessed_eligibility": len(articles),
        "excluded_reasons": {},
        "included_studies": len(articles),
    }
    prisma_path = generate_prisma_flow(
        prisma_counts, os.path.join(output_dir, "prisma_flow.png")
    )
    print(f"📊 PRISMA flow: {prisma_path}")

    # Save query and metadata
    meta = {
        "pico": pico,
        "query": query,
        "total_results": total_count,
        "fetched": len(articles),
        "date": datetime.now().isoformat(),
    }
    meta_path = os.path.join(output_dir, "search_metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)

    return {
        "query": query,
        "total_count": total_count,
        "articles": articles,
        "files": {
            "csv": csv_path,
            "markdown": md_path,
            "extraction_template": template_path,
            "prisma_flow": prisma_path,
            "metadata": meta_path,
        },
    }


# ============================================================
# Example PICO
# ============================================================

EXAMPLE_PICO = {
    "review_title": "Hydrodilatation for Adhesive Capsulitis: A Systematic Review",
    "population": ["adhesive capsulitis", "frozen shoulder"],
    "intervention": ["hydrodilatation", "hydraulic distension", "distension arthrography"],
    "comparison": ["corticosteroid injection", "physiotherapy", "placebo"],
    "outcome": ["SPADI", "range of motion", "pain", "function"],
    "study_types": ["Randomized Controlled Trial", "Clinical Trial"],
    "additional_terms": [],
    "exclude_terms": ["animal", "cadaver"],
    "language": "English",
    "min_year": 2015,
    "max_year": 2026,
}


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Systematic Literature Review Pipeline (PubMed)"
    )
    parser.add_argument("pico_config", nargs="?", help="PICO config JSON file")
    parser.add_argument("--email", "-e", help="Email for NCBI E-utilities (required)")
    parser.add_argument("--output", "-o", default="./lit_review_output",
                        help="Output directory")
    parser.add_argument("--max-results", "-n", type=int, default=100,
                        help="Max results to fetch")
    parser.add_argument("--example", action="store_true", help="Use example PICO config")
    parser.add_argument("--run", action="store_true", help="Run pipeline (use with --example)")
    parser.add_argument("--query-only", action="store_true",
                        help="Only build and print search query (no API call)")

    args = parser.parse_args()

    if args.example and not args.run:
        print(json.dumps(EXAMPLE_PICO, indent=2, ensure_ascii=False))
        return

    # Load config
    if args.example:
        pico = EXAMPLE_PICO
    elif args.pico_config:
        with open(args.pico_config) as f:
            pico = json.load(f)
    else:
        parser.print_help()
        print("\nUse --example --run --email your@email.com to test.")
        sys.exit(1)

    # Query only
    if args.query_only:
        query = build_search_query(pico, min_year=pico.get("min_year"),
                                   max_year=pico.get("max_year"))
        print(f"PubMed Query:\n{query}")
        return

    if not args.email:
        print("ERROR: --email is required for PubMed API calls.")
        sys.exit(1)

    result = run_pipeline(pico, args.email, args.output, args.max_results)
    print(f"\n✅ Pipeline complete: {result['total_count']} results found, "
          f"{len(result['articles'])} fetched")


if __name__ == "__main__":
    main()

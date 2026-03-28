#!/usr/bin/env python3
"""
protocol_crf_mapper.py — Protocol ↔ CRF 일관성 검증 + SPIRIT 2025 Checklist Validator

Phase 2 of the clinical research workflow:
  1. Protocol-CRF Mapping Matrix: Cross-reference all variables between protocol,
     CRF, SAP, pre-exam form, and quick sheet → gap report
  2. SPIRIT 2025 Checklist Validator: Check protocol config against 50 sub-items

=== Usage ===

Python:
    from protocol_crf_mapper import (
        generate_mapping_matrix, validate_spirit_2025,
        generate_full_report, EXAMPLE_SCAFFOLD_CONFIG, EXAMPLE_SAP_CONFIG
    )

    # Mapping matrix
    matrix = generate_mapping_matrix(scaffold_config, sap_config)
    # → {'variables': [...], 'gaps': [...], 'summary': {...}}

    # SPIRIT 2025 checklist
    result = validate_spirit_2025(scaffold_config, sap_config)
    # → {'items': [...], 'score': 32/50, 'missing': [...]}

    # Full report (docx)
    generate_full_report(scaffold_config, sap_config, "consistency_report.docx")

CLI:
    python3 protocol_crf_mapper.py scaffold.json sap.json --output report.docx
    python3 protocol_crf_mapper.py --example --run

Dependencies: python-docx
"""

import os
import sys
import json
import argparse
from datetime import datetime
from collections import OrderedDict

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ============================================================
# SPIRIT 2025 Checklist (50 sub-items, 9 sections)
# ============================================================

SPIRIT_2025_CHECKLIST = [
    # Section 1: Administrative Information
    {"id": "1a", "section": "Administrative", "name": "Title",
     "description": "Title stating trial design, population, interventions; identified as protocol",
     "keywords": ["title", "study_title", "design", "protocol"],
     "config_paths": [("sap", "study.title"), ("scaffold", "study.title_en"), ("scaffold", "study.title_ko")]},
    {"id": "1b", "section": "Administrative", "name": "Structured Summary",
     "description": "Structured summary with WHO Trial Registration Data Set items",
     "keywords": ["summary", "abstract", "registration"],
     "config_paths": [("sap", "objectives.primary.description"), ("sap", "design.type")]},
    {"id": "2", "section": "Administrative", "name": "Protocol Version",
     "description": "Version date and identifier",
     "keywords": ["version", "date", "protocol_number"],
     "config_paths": [("sap", "study.sap_version"), ("sap", "study.sap_date"), ("sap", "study.protocol_number")]},
    {"id": "3a", "section": "Administrative", "name": "Roles & Responsibilities",
     "description": "Names, affiliations, roles of protocol contributors",
     "keywords": ["pi", "investigator", "sponsor", "role"],
     "config_paths": [("sap", "study.pi"), ("sap", "study.sponsor"), ("scaffold", "study.pi"), ("scaffold", "study.department")]},
    {"id": "3b", "section": "Administrative", "name": "Sponsor Contact",
     "description": "Name and contact information for trial sponsor",
     "keywords": ["sponsor", "contact"],
     "config_paths": [("sap", "study.sponsor")]},
    {"id": "3c", "section": "Administrative", "name": "Sponsor/Funder Role",
     "description": "Role in design, conduct, analysis, reporting",
     "keywords": ["funder", "funding_role", "sponsor_role"],
     "config_paths": []},
    {"id": "3d", "section": "Administrative", "name": "Trial Oversight",
     "description": "Coordinating site, steering committee, endpoint adjudication, data management",
     "keywords": ["oversight", "committee", "steering", "adjudication"],
     "config_paths": []},

    # Section 2: Open Science
    {"id": "4", "section": "Open Science", "name": "Trial Registration",
     "description": "Registry name, identifying number (URL), registration date",
     "keywords": ["registration", "registry", "clinicaltrials", "cris"],
     "config_paths": [("sap", "study.protocol_number")]},
    {"id": "5", "section": "Open Science", "name": "Protocol & SAP Accessibility",
     "description": "Where trial protocol and SAP can be accessed",
     "keywords": ["access", "protocol_access", "sap_access"],
     "config_paths": []},
    {"id": "6", "section": "Open Science", "name": "Data Sharing",
     "description": "Where/how deidentified data, data dictionary, code accessible",
     "keywords": ["data_sharing", "deidentified", "repository"],
     "config_paths": []},
    {"id": "7a", "section": "Open Science", "name": "Funding Sources",
     "description": "Sources of funding and other support",
     "keywords": ["funding", "grant", "support"],
     "config_paths": []},
    {"id": "7b", "section": "Open Science", "name": "Conflicts of Interest",
     "description": "Financial/other COI for PI and steering committee",
     "keywords": ["conflict", "coi", "disclosure"],
     "config_paths": []},
    {"id": "8", "section": "Open Science", "name": "Dissemination Plans",
     "description": "Plans to communicate results to participants, professionals, public",
     "keywords": ["dissemination", "publication", "communication"],
     "config_paths": []},

    # Section 3: Introduction
    {"id": "9a", "section": "Introduction", "name": "Scientific Background",
     "description": "Background and rationale including summary of relevant studies",
     "keywords": ["background", "rationale", "literature"],
     "config_paths": []},
    {"id": "9b", "section": "Introduction", "name": "Comparator Justification",
     "description": "Explanation for choice of comparator",
     "keywords": ["comparator", "control", "justification"],
     "config_paths": [("sap", "design.arms")]},
    {"id": "10", "section": "Introduction", "name": "Objectives",
     "description": "Specific objectives related to benefits and harms",
     "keywords": ["objective", "aim", "hypothesis"],
     "config_paths": [("sap", "objectives.primary"), ("sap", "objectives.secondary")]},

    # Section 4: Methods - Patient Involvement & Design
    {"id": "11", "section": "Methods-PPI", "name": "Patient & Public Involvement (NEW)",
     "description": "Details of patient/public involvement in design, conduct, reporting",
     "keywords": ["patient_involvement", "ppi", "public"],
     "config_paths": []},
    {"id": "12", "section": "Methods-Design", "name": "Trial Design",
     "description": "Type, allocation ratio, framework (superiority/equivalence/non-inferiority)",
     "keywords": ["design", "parallel", "crossover", "allocation_ratio", "superiority", "non_inferiority"],
     "config_paths": [("sap", "design.type"), ("sap", "design.arms")]},

    # Section 5: Methods - Participants, Interventions, Outcomes
    {"id": "13", "section": "Methods-Participants", "name": "Trial Setting",
     "description": "Settings (community, hospital) and locations",
     "keywords": ["setting", "hospital", "site", "location"],
     "config_paths": [("scaffold", "study.department"), ("sap", "study.sponsor")]},
    {"id": "14a", "section": "Methods-Participants", "name": "Participant Eligibility",
     "description": "Eligibility criteria for participants",
     "keywords": ["inclusion", "exclusion", "eligibility", "criteria"],
     "config_paths": [("scaffold", "inclusion"), ("scaffold", "exclusion")]},
    {"id": "14b", "section": "Methods-Participants", "name": "Site/Provider Eligibility",
     "description": "Eligibility criteria for sites and intervention providers",
     "keywords": ["provider", "site_eligibility", "qualifications"],
     "config_paths": []},
    {"id": "15a", "section": "Methods-Interventions", "name": "Intervention Description",
     "description": "Sufficient detail for replication: how, when, by whom",
     "keywords": ["intervention", "procedure", "treatment", "protocol"],
     "config_paths": [("scaffold", "procedure")]},
    {"id": "15b", "section": "Methods-Interventions", "name": "Intervention Modification",
     "description": "Criteria for discontinuing/modifying allocated intervention",
     "keywords": ["discontinuation", "modification", "stopping"],
     "config_paths": []},
    {"id": "15c", "section": "Methods-Interventions", "name": "Adherence Strategies",
     "description": "Strategies to improve adherence; monitoring procedures",
     "keywords": ["adherence", "compliance", "monitoring"],
     "config_paths": []},
    {"id": "15d", "section": "Methods-Interventions", "name": "Concomitant Care",
     "description": "Permitted or prohibited concomitant care",
     "keywords": ["concomitant", "prohibited", "permitted", "medication"],
     "config_paths": []},
    {"id": "16", "section": "Methods-Outcomes", "name": "Outcomes",
     "description": "Primary/secondary outcomes: variable, metric, aggregation, time point",
     "keywords": ["primary_outcome", "secondary_outcome", "endpoint", "measure"],
     "config_paths": [("sap", "primary_outcome"), ("sap", "secondary_outcomes"),
                      ("scaffold", "baseline"), ("scaffold", "followup")]},
    {"id": "17", "section": "Methods-Outcomes", "name": "Harms Assessment",
     "description": "How harms are defined and assessed",
     "keywords": ["harm", "adverse_event", "safety", "ae"],
     "config_paths": [("scaffold", "procedure")]},
    {"id": "18", "section": "Methods-Outcomes", "name": "Participant Timeline",
     "description": "Schedule of enrolment, interventions, assessments, visits",
     "keywords": ["timeline", "schedule", "visits", "followup"],
     "config_paths": [("scaffold", "followup")]},
    {"id": "19", "section": "Methods-Outcomes", "name": "Sample Size",
     "description": "How sample size was determined with supporting assumptions",
     "keywords": ["sample_size", "power", "alpha", "mcid", "effect_size"],
     "config_paths": [("sap", "primary_outcome.mcid"), ("sap", "primary_outcome.sd"),
                      ("sap", "primary_outcome.alpha"), ("sap", "primary_outcome.power")]},
    {"id": "20", "section": "Methods-Outcomes", "name": "Recruitment Strategies",
     "description": "Strategies for achieving adequate enrolment",
     "keywords": ["recruitment", "enrolment", "screening"],
     "config_paths": []},

    # Section 6: Methods - Assignment of Interventions
    {"id": "21a", "section": "Methods-Assignment", "name": "Sequence Generation",
     "description": "Who generates random allocation and method used",
     "keywords": ["randomization", "sequence", "generation"],
     "config_paths": [("sap", "design.randomization.method")]},
    {"id": "21b", "section": "Methods-Assignment", "name": "Randomization Type",
     "description": "Simple/restricted; stratification factors; planned restrictions",
     "keywords": ["stratification", "block", "restriction"],
     "config_paths": [("sap", "design.randomization.block_sizes"),
                      ("sap", "design.randomization.stratification")]},
    {"id": "22", "section": "Methods-Assignment", "name": "Allocation Concealment",
     "description": "Mechanism to conceal allocation sequence until assignment",
     "keywords": ["concealment", "sealed", "opaque", "central"],
     "config_paths": []},
    {"id": "23", "section": "Methods-Assignment", "name": "Implementation",
     "description": "Whether enrolling/assigning personnel access allocation sequence",
     "keywords": ["implementation", "enrollment_personnel"],
     "config_paths": []},
    {"id": "24a", "section": "Methods-Assignment", "name": "Blinding Specification",
     "description": "Who is blinded: participants, providers, assessors, analysts",
     "keywords": ["blinding", "masking", "blind"],
     "config_paths": [("sap", "design.blinding")]},
    {"id": "24b", "section": "Methods-Assignment", "name": "Blinding Method",
     "description": "How blinding achieved; intervention similarity description",
     "keywords": ["blinding_method", "placebo", "sham"],
     "config_paths": []},
    {"id": "24c", "section": "Methods-Assignment", "name": "Unblinding Procedures",
     "description": "Circumstances permitting unblinding; allocation reveal procedure",
     "keywords": ["unblinding", "code_breaking", "emergency"],
     "config_paths": []},

    # Section 7: Methods - Data Collection and Analysis
    {"id": "25a", "section": "Methods-Data", "name": "Data Collection Methods",
     "description": "Assessment/collection plans, quality processes, instrument reliability",
     "keywords": ["data_collection", "instrument", "reliability", "validity"],
     "config_paths": [("scaffold", "baseline"), ("scaffold", "procedure")]},
    {"id": "25b", "section": "Methods-Data", "name": "Participant Retention",
     "description": "Retention/follow-up plans; outcome data for discontinuers",
     "keywords": ["retention", "dropout", "lost_to_followup"],
     "config_paths": [("sap", "primary_outcome.dropout_rate")]},
    {"id": "26", "section": "Methods-Data", "name": "Data Management",
     "description": "Data entry, coding, security, storage; quality processes",
     "keywords": ["data_management", "entry", "coding", "security"],
     "config_paths": []},
    {"id": "27a", "section": "Methods-Analysis", "name": "Statistical Methods",
     "description": "Methods comparing groups for primary/secondary/harms outcomes",
     "keywords": ["statistical_method", "t_test", "ancova", "regression"],
     "config_paths": [("sap", "analysis.primary_method")]},
    {"id": "27b", "section": "Methods-Analysis", "name": "Analysis Population",
     "description": "Definition of each analysis population (ITT, PP, safety)",
     "keywords": ["itt", "per_protocol", "intention_to_treat", "population"],
     "config_paths": [("sap", "populations")]},
    {"id": "27c", "section": "Methods-Analysis", "name": "Missing Data Handling",
     "description": "How missing data will be handled",
     "keywords": ["missing", "imputation", "mice", "last_observation"],
     "config_paths": [("sap", "analysis.missing_data")]},
    {"id": "27d", "section": "Methods-Analysis", "name": "Additional Analyses",
     "description": "Subgroup and sensitivity analyses methods",
     "keywords": ["subgroup", "sensitivity", "exploratory"],
     "config_paths": [("sap", "analysis.subgroup"), ("sap", "analysis.sensitivity")]},

    # Section 8: Methods - Monitoring
    {"id": "28a", "section": "Methods-Monitoring", "name": "Data Monitoring Committee",
     "description": "DMC composition, role, reporting, independence, charter",
     "keywords": ["dmc", "dsmb", "monitoring_committee"],
     "config_paths": []},
    {"id": "28b", "section": "Methods-Monitoring", "name": "Interim Analyses",
     "description": "Stopping guidelines; interim results access; decision authority",
     "keywords": ["interim", "stopping", "futility"],
     "config_paths": []},
    {"id": "29", "section": "Methods-Monitoring", "name": "Trial Monitoring (NEW)",
     "description": "Frequency and procedures for monitoring trial conduct",
     "keywords": ["monitoring", "audit", "site_visit"],
     "config_paths": []},

    # Section 9: Ethics and Dissemination
    {"id": "30", "section": "Ethics", "name": "Research Ethics Approval",
     "description": "Plans for seeking REC/IRB approval",
     "keywords": ["irb", "ethics", "rec", "approval"],
     "config_paths": [("sap", "study.protocol_number")]},
    {"id": "31", "section": "Ethics", "name": "Protocol Amendments",
     "description": "Plans for communicating modifications to relevant parties",
     "keywords": ["amendment", "modification", "revision"],
     "config_paths": [("sap", "study.sap_version")]},
    {"id": "32a", "section": "Ethics", "name": "Informed Consent Process",
     "description": "Who obtains consent/assent, how",
     "keywords": ["consent", "informed_consent", "assent"],
     "config_paths": []},
    {"id": "32b", "section": "Ethics", "name": "Ancillary Study Consent",
     "description": "Additional consent for data/specimens in ancillary studies",
     "keywords": ["ancillary", "biobank", "specimen"],
     "config_paths": []},
    {"id": "33", "section": "Ethics", "name": "Confidentiality",
     "description": "Protection of personal information",
     "keywords": ["confidentiality", "privacy", "deidentification"],
     "config_paths": []},
    {"id": "34", "section": "Ethics", "name": "Post-Trial Care",
     "description": "Provisions for post-trial care and compensation for harm",
     "keywords": ["post_trial", "compensation", "care"],
     "config_paths": []},
]


# ============================================================
# Variable extraction helpers
# ============================================================

def _deep_get(obj, path):
    """Get nested dict value by dot-separated path. Returns None if not found."""
    keys = path.split(".")
    for k in keys:
        if isinstance(obj, dict) and k in obj:
            obj = obj[k]
        else:
            return None
    return obj


def _extract_scaffold_variables(config):
    """Extract all variables from a clinical_research_scaffold config.

    Returns list of dicts: {name, section, source, form, hint}
    - form: 'preexam' (예진지), 'quicksheet', 'crf' (all), 'enrollment', 'followup'
    """
    variables = []

    # Enrollment fields
    for item in config.get("enrollment", []):
        variables.append({
            "name": item["label"],
            "section": "Enrollment",
            "source": item.get("source", "unknown"),
            "form": "preexam",
            "hint": item.get("hint", ""),
            "origin": "scaffold",
        })

    # Inclusion criteria → implicit variables
    for i, criterion in enumerate(config.get("inclusion", []), 1):
        variables.append({
            "name": f"Inclusion #{i}: {criterion}",
            "section": "Inclusion",
            "source": "protocol",
            "form": "screening",
            "hint": criterion,
            "origin": "scaffold",
        })

    # Exclusion criteria → implicit variables
    for i, criterion in enumerate(config.get("exclusion", []), 1):
        variables.append({
            "name": f"Exclusion #{i}: {criterion}",
            "section": "Exclusion",
            "source": "protocol",
            "form": "screening",
            "hint": criterion,
            "origin": "scaffold",
        })

    # Baseline sections
    for sec in config.get("baseline", {}).get("sections", []):
        for item in sec.get("items", []):
            variables.append({
                "name": item["label"],
                "section": f"Baseline: {sec['name']}",
                "source": item.get("source", "unknown"),
                "form": "preexam",
                "hint": item.get("hint", ""),
                "origin": "scaffold",
            })

    # Procedure sections
    for sec in config.get("procedure", {}).get("sections", []):
        for item in sec.get("items", []):
            variables.append({
                "name": item["label"],
                "section": f"Procedure: {sec['name']}",
                "source": item.get("source", "unknown"),
                "form": "quicksheet",
                "hint": item.get("hint", ""),
                "origin": "scaffold",
            })

    # Follow-up
    followup = config.get("followup", {})
    for measure in followup.get("measures", []):
        variables.append({
            "name": measure,
            "section": "Follow-up",
            "source": "followup",
            "form": "crf",
            "hint": f"Visits: {', '.join(followup.get('visits', []))}",
            "origin": "scaffold",
        })

    return variables


def _extract_sap_variables(config):
    """Extract outcome variables from SAP config.

    Returns list of dicts: {name, section, type, role}
    - role: 'primary', 'secondary', 'exploratory', 'subgroup_factor', 'safety'
    """
    variables = []

    # Primary outcome
    po = config.get("primary_outcome", {})
    if po:
        variables.append({
            "name": po.get("name", ""),
            "section": "Primary Outcome",
            "type": po.get("type", "continuous"),
            "role": "primary",
            "test": po.get("test", ""),
            "origin": "sap",
        })

    # Secondary outcomes
    for so in config.get("secondary_outcomes", []):
        variables.append({
            "name": so.get("name", ""),
            "section": "Secondary Outcome",
            "type": so.get("type", "continuous"),
            "role": "secondary",
            "origin": "sap",
        })

    # Exploratory outcomes
    for eo in config.get("objectives", {}).get("exploratory", []):
        variables.append({
            "name": eo if isinstance(eo, str) else eo.get("name", ""),
            "section": "Exploratory Outcome",
            "type": "mixed",
            "role": "exploratory",
            "origin": "sap",
        })

    # Subgroup factors
    for sg in config.get("analysis", {}).get("subgroup", []):
        variables.append({
            "name": sg if isinstance(sg, str) else sg.get("factor", ""),
            "section": "Subgroup Factor",
            "type": "categorical",
            "role": "subgroup_factor",
            "origin": "sap",
        })

    return variables


# ============================================================
# Fuzzy matching for variable cross-referencing
# ============================================================

def _normalize(text):
    """Normalize text for fuzzy matching."""
    import re
    text = text.lower().strip()
    # Remove parenthetical hints
    text = re.sub(r'\([^)]*\)', '', text)
    # Remove common measurement units and placeholders
    text = re.sub(r'(___|\d+)\s*(°|mm|ml|cm|kg|/\d+)', '', text)
    # Remove filler words
    text = re.sub(r'\b(the|a|an|at|in|of|to|from|for|and|vs|with)\b', '', text)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _tokenize(text):
    """Extract meaningful tokens (length >= 2) from normalized text."""
    return set(t for t in _normalize(text).split() if len(t) >= 2)


def _match_score(name_a, name_b):
    """Token overlap matching score (0-1) with minimum token length filter."""
    tokens_a = _tokenize(name_a)
    tokens_b = _tokenize(name_b)
    if not tokens_a or not tokens_b:
        return 0.0
    # Require at least one meaningful token match
    intersection = tokens_a & tokens_b
    if not intersection:
        return 0.0
    # Weighted Jaccard: penalize when match is only trivial tokens
    union = tokens_a | tokens_b
    return len(intersection) / len(union) if union else 0.0


# Known synonym groups for clinical research variables.
# Each key maps to a list of terms that should be considered equivalent.
_SYNONYMS = {
    "spadi": ["spadi", "shoulder pain and disability index"],
    "nrs_rest": ["resting pain", "pain nrs rest", "nrs rest"],
    "nrs_activity": ["activity pain", "pain nrs activity", "nrs during activity"],
    "nrs_night": ["night pain", "pain nrs night", "night pain nrs"],
    "rom_er": ["external rotation", "er rom", "shoulder er"],
    "rom_ir": ["internal rotation", "ir rom", "shoulder ir"],
    "rom_abd": ["abduction", "abd rom", "shoulder abd"],
    "rom_flex": ["forward flexion", "flexion", "shoulder flexion"],
    "rom_ext": ["extension"],
    "pga": ["pga", "patient global assessment", "global assessment"],
    "sst": ["sst", "simple shoulder test", "sst score", "sst status"],
    "bmi": ["bmi", "body mass index"],
    "vas": ["vas", "visual analog scale", "visual analogue scale"],
    "age": ["age", "나이", "생년월일"],
    "sex": ["sex", "성별", "gender", "male", "female"],
    "diabetes": ["diabetes", "당뇨", "hba1c", "dm", "diabetes mellitus"],
    "adverse": ["adverse", "harm", "safety", "vasovagal", "swelling", "ecchymosis"],
    "dominant_hand": ["dominant", "우세손", "dominant hand"],
    "symptom_duration": ["symptom duration", "이환 기간", "증상 시작일", "이환기간"],
    "reinjection": ["reinjection", "재주사"],
    "ultrasound": ["ultrasound", "us finding", "capsular thickness", "chl thickness",
                    "axillary recess", "biceps effusion", "power doppler"],
    "capsule": ["capsule", "capsular", "capsule status", "rupture"],
}


def _synonym_match(name_a, name_b):
    """Check if two variable names match via synonyms. Returns group key or None."""
    na = _normalize(name_a)
    nb = _normalize(name_b)
    for group_key, synonyms in _SYNONYMS.items():
        match_a = any(s in na for s in synonyms)
        match_b = any(s in nb for s in synonyms)
        if match_a and match_b:
            return group_key
    return None


def _find_best_match(target_name, candidate_list, threshold=0.35):
    """Find best matching variable from candidates.

    Uses synonym matching (highest priority), then token overlap, then substring.
    Returns (best_match_dict, score) or (None, 0) if no match above threshold.
    """
    best = None
    best_score = 0.0

    target_norm = _normalize(target_name)
    target_tokens = _tokenize(target_name)

    for candidate in candidate_list:
        cname = candidate["name"]
        cname_norm = _normalize(cname)
        score = 0.0

        # 1. Synonym match (strongest signal)
        syn_group = _synonym_match(target_name, cname)
        if syn_group:
            score = max(score, 0.6)

        # 2. Token overlap (Jaccard)
        token_score = _match_score(target_name, cname)
        score = max(score, token_score)

        # 3. Substring containment (one contains the other)
        if len(target_norm) >= 3 and len(cname_norm) >= 3:
            if target_norm in cname_norm or cname_norm in target_norm:
                score = max(score, 0.55)

        if score > best_score:
            best_score = score
            best = candidate

    if best_score >= threshold:
        return best, best_score
    return None, 0.0


# ============================================================
# Core: Mapping Matrix Generation
# ============================================================

def generate_mapping_matrix(scaffold_config, sap_config):
    """Generate Protocol ↔ CRF Mapping Matrix.

    Cross-references variables between scaffold (CRF/preexam/quicksheet)
    and SAP (primary/secondary/exploratory outcomes).

    Returns dict:
        variables: list of mapped variable records
        gaps: list of unmatched variables (gaps in coverage)
        summary: {total, matched, unmatched_scaffold, unmatched_sap}
    """
    scaffold_vars = _extract_scaffold_variables(scaffold_config)
    sap_vars = _extract_sap_variables(sap_config)

    mapped = []
    matched_sap_indices = set()

    # For each SAP variable, find best CRF match
    for sap_var in sap_vars:
        match, score = _find_best_match(sap_var["name"], scaffold_vars)
        record = {
            "sap_variable": sap_var["name"],
            "sap_role": sap_var["role"],
            "sap_section": sap_var["section"],
        }
        if match:
            record.update({
                "crf_variable": match["name"],
                "crf_section": match["section"],
                "crf_form": match["form"],
                "data_source": match.get("source", ""),
                "match_score": round(score, 2),
                "status": "matched" if score >= 0.4 else "weak_match",
            })
        else:
            record.update({
                "crf_variable": "—",
                "crf_section": "—",
                "crf_form": "—",
                "data_source": "—",
                "match_score": 0,
                "status": "MISSING_IN_CRF",
            })
        mapped.append(record)

    # Find scaffold variables not referenced in SAP
    sap_names_norm = [_normalize(v["name"]) for v in sap_vars]
    unreferenced = []
    for sv in scaffold_vars:
        if sv["section"].startswith("Inclusion") or sv["section"].startswith("Exclusion"):
            continue  # Inclusion/exclusion are protocol-level, not SAP outcomes
        match, score = _find_best_match(sv["name"], sap_vars)
        if not match:
            unreferenced.append({
                "crf_variable": sv["name"],
                "crf_section": sv["section"],
                "crf_form": sv["form"],
                "data_source": sv.get("source", ""),
                "note": "Collected in CRF but not referenced in SAP analysis plan",
            })

    # Gaps = SAP variables missing from CRF
    gaps = [r for r in mapped if r["status"] == "MISSING_IN_CRF"]

    n_matched = sum(1 for r in mapped if r["status"] in ("matched", "weak_match"))
    summary = {
        "total_sap_variables": len(sap_vars),
        "total_crf_variables": len(scaffold_vars),
        "matched": n_matched,
        "missing_in_crf": len(gaps),
        "unreferenced_in_sap": len(unreferenced),
        "coverage_pct": round(n_matched / max(len(sap_vars), 1) * 100, 1),
    }

    return {
        "variables": mapped,
        "gaps": gaps,
        "unreferenced": unreferenced,
        "summary": summary,
    }


# ============================================================
# Core: SPIRIT 2025 Checklist Validation
# ============================================================

def validate_spirit_2025(scaffold_config=None, sap_config=None):
    """Validate protocol configs against SPIRIT 2025 checklist (50 sub-items).

    Checks if each item has relevant data in scaffold/SAP configs.
    Items requiring narrative text (background, rationale) are marked as
    'needs_narrative' since they can't be auto-validated from config alone.

    Returns dict:
        items: list of {id, name, section, status, evidence, note}
        score: {addressed, total, percentage}
        missing: list of unaddressed items
        needs_narrative: items requiring human review
    """
    results = []

    for item in SPIRIT_2025_CHECKLIST:
        status = "not_addressed"
        evidence = []
        note = ""

        # Check if any config_paths have data
        for config_source, path in item.get("config_paths", []):
            cfg = None
            if config_source == "scaffold" and scaffold_config:
                cfg = scaffold_config
            elif config_source == "sap" and sap_config:
                cfg = sap_config

            if cfg:
                val = _deep_get(cfg, path)
                if val is not None and val != "" and val != []:
                    evidence.append(f"{config_source}.{path}")
                    status = "addressed"

        # Items with no config_paths → needs human/narrative input
        if not item.get("config_paths"):
            status = "needs_narrative"
            note = "Requires narrative text in protocol document (not derivable from config)"

        # Special checks for items that need deeper validation
        if item["id"] == "19" and sap_config:
            po = sap_config.get("primary_outcome", {})
            if all(k in po for k in ["mcid", "sd", "alpha", "power"]):
                status = "addressed"
                evidence.append("sap.primary_outcome (complete sample size params)")

        if item["id"] == "14a" and scaffold_config:
            if scaffold_config.get("inclusion") and scaffold_config.get("exclusion"):
                status = "addressed"
                evidence.append(f"scaffold.inclusion ({len(scaffold_config['inclusion'])} criteria)")
                evidence.append(f"scaffold.exclusion ({len(scaffold_config['exclusion'])} criteria)")

        results.append({
            "id": item["id"],
            "name": item["name"],
            "section": item["section"],
            "description": item["description"],
            "status": status,
            "evidence": evidence,
            "note": note,
        })

    addressed = [r for r in results if r["status"] == "addressed"]
    missing = [r for r in results if r["status"] == "not_addressed"]
    needs_narrative = [r for r in results if r["status"] == "needs_narrative"]

    score = {
        "addressed": len(addressed),
        "not_addressed": len(missing),
        "needs_narrative": len(needs_narrative),
        "total": len(results),
        "auto_checkable": len(results) - len(needs_narrative),
        "percentage": round(len(addressed) / max(len(results) - len(needs_narrative), 1) * 100, 1),
    }

    return {
        "items": results,
        "score": score,
        "missing": missing,
        "needs_narrative": needs_narrative,
    }


# ============================================================
# Docx Report Generation
# ============================================================

def _add_table_row(table, cells_data, bold=False, header=False):
    """Add a row to a docx table."""
    row = table.add_row()
    for i, text in enumerate(cells_data):
        cell = row.cells[i]
        p = cell.paragraphs[0]
        run = p.add_run(str(text))
        run.font.size = Pt(9)
        if bold or header:
            run.bold = True
        if header:
            from lxml import etree
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            tc = cell._tc
            tcPr = tc.find(f"{{{W}}}tcPr")
            if tcPr is None:
                tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
                tc.insert(0, tcPr)
            shd = etree.SubElement(tcPr, f"{{{W}}}shd")
            shd.set(f"{{{W}}}fill", "2F5496")
            shd.set(f"{{{W}}}val", "clear")
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return row


def generate_full_report(scaffold_config, sap_config, output_path="consistency_report.docx"):
    """Generate complete Protocol ↔ CRF Consistency Report as docx.

    Includes:
    1. Mapping Matrix (SAP ↔ CRF cross-reference)
    2. Gap Analysis (missing variables)
    3. SPIRIT 2025 Checklist Validation
    4. Recommendations
    """
    if not HAS_DOCX:
        raise ImportError("python-docx is required. Install: pip install python-docx")

    doc = Document()

    # ── Title ──
    study_name = (scaffold_config.get("study", {}).get("short_name", "")
                  or sap_config.get("study", {}).get("title", "Study")[:40])

    title = doc.add_heading(f"Protocol ↔ CRF Consistency Report", level=0)
    doc.add_paragraph(f"Study: {study_name}")
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"Tool: protocol_crf_mapper.py (pytool Phase 2)")
    doc.add_paragraph("")

    # ── Part 1: Mapping Matrix ──
    doc.add_heading("1. Protocol ↔ CRF Mapping Matrix", level=1)

    matrix = generate_mapping_matrix(scaffold_config, sap_config)
    summary = matrix["summary"]

    # Summary paragraph
    p = doc.add_paragraph()
    p.add_run("Coverage Summary: ").bold = True
    p.add_run(
        f"{summary['matched']}/{summary['total_sap_variables']} SAP variables "
        f"mapped to CRF ({summary['coverage_pct']}%). "
        f"{summary['missing_in_crf']} gaps found. "
        f"{summary['unreferenced_in_sap']} CRF variables not in SAP."
    )

    # Mapping table
    doc.add_heading("1.1 SAP → CRF Variable Mapping", level=2)
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"

    # Header
    headers = ["SAP Variable", "Role", "CRF Variable", "CRF Form", "Source", "Status"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # Blue background
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        tc = cell._tc
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
            tc.insert(0, tcPr)
        shd = etree.SubElement(tcPr, f"{{{W}}}shd")
        shd.set(f"{{{W}}}fill", "2F5496")
        shd.set(f"{{{W}}}val", "clear")

    # Status symbols
    status_map = {"matched": "✅", "weak_match": "⚠️", "MISSING_IN_CRF": "❌"}

    for var in matrix["variables"]:
        row = table.add_row()
        cells = [
            var["sap_variable"],
            var["sap_role"],
            var["crf_variable"],
            var["crf_form"],
            var["data_source"],
            status_map.get(var["status"], var["status"]),
        ]
        for i, text in enumerate(cells):
            run = row.cells[i].paragraphs[0].add_run(str(text))
            run.font.size = Pt(8)
            if var["status"] == "MISSING_IN_CRF":
                run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

    # Gap report
    if matrix["gaps"]:
        doc.add_heading("1.2 Gap Analysis: SAP Variables Missing from CRF", level=2)
        for gap in matrix["gaps"]:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(f"❌ {gap['sap_variable']}")
            run.bold = True
            run.font.size = Pt(10)
            p.add_run(f" ({gap['sap_role']})")

    # Unreferenced CRF variables
    if matrix["unreferenced"]:
        doc.add_heading("1.3 CRF Variables Not Referenced in SAP", level=2)
        p = doc.add_paragraph(
            "These variables are collected in the CRF but not explicitly "
            "referenced in the SAP. This may be intentional (e.g., descriptive "
            "data, safety monitoring) or indicate a gap in the analysis plan."
        )
        p.runs[0].font.size = Pt(9)

        table2 = doc.add_table(rows=1, cols=4)
        table2.style = "Table Grid"
        for i, h in enumerate(["Variable", "Section", "Form", "Source"]):
            run = table2.rows[0].cells[i].paragraphs[0].add_run(h)
            run.bold = True
            run.font.size = Pt(9)

        for uv in matrix["unreferenced"][:30]:  # Limit to 30 rows
            row = table2.add_row()
            for i, text in enumerate([
                uv["crf_variable"], uv["crf_section"],
                uv["crf_form"], uv["data_source"]
            ]):
                run = row.cells[i].paragraphs[0].add_run(str(text))
                run.font.size = Pt(8)

    # ── Part 2: SPIRIT 2025 Checklist ──
    doc.add_page_break()
    doc.add_heading("2. SPIRIT 2025 Checklist Validation", level=1)

    spirit = validate_spirit_2025(scaffold_config, sap_config)
    score = spirit["score"]

    p = doc.add_paragraph()
    p.add_run("Validation Score: ").bold = True
    p.add_run(
        f"{score['addressed']}/{score['auto_checkable']} auto-checkable items addressed "
        f"({score['percentage']}%). "
        f"{score['needs_narrative']} items require narrative review."
    )

    # Checklist table
    doc.add_heading("2.1 Item-by-Item Assessment", level=2)
    table3 = doc.add_table(rows=1, cols=5)
    table3.style = "Table Grid"
    for i, h in enumerate(["Item", "Name", "Section", "Status", "Evidence"]):
        run = table3.rows[0].cells[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        from lxml import etree
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        tc = table3.rows[0].cells[i]._tc
        tcPr = tc.find(f"{{{W}}}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(tc, f"{{{W}}}tcPr")
            tc.insert(0, tcPr)
        shd = etree.SubElement(tcPr, f"{{{W}}}shd")
        shd.set(f"{{{W}}}fill", "2F5496")
        shd.set(f"{{{W}}}val", "clear")

    spirit_status_map = {
        "addressed": "✅ Addressed",
        "not_addressed": "❌ Missing",
        "needs_narrative": "📝 Narrative",
    }

    for item in spirit["items"]:
        row = table3.add_row()
        evidence_text = "; ".join(item["evidence"][:2]) if item["evidence"] else item.get("note", "—")
        cells = [
            item["id"],
            item["name"],
            item["section"],
            spirit_status_map.get(item["status"], item["status"]),
            evidence_text,
        ]
        for i, text in enumerate(cells):
            run = row.cells[i].paragraphs[0].add_run(str(text))
            run.font.size = Pt(8)
            if item["status"] == "not_addressed":
                run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

    # Missing items summary
    if spirit["missing"]:
        doc.add_heading("2.2 Action Required: Missing Items", level=2)
        for m in spirit["missing"]:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(f"Item {m['id']} — {m['name']}: ")
            run.bold = True
            p.add_run(m["description"])

    # ── Part 3: Recommendations ──
    doc.add_page_break()
    doc.add_heading("3. Recommendations", level=1)

    # Auto-generate recommendations based on gaps
    recs = []

    if matrix["gaps"]:
        recs.append(
            f"Add {len(matrix['gaps'])} missing variable(s) to CRF: "
            + ", ".join(g["sap_variable"] for g in matrix["gaps"])
        )

    if spirit["missing"]:
        missing_ids = [m["id"] for m in spirit["missing"]]
        recs.append(
            f"Address {len(spirit['missing'])} missing SPIRIT 2025 item(s) in protocol: "
            + ", ".join(missing_ids)
        )

    if summary["unreferenced_in_sap"] > 10:
        recs.append(
            f"Review {summary['unreferenced_in_sap']} CRF variables not referenced "
            f"in SAP. Consider adding relevant ones to exploratory or descriptive analyses."
        )

    if score["needs_narrative"] > 10:
        recs.append(
            f"{score['needs_narrative']} SPIRIT items require narrative text. "
            f"Ensure these are addressed in the full protocol document."
        )

    if not recs:
        recs.append("No critical gaps found. Protocol and CRF are well-aligned.")

    for i, rec in enumerate(recs, 1):
        p = doc.add_paragraph(style="List Number")
        p.add_run(rec).font.size = Pt(10)

    doc.save(output_path)

    return {
        "output_path": output_path,
        "mapping_summary": summary,
        "spirit_score": score,
        "recommendations": recs,
    }


# ============================================================
# Print helpers (for terminal/markdown output)
# ============================================================

def print_mapping_summary(scaffold_config, sap_config):
    """Print mapping matrix summary to stdout."""
    matrix = generate_mapping_matrix(scaffold_config, sap_config)
    s = matrix["summary"]

    print("=" * 70)
    print("  Protocol ↔ CRF Mapping Matrix Summary")
    print("=" * 70)
    print(f"  SAP variables:           {s['total_sap_variables']}")
    print(f"  CRF variables:           {s['total_crf_variables']}")
    print(f"  Matched:                 {s['matched']} ({s['coverage_pct']}%)")
    print(f"  Missing in CRF:          {s['missing_in_crf']}")
    print(f"  Unreferenced in SAP:     {s['unreferenced_in_sap']}")
    print()

    status_sym = {"matched": "✅", "weak_match": "⚠️", "MISSING_IN_CRF": "❌"}
    print(f"{'SAP Variable':<45} {'Role':<12} {'CRF Match':<30} {'Status'}")
    print("-" * 100)
    for var in matrix["variables"]:
        sym = status_sym.get(var["status"], "?")
        print(f"{var['sap_variable']:<45} {var['sap_role']:<12} {var['crf_variable']:<30} {sym}")

    if matrix["gaps"]:
        print(f"\n⚠️  GAPS ({len(matrix['gaps'])}):")
        for g in matrix["gaps"]:
            print(f"  ❌ {g['sap_variable']} ({g['sap_role']})")

    print()


def print_spirit_summary(scaffold_config, sap_config):
    """Print SPIRIT 2025 validation summary to stdout."""
    result = validate_spirit_2025(scaffold_config, sap_config)
    sc = result["score"]

    print("=" * 70)
    print("  SPIRIT 2025 Checklist Validation")
    print("=" * 70)
    print(f"  Auto-checkable items:    {sc['auto_checkable']}/{sc['total']}")
    print(f"  Addressed:               {sc['addressed']} ({sc['percentage']}%)")
    print(f"  Not addressed:           {sc['not_addressed']}")
    print(f"  Needs narrative review:  {sc['needs_narrative']}")
    print()

    status_sym = {"addressed": "✅", "not_addressed": "❌", "needs_narrative": "📝"}
    current_section = ""
    for item in result["items"]:
        if item["section"] != current_section:
            current_section = item["section"]
            print(f"\n  [{current_section}]")
        sym = status_sym.get(item["status"], "?")
        ev = " — " + "; ".join(item["evidence"][:1]) if item["evidence"] else ""
        print(f"    {sym} {item['id']:>4} {item['name']:<35}{ev}")

    if result["missing"]:
        print(f"\n  ❌ ACTION REQUIRED ({len(result['missing'])} items):")
        for m in result["missing"]:
            print(f"    • Item {m['id']} — {m['name']}: {m['description']}")
    print()


# ============================================================
# Example configs (import from Phase 1 tools)
# ============================================================

# Minimal inline configs for standalone testing.
# For full configs, import from clinical_research_scaffold and sap_generator.

EXAMPLE_SCAFFOLD_CONFIG = {
    "study": {
        "short_name": "FSHD",
        "file_prefix": "FSHD",
        "title_ko": "동결견 환자의 초음파 유도 관절낭 보존 수압팽창술",
        "title_en": "US-guided capsule-preserving hydrodilatation for adhesive capsulitis",
        "department": "부산대학교병원 재활의학과",
        "pi": "이재현",
        "preexam_title": "어깨 통증 예진 문진표",
        "quicksheet_title": "FSHD Ultrasound Quick Sheet",
    },
    "subjects": {
        "fields": ["이름", "등록번호", "날짜"],
        "crf_fields": ["Subject ID", "병록번호(MRN)", "이니셜", "성명", "동의일"],
    },
    "enrollment": [
        {"label": "성별", "hint": "M(0) / F(1)", "source": "emr"},
        {"label": "생년월일 / 나이", "hint": "____.__.__  /  만 ___세", "source": "emr"},
        {"label": "이환측", "hint": "Rt(0) / Lt(1)", "source": "patient"},
        {"label": "우세손", "hint": "Rt(0) / Lt(1)", "source": "patient"},
        {"label": "키 (cm)", "hint": "___ cm", "source": "patient"},
        {"label": "몸무게 (kg)", "hint": "___ kg", "source": "patient"},
        {"label": "BMI", "hint": "자동계산", "source": "calc"},
    ],
    "inclusion": [
        "40–75세",
        "Active + Passive ROM 제한 ≥ 30° (2방향 이상)",
        "Primary adhesive capsulitis (이차성 원인 배제)",
        "증상 기간 ≥ 4주",
    ],
    "exclusion": [
        "이차성 원인 (골절, 수술, 석회건염 등)",
        "신경학적 결손 (상완신경총 등)",
        "최근 3개월 이내 스테로이드 주사",
        "전신 스테로이드 사용 중",
        "조절되지 않는 당뇨 (HbA1c > 9%)",
    ],
    "baseline": {
        "sections": [
            {
                "name": "동결견 병력",
                "items": [
                    {"label": "증상 시작일", "hint": "____년 ____월", "source": "patient"},
                    {"label": "이환 기간", "hint": "자동계산 (____개월)", "source": "calc"},
                    {"label": "Phase", "hint": "1 Freezing / 2 Frozen / 3 Thawing", "source": "examiner"},
                    {"label": "이전 치료 (주사 외)", "hint": "PT / 약물 / 도수 / 기타", "source": "patient"},
                    {"label": "이전 주사 횟수", "hint": "__회, 최종: ____년 ____월", "source": "patient"},
                    {"label": "당뇨 여부", "hint": "없음(0) / 있음(1)", "source": "patient"},
                ],
            },
            {
                "name": "Pain NRS (/10)",
                "items": [
                    {"label": "Resting pain", "hint": "___ /10", "source": "patient"},
                    {"label": "Activity pain", "hint": "___ /10", "source": "patient"},
                    {"label": "Night pain", "hint": "___ /10", "source": "patient"},
                ],
            },
            {
                "name": "Passive ROM (Goniometer)",
                "items": [
                    {"label": "Forward Flexion", "hint": "Affected ___° / Unaffected ___°", "source": "examiner"},
                    {"label": "Abduction", "hint": "Affected ___° / Unaffected ___°", "source": "examiner"},
                    {"label": "External Rotation", "hint": "Affected ___° / Unaffected ___°", "source": "examiner"},
                    {"label": "Extension", "hint": "Affected ___° / Unaffected ___°", "source": "examiner"},
                    {"label": "Internal Rotation", "hint": "Affected ___ / Unaffected ___", "source": "examiner"},
                ],
            },
            {
                "name": "SPADI + PGA",
                "items": [
                    {"label": "SPADI Pain (P1–P5)", "hint": "___ /50", "source": "patient"},
                    {"label": "SPADI Disability (D1–D8)", "hint": "___ /80", "source": "patient"},
                    {"label": "SPADI Total (%)", "hint": "자동계산", "source": "calc"},
                    {"label": "Patient Global Assessment", "hint": "___ /10", "source": "patient"},
                ],
            },
        ],
    },
    "procedure": {
        "sections": [
            {
                "name": "Ultrasound Baseline",
                "items": [
                    {"label": "CHL thickness (mm)", "hint": "Affected ___mm / Unaffected ___mm", "source": "examiner"},
                    {"label": "Axillary recess (mm)", "hint": "Affected ___mm / Unaffected ___mm", "source": "examiner"},
                    {"label": "SST status", "hint": "Intact / Partial / Full tear", "source": "examiner"},
                    {"label": "RI thickness (mm)", "hint": "Affected ___mm / Unaffected ___mm", "source": "examiner"},
                    {"label": "Power Doppler", "hint": "+  /  −", "source": "examiner"},
                    {"label": "Image file no.", "hint": "________", "source": "examiner"},
                ],
            },
            {
                "name": "Intra-Procedure",
                "items": [
                    {"label": "Approach", "hint": "Anterior(0) / Posterior(1)", "source": "procedure"},
                    {"label": "5 mL", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                    {"label": "10 mL", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                    {"label": "15 mL", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                    {"label": "20 mL", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                    {"label": "25 mL", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                    {"label": "30 mL", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                    {"label": "__ mL (custom)", "hint": "Time / Resistance(0-3) / Discomfort / US", "source": "procedure"},
                ],
            },
            {
                "name": "Procedure End",
                "items": [
                    {"label": "Total injected volume", "hint": "___ mL", "source": "procedure"},
                    {"label": "Total injection time", "hint": "___ sec", "source": "procedure"},
                    {"label": "Stopping reason", "hint": "1 Rupture / 2 Patient / 3 Resistance / 4 Plateau / 5 Target", "source": "procedure"},
                    {"label": "Capsule status", "hint": "Preserved(0) / Ruptured(1)", "source": "procedure"},
                    {"label": "Reinjection", "hint": "No(0) / Yes(1) → 추가 ___mL", "source": "procedure"},
                    {"label": "Rupture location", "hint": "Ant / Post / Inf / Sup / N/A", "source": "procedure"},
                ],
            },
            {
                "name": "AE Observation (15–30 min)",
                "items": [
                    {"label": "Vasovagal reaction", "hint": "No(0) / Yes(1)", "source": "procedure"},
                    {"label": "Injection site pain", "hint": "None(0) / Mild(1) / Moderate(2) / Severe(3)", "source": "procedure"},
                    {"label": "Swelling / Ecchymosis", "hint": "No(0) / Yes(1)", "source": "procedure"},
                    {"label": "Other", "hint": "________________________", "source": "procedure"},
                ],
            },
        ],
    },
    "followup": {
        "visits": ["V1 Baseline", "V2 30min", "V3 2wk", "V4 4wk", "V5 8wk", "V6 12wk"],
        "measures": ["NRS (rest/act/night)", "ROM", "SPADI", "PGA", "AE check"],
    },
    "emr_post_entry": ["Subject ID", "성별", "생년월일", "시술일", "시술자"],
    "auto_calc": ["BMI", "이환기간", "SPADI Total %", "ROM ratio", "ROM change"],
}


EXAMPLE_SAP_CONFIG = {
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
        "itt": "All randomized subjects, analyzed according to assigned treatment group.",
        "pp": "All randomized subjects who completed 3-month follow-up without major deviations.",
        "safety": "All subjects who received at least one injection procedure.",
    },
    "analysis": {
        "primary_method": "Independent samples t-test with ANCOVA confirmatory.",
        "missing_data": "Multiple imputation (MICE) under MAR assumption.",
        "multiplicity": "Bonferroni correction for secondary endpoints.",
        "sensitivity": [
            "Per-protocol analysis (PP population)",
            "ANCOVA with baseline SPADI as covariate",
            "Complete-case analysis",
            "Worst-case imputation",
        ],
        "subgroup": [
            "Sex (male vs female)",
            "Diabetes mellitus (yes vs no)",
            "Dominant hand affected (yes vs no)",
            "Symptom duration (3-6 months vs >6 months)",
        ],
        "software": "Python 3.x (scipy, statsmodels, pandas)",
        "significance_level": "Two-sided α = 0.05",
    },
}


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Protocol ↔ CRF Mapping Matrix + SPIRIT 2025 Validator"
    )
    parser.add_argument("scaffold_config", nargs="?", help="Scaffold config JSON file")
    parser.add_argument("sap_config", nargs="?", help="SAP config JSON file")
    parser.add_argument("--output", "-o", default="consistency_report.docx",
                        help="Output docx path")
    parser.add_argument("--example", action="store_true",
                        help="Print example configs as JSON")
    parser.add_argument("--run", action="store_true",
                        help="Run with example configs (use with --example)")
    parser.add_argument("--summary", action="store_true",
                        help="Print text summary only (no docx)")

    args = parser.parse_args()

    if args.example and not args.run:
        print(json.dumps({
            "scaffold_config": EXAMPLE_SCAFFOLD_CONFIG,
            "sap_config": EXAMPLE_SAP_CONFIG,
        }, indent=2, ensure_ascii=False))
        return

    # Load configs
    if args.example and args.run:
        scaffold_cfg = EXAMPLE_SCAFFOLD_CONFIG
        sap_cfg = EXAMPLE_SAP_CONFIG
    elif args.scaffold_config and args.sap_config:
        with open(args.scaffold_config) as f:
            scaffold_cfg = json.load(f)
        with open(args.sap_config) as f:
            sap_cfg = json.load(f)
    else:
        parser.print_help()
        print("\nUse --example --run to test with FSHD example configs.")
        sys.exit(1)

    # Print summaries
    print_mapping_summary(scaffold_cfg, sap_cfg)
    print_spirit_summary(scaffold_cfg, sap_cfg)

    # Generate docx report
    if not args.summary:
        result = generate_full_report(scaffold_cfg, sap_cfg, args.output)
        print(f"\n📄 Report saved: {result['output_path']}")
        print(f"   Mapping: {result['mapping_summary']['matched']}/"
              f"{result['mapping_summary']['total_sap_variables']} matched "
              f"({result['mapping_summary']['coverage_pct']}%)")
        print(f"   SPIRIT:  {result['spirit_score']['addressed']}/"
              f"{result['spirit_score']['auto_checkable']} addressed "
              f"({result['spirit_score']['percentage']}%)")


if __name__ == "__main__":
    main()

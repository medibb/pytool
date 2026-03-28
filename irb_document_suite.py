#!/usr/bin/env python3
"""
irb_document_suite.py — IRB 서류 일괄 생성기

Phase 4 of the clinical research workflow:
  Config dict → IRB 제출에 필요한 전체 서류 세트 자동 생성
  - 연구 요약서 (Study Summary)
  - 동의서 (Informed Consent Form)
  - IRB 제출 체크리스트

기존 pytool과 통합:
  - clinical_research_scaffold.py → CRF, 예진지, Quick Sheet (이미 구현)
  - sap_generator.py → SAP (이미 구현)
  → 이 도구 추가로 Full IRB Suite 완성

=== Usage ===

Python:
    from irb_document_suite import generate_irb_suite, EXAMPLE_CONFIG

    generate_irb_suite(EXAMPLE_CONFIG, output_dir="./irb_docs")
    # → study_summary.docx, informed_consent.docx, irb_checklist.docx

CLI:
    python3 irb_document_suite.py --example --run
    python3 irb_document_suite.py config.json -o ./irb_docs

Dependencies: python-docx
"""

import os
import sys
import json
import argparse
from datetime import datetime

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ============================================================
# Document 1: Study Summary (연구 요약서)
# ============================================================

def generate_study_summary(config, output_path="study_summary.docx"):
    """Generate 1-2 page study summary for IRB submission.

    Covers: title, PI, background, objectives, design, population,
    intervention, outcomes, sample size, timeline.
    """
    if not HAS_DOCX:
        raise ImportError("python-docx required.")

    doc = Document()
    study = config.get("study", {})

    # Title
    title_p = doc.add_heading("연구 요약서 (Study Summary)", level=0)
    doc.add_paragraph("")

    # Basic info table
    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = "Table Grid"
    info_data = [
        ("연구 제목", study.get("title_ko", study.get("title", ""))),
        ("영문 제목", study.get("title_en", study.get("title", ""))),
        ("책임연구자", f"{study.get('pi', '')} ({study.get('department', '')})"),
        ("연구 기간", study.get("study_period", "IRB 승인일 ~ 종료일")),
        ("프로토콜 번호", study.get("protocol_number", "")),
        ("연구비 지원", study.get("funding", "해당 없음")),
    ]
    for i, (label, value) in enumerate(info_data):
        row = info_table.rows[i]
        run = row.cells[0].paragraphs[0].add_run(label)
        run.bold = True
        run.font.size = Pt(10)
        run2 = row.cells[1].paragraphs[0].add_run(value)
        run2.font.size = Pt(10)

    doc.add_paragraph("")

    # Sections
    sections = [
        ("1. 연구 배경 및 필요성", config.get("background", "연구 배경을 기술하세요.")),
        ("2. 연구 목적", _format_objectives(config)),
        ("3. 연구 설계", _format_design(config)),
        ("4. 연구 대상", _format_population(config)),
        ("5. 중재 (시술) 방법", config.get("intervention_description",
                              "시술 방법을 상세히 기술하세요.")),
        ("6. 평가 변수", _format_outcomes(config)),
        ("7. 표본 크기", _format_sample_size(config)),
        ("8. 통계 분석", config.get("analysis_summary",
                        "1차 변수: 독립표본 t-검정, 2차 변수: 보정 분석 (ANCOVA)")),
        ("9. 연구 일정", _format_timeline(config)),
        ("10. 윤리적 고려사항", _format_ethics(config)),
    ]

    for heading, content in sections:
        doc.add_heading(heading, level=2)
        if isinstance(content, list):
            for item in content:
                doc.add_paragraph(item, style="List Bullet")
        else:
            doc.add_paragraph(content)

    doc.save(output_path)
    return output_path


def _format_objectives(config):
    """Format objectives section."""
    lines = []
    obj = config.get("objectives", {})
    if obj.get("primary"):
        lines.append(f"주요 목적: {obj['primary']}")
    for i, sec in enumerate(obj.get("secondary", []), 1):
        lines.append(f"부차 목적 {i}: {sec}")
    return lines if lines else "연구 목적을 기술하세요."


def _format_design(config):
    """Format study design section."""
    design = config.get("design", {})
    parts = []
    dtype = design.get("type", "")
    type_map = {
        "rct_parallel": "무작위 배정 평행군 비교 임상시험",
        "rct_crossover": "무작위 배정 교차 임상시험",
        "cohort": "전향적 코호트 연구",
        "case_control": "환자-대조군 연구",
        "cross_sectional": "단면 연구",
    }
    parts.append(f"연구 유형: {type_map.get(dtype, dtype)}")

    arms = design.get("arms", [])
    if arms:
        arm_names = ", ".join(a.get("name", "") for a in arms)
        parts.append(f"연구군: {arm_names}")

    blinding = design.get("blinding", "")
    blinding_map = {
        "open_label": "개방형 (open-label)",
        "single_blind_assessor": "단일 맹검 (평가자 맹검)",
        "double_blind": "이중 맹검",
    }
    parts.append(f"맹검: {blinding_map.get(blinding, blinding)}")

    rand = design.get("randomization", {})
    if rand:
        parts.append(f"무작위 배정: {rand.get('method', '')} 방법, "
                     f"블록 크기 {rand.get('block_sizes', [])}")
        if rand.get("stratification"):
            parts.append(f"층화 요인: {', '.join(rand['stratification'])}")

    return "\n".join(parts)


def _format_population(config):
    """Format population section."""
    lines = []
    inclusion = config.get("inclusion", [])
    if inclusion:
        lines.append("선정 기준:")
        for i, c in enumerate(inclusion, 1):
            lines.append(f"  {i}) {c}")
    exclusion = config.get("exclusion", [])
    if exclusion:
        lines.append("제외 기준:")
        for i, c in enumerate(exclusion, 1):
            lines.append(f"  {i}) {c}")
    return lines if lines else "연구 대상 기준을 기술하세요."


def _format_outcomes(config):
    """Format outcomes section."""
    lines = []
    po = config.get("primary_outcome", {})
    if po:
        lines.append(f"1차 평가 변수: {po.get('name', '')}")
    for i, so in enumerate(config.get("secondary_outcomes", []), 1):
        name = so.get("name", so) if isinstance(so, dict) else so
        lines.append(f"2차 평가 변수 {i}: {name}")
    return lines if lines else "평가 변수를 기술하세요."


def _format_sample_size(config):
    """Format sample size section."""
    po = config.get("primary_outcome", {})
    if not po.get("mcid"):
        return "표본 크기 산출 근거를 기술하세요."

    parts = [
        f"1차 변수: {po.get('name', '')}",
        f"최소 임상적 유의차 (MCID): {po.get('mcid', '')}",
        f"표준편차 (SD): {po.get('sd', '')}",
        f"유의수준 α = {po.get('alpha', 0.05)}, 검정력 = {po.get('power', 0.80)}",
        f"탈락률: {po.get('dropout_rate', 0) * 100:.0f}%",
    ]

    # Calculate sample size if possible
    try:
        from scipy.stats import norm
        import math
        mcid = po["mcid"]
        sd = po["sd"]
        alpha = po.get("alpha", 0.05)
        power = po.get("power", 0.80)
        z_alpha = norm.ppf(1 - alpha / 2)
        z_beta = norm.ppf(power)
        n = 2 * ((z_alpha + z_beta) * sd / mcid) ** 2
        n = math.ceil(n)
        dropout = po.get("dropout_rate", 0)
        n_adj = math.ceil(n / (1 - dropout)) if dropout > 0 else n
        total = n_adj * 2
        parts.append(f"산출 결과: 군당 {n}명 → 탈락 보정 군당 {n_adj}명 → 총 {total}명")
    except Exception:
        parts.append("(sap_generator.py로 정확한 산출 가능)")

    return "\n".join(parts)


def _format_timeline(config):
    """Format timeline section."""
    timeline = config.get("timeline", {})
    if not timeline:
        return ("IRB 심의: 승인 후 1개월 내 연구 시작\n"
                "대상자 모집: 승인 후 6개월\n"
                "자료 분석: 모집 완료 후 3개월\n"
                "논문 작성: 분석 완료 후 6개월")
    parts = []
    for phase, period in timeline.items():
        parts.append(f"{phase}: {period}")
    return "\n".join(parts)


def _format_ethics(config):
    """Format ethics section."""
    return ("본 연구는 헬싱키 선언의 윤리적 원칙에 따라 수행됩니다.\n"
            "연구 참여 전 서면 동의를 받으며, 참여자는 언제든지 동의를 철회할 수 있습니다.\n"
            "수집된 개인정보는 연구 목적으로만 사용되며, 관련 법규에 따라 보호됩니다.\n"
            "모든 연구 자료는 잠금장치가 있는 장소에 보관하고, 전자 자료는 암호화합니다.")


# ============================================================
# Document 2: Informed Consent Form (동의서)
# ============================================================

def generate_informed_consent(config, output_path="informed_consent.docx"):
    """Generate informed consent form for IRB.

    Based on standard Korean IRB template (부산대학교병원 양식).
    """
    if not HAS_DOCX:
        raise ImportError("python-docx required.")

    doc = Document()
    study = config.get("study", {})
    title_ko = study.get("title_ko", study.get("title", "연구 제목"))
    pi = study.get("pi", "연구책임자")
    dept = study.get("department", "부서명")

    # Header
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("연구참여 동의서")
    run.bold = True
    run.font.size = Pt(16)

    doc.add_paragraph("")

    # Study title
    p = doc.add_paragraph()
    run = p.add_run("연구 제목: ")
    run.bold = True
    p.add_run(title_ko)
    doc.add_paragraph("")

    # Consent sections (standard Korean IRB format)
    consent_sections = [
        ("1. 연구의 목적",
         config.get("objectives", {}).get("primary",
         "이 연구는 [연구 목적]을 알아보기 위해 수행됩니다.")),

        ("2. 연구 참여 대상",
         f"선정 기준:\n" +
         "\n".join(f"  - {c}" for c in config.get("inclusion", ["선정 기준을 기술"])) +
         f"\n\n제외 기준:\n" +
         "\n".join(f"  - {c}" for c in config.get("exclusion", ["제외 기준을 기술"]))),

        ("3. 연구 방법 (시술/검사 내용)",
         config.get("intervention_description",
         "연구에 참여하시면 다음과 같은 시술/검사를 받게 됩니다:\n"
         "[시술 방법을 환자가 이해할 수 있는 쉬운 말로 기술]")),

        ("4. 연구 참여 기간 및 방문 횟수",
         _format_visit_schedule(config)),

        ("5. 예상되는 위험 및 부작용",
         config.get("risks",
         "이 연구에 참여함으로써 다음과 같은 위험이 있을 수 있습니다:\n"
         "- [예상 부작용 1]\n- [예상 부작용 2]\n"
         "예상하지 못한 부작용이 발생할 경우 즉시 연구팀에 알려주십시오.")),

        ("6. 예상되는 이익",
         config.get("benefits",
         "이 연구에 참여하여 직접적인 이익을 얻을 수도 있으나, "
         "그렇지 않을 수도 있습니다. 그러나 귀하의 참여가 향후 "
         "같은 질환을 가진 환자들의 치료 발전에 기여할 수 있습니다.")),

        ("7. 연구 참여에 대한 보상",
         config.get("compensation",
         "본 연구 참여에 대한 별도의 금전적 보상은 없습니다. "
         "연구 관련 검사 비용은 연구비에서 지원됩니다.")),

        ("8. 개인정보 보호",
         "귀하의 개인정보는 관련 법규(개인정보보호법, 생명윤리법)에 따라 보호됩니다. "
         "연구 결과가 발표될 때에도 귀하의 신원을 파악할 수 있는 정보는 사용되지 않습니다. "
         "연구 관련 모니터링, 점검, 심사 시 관련 기관에 한하여 자료 열람이 가능하며, "
         "이 경우에도 비밀은 보장됩니다."),

        ("9. 동의 철회",
         "귀하는 자유 의사에 따라 연구 참여에 동의하거나 거부할 수 있습니다. "
         "또한 연구 도중 언제든지 동의를 철회할 수 있으며, 이 경우 불이익은 전혀 없습니다. "
         "동의 철회 시 이미 수집된 자료는 연구에 사용될 수 있습니다."),

        ("10. 연구 관련 손상에 대한 보상",
         "연구 참여로 인해 발생한 유해사례에 대해서는 적절한 치료를 받으실 수 있습니다."),

        ("11. 문의처",
         f"연구에 대해 궁금한 점이 있으시면 아래로 연락해 주십시오.\n\n"
         f"연구책임자: {pi}\n"
         f"소속: {dept}\n"
         f"연락처: {study.get('contact', '000-000-0000')}\n\n"
         f"연구대상자의 권리에 대해 궁금한 점이 있으시면 "
         f"부산대학교병원 임상시험심사위원회(IRB)로 연락해 주십시오.\n"
         f"IRB 전화: 051-240-7476"),
    ]

    for heading, content in consent_sections:
        doc.add_heading(heading, level=2)
        for para_text in content.split("\n"):
            if para_text.strip():
                doc.add_paragraph(para_text.strip())

    # Signature section
    doc.add_paragraph("")
    doc.add_paragraph("─" * 40)
    doc.add_paragraph("")

    sig_text = (
        "본인은 위의 내용을 충분히 이해하였으며, 자발적으로 이 연구에 참여하는 것에 "
        "동의합니다. 본인은 이 동의서의 사본을 받았습니다."
    )
    p = doc.add_paragraph(sig_text)
    doc.add_paragraph("")

    # Signature table
    sig_table = doc.add_table(rows=4, cols=3)
    sig_table.style = "Table Grid"
    sig_data = [
        ("", "서명 / 날짜", "이름"),
        ("연구대상자", "________________  /  ____.__.__", "________________"),
        ("법정대리인\n(필요 시)", "________________  /  ____.__.__", "________________"),
        ("동의 취득자\n(연구자)", "________________  /  ____.__.__", "________________"),
    ]
    for i, (role, sig, name) in enumerate(sig_data):
        row = sig_table.rows[i]
        for j, text in enumerate([role, sig, name]):
            run = row.cells[j].paragraphs[0].add_run(text)
            run.font.size = Pt(10)
            if i == 0:
                run.bold = True

    doc.save(output_path)
    return output_path


def _format_visit_schedule(config):
    """Format visit schedule for consent form."""
    followup = config.get("followup", {})
    visits = followup.get("visits", [])
    if visits:
        lines = [f"본 연구의 총 방문 횟수는 {len(visits)}회이며, 각 방문은 다음과 같습니다:"]
        for v in visits:
            lines.append(f"  - {v}")
        measures = followup.get("measures", [])
        if measures:
            lines.append(f"\n각 방문 시 다음 검사를 시행합니다: {', '.join(measures)}")
        return "\n".join(lines)
    return ("연구 기간 동안 총 [  ]회 방문이 예정되어 있습니다.\n"
            "[각 방문 일정 및 검사 내용을 기술]")


# ============================================================
# Document 3: IRB Submission Checklist (IRB 제출 체크리스트)
# ============================================================

def generate_irb_checklist(config, output_path="irb_checklist.docx"):
    """Generate IRB submission checklist docx.

    Standard PNUH IRB checklist items with auto-filled status.
    """
    if not HAS_DOCX:
        raise ImportError("python-docx required.")

    doc = Document()
    study = config.get("study", {})

    doc.add_heading("IRB 제출 서류 체크리스트", level=0)
    doc.add_paragraph(f"연구 제목: {study.get('title_ko', study.get('title', ''))}")
    doc.add_paragraph(f"책임연구자: {study.get('pi', '')}")
    doc.add_paragraph(f"작성일: {datetime.now().strftime('%Y-%m-%d')}")
    doc.add_paragraph("")

    # Checklist items
    items = [
        ("연구계획서 (Protocol)", _check_has(config, "objectives"),
         "protocol_crf_mapper.py + 수동 작성"),
        ("연구 요약서", True,
         "irb_document_suite.py → study_summary.docx"),
        ("동의서 (Informed Consent)", True,
         "irb_document_suite.py → informed_consent.docx"),
        ("증례기록서 (CRF)", _check_has(config, "baseline"),
         "clinical_research_scaffold.py → CRF.docx"),
        ("예진지 / 설문지", _check_has(config, "baseline"),
         "clinical_research_scaffold.py → 예진지.docx"),
        ("통계분석계획서 (SAP)", _check_has(config, "primary_outcome"),
         "sap_generator.py → SAP.docx"),
        ("연구자 이력서 (CV)", False, "수동 준비"),
        ("GCP 교육 이수증", False, "NECA/CITI 교육 이수 후 첨부"),
        ("보험 가입 증명서", False, "해당 시 첨부 (중재 연구)"),
        ("연구비 지원 증빙", False, "해당 시 첨부"),
        ("이해상충 신고서 (COI)", False, "서식 작성"),
        ("모집 공고문 (광고문)", False, "해당 시 작성"),
        ("연구 대상자 보호 계획", True, "동의서 내 포함"),
        ("개인정보 처리 동의", True, "동의서 내 포함"),
        ("임상시험 등록", _check_has(config, "study.protocol_number"),
         "ClinicalTrials.gov 또는 CRIS"),
    ]

    table = doc.add_table(rows=len(items) + 1, cols=4)
    table.style = "Table Grid"

    # Header
    headers = ["서류 항목", "준비 여부", "생성 도구", "비고"]
    for i, h in enumerate(headers):
        run = table.rows[0].cells[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(10)

    for row_idx, (item_name, is_ready, tool) in enumerate(items, 1):
        row = table.rows[row_idx]
        row.cells[0].paragraphs[0].add_run(item_name).font.size = Pt(9)

        status = "✅ 자동 생성" if is_ready else "⬜ 수동 준비"
        run = row.cells[1].paragraphs[0].add_run(status)
        run.font.size = Pt(9)
        if not is_ready:
            run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

        row.cells[2].paragraphs[0].add_run(tool).font.size = Pt(8)
        row.cells[3].paragraphs[0].add_run("").font.size = Pt(8)

    doc.add_paragraph("")
    doc.add_paragraph("※ ✅ 표시 항목은 pytool로 자동 생성 가능합니다.")
    doc.add_paragraph("※ ⬜ 표시 항목은 별도 준비가 필요합니다.")

    doc.save(output_path)
    return output_path


def _check_has(config, key):
    """Check if config has a non-empty value at the given key path."""
    keys = key.split(".")
    obj = config
    for k in keys:
        if isinstance(obj, dict) and k in obj:
            obj = obj[k]
        else:
            return False
    return bool(obj)


# ============================================================
# Full Suite Generator
# ============================================================

def generate_irb_suite(config, output_dir="./irb_docs"):
    """Generate complete IRB document suite.

    Creates output_dir/ with:
      - study_summary.docx
      - informed_consent.docx
      - irb_checklist.docx

    Returns dict with file paths and status.
    """
    os.makedirs(output_dir, exist_ok=True)
    results = {}

    print("📋 Generating IRB document suite...")

    # Study summary
    path = generate_study_summary(config, os.path.join(output_dir, "study_summary.docx"))
    results["study_summary"] = path
    print(f"  ✅ {path}")

    # Informed consent
    path = generate_informed_consent(config, os.path.join(output_dir, "informed_consent.docx"))
    results["informed_consent"] = path
    print(f"  ✅ {path}")

    # IRB checklist
    path = generate_irb_checklist(config, os.path.join(output_dir, "irb_checklist.docx"))
    results["irb_checklist"] = path
    print(f"  ✅ {path}")

    # Note about other documents from existing tools
    print("\n  💡 Complete IRB suite with existing pytool:")
    print("     clinical_research_scaffold.py → CRF, 예진지, Quick Sheet")
    print("     sap_generator.py → SAP")
    print("     protocol_crf_mapper.py → Protocol consistency report")

    return results


# ============================================================
# Example Config
# ============================================================

EXAMPLE_CONFIG = {
    "study": {
        "title": "Comparison of Anterior vs Posterior Approach Ultrasound-guided "
                 "Hydrodilatation for Frozen Shoulder",
        "title_ko": "동결견 환자의 초음파 유도 전방 접근법 vs 후방 접근법 "
                     "관절낭 보존 수압팽창술 비교: 무작위 대조 시험",
        "title_en": "Anterior vs Posterior Approach US-guided Capsule-preserving "
                    "Hydrodilatation for Adhesive Capsulitis: An RCT",
        "protocol_number": "PNUH-2026-IRB-XXX",
        "pi": "이재현",
        "department": "부산대학교병원 재활의학과",
        "contact": "051-240-XXXX",
        "study_period": "IRB 승인 후 12개월",
        "funding": "원내 연구비",
    },
    "background": (
        "동결견(adhesive capsulitis)은 견관절의 관절낭 섬유화와 수축으로 인해 "
        "심한 통증과 운동 범위 제한을 유발하는 질환입니다. "
        "초음파 유도 수압팽창술(hydrodilatation)은 관절낭에 생리식염수를 "
        "주입하여 관절낭을 확장시키는 효과적인 치료법으로 알려져 있습니다. "
        "현재 전방 접근법과 후방 접근법 모두 사용되고 있으나, "
        "두 접근법의 효과를 직접 비교한 연구는 제한적입니다."
    ),
    "objectives": {
        "primary": "동결견 환자에서 초음파 유도 수압팽창술의 전방 접근법과 "
                   "후방 접근법의 3개월 후 SPADI 총점 변화를 비교",
        "secondary": [
            "어깨 관절 운동 범위(ROM) 변화 비교",
            "통증 NRS 변화 비교",
            "환자 전반적 평가(PGA) 변화 비교",
        ],
    },
    "design": {
        "type": "rct_parallel",
        "arms": [
            {"name": "전방 접근법 수압팽창술", "ratio": 1},
            {"name": "후방 접근법 수압팽창술", "ratio": 1},
        ],
        "blinding": "single_blind_assessor",
        "randomization": {
            "method": "block",
            "block_sizes": [4, 6],
            "stratification": ["성별"],
        },
    },
    "inclusion": [
        "40–75세 성인",
        "Passive ROM 제한 ≥ 30° (2방향 이상)",
        "Primary adhesive capsulitis (이차성 원인 배제)",
        "증상 기간 ≥ 4주",
        "자발적 연구 동의",
    ],
    "exclusion": [
        "이차성 원인 (골절, 수술, 석회건염 등)",
        "신경학적 결손 (상완신경총 손상 등)",
        "최근 3개월 이내 스테로이드 관절 주사",
        "전신 스테로이드 복용 중",
        "조절되지 않는 당뇨 (HbA1c > 9%)",
        "항응고제 복용 중",
    ],
    "intervention_description": (
        "초음파 유도 하에 견관절 관절낭에 22G 척추침으로 접근합니다.\n"
        "전방 접근법: rotator interval 경유, 전방 관절낭 천자\n"
        "후방 접근법: infraspinatus 하방 경유, 후방 관절낭 천자\n"
        "생리식염수를 5mL 단위로 최대 30mL까지 주입하며, "
        "저항, 통증, 초음파 소견을 기록합니다."
    ),
    "risks": (
        "이 연구에 참여함으로써 다음과 같은 위험이 있을 수 있습니다:\n"
        "- 시술 부위 통증 (일시적, 대부분 24-48시간 내 호전)\n"
        "- 관절낭 파열 (시술 중 발생 가능, 임상적 의미 미미)\n"
        "- 미주신경 반사 (드물게 발생, 안정 시 호전)\n"
        "- 감염 (매우 드묾, 무균 시술로 예방)"
    ),
    "primary_outcome": {
        "name": "SPADI total score change at 3 months",
        "mcid": 13,
        "sd": 20,
        "alpha": 0.05,
        "power": 0.80,
        "dropout_rate": 0.20,
    },
    "secondary_outcomes": [
        {"name": "Shoulder ROM (ER, IR, ABD, Flexion)"},
        {"name": "Pain NRS (rest, activity, night)"},
        {"name": "PGA (patient global assessment)"},
    ],
    "followup": {
        "visits": [
            "V1: Baseline (시술 전)",
            "V2: 시술 직후 30분",
            "V3: 2주 후",
            "V4: 4주 후",
            "V5: 8주 후",
            "V6: 12주 후 (3개월)",
        ],
        "measures": ["Pain NRS", "ROM", "SPADI", "PGA", "이상반응 확인"],
    },
    "baseline": {"sections": [{"name": "dummy", "items": [{"label": "x"}]}]},
}


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="IRB Document Suite Generator"
    )
    parser.add_argument("config", nargs="?", help="Study config JSON file")
    parser.add_argument("--output", "-o", default="./irb_docs",
                        help="Output directory")
    parser.add_argument("--example", action="store_true", help="Use example config")
    parser.add_argument("--run", action="store_true", help="Generate documents")

    args = parser.parse_args()

    if args.example and not args.run:
        print(json.dumps(EXAMPLE_CONFIG, indent=2, ensure_ascii=False))
        return

    if args.example and args.run:
        cfg = EXAMPLE_CONFIG
    elif args.config:
        with open(args.config) as f:
            cfg = json.load(f)
    else:
        parser.print_help()
        print("\nUse --example --run to generate FSHD example documents.")
        sys.exit(1)

    results = generate_irb_suite(cfg, args.output)
    print(f"\n✅ IRB suite generated: {len(results)} documents in {args.output}/")


if __name__ == "__main__":
    main()

# pytool — 부산대병원 재활의학과 연구 자동화 도구

임상연구 문서 생성 및 편집 자동화를 위한 Python 스크립트 모음.

## 도구 목록

| 파일 | 기능 | 의존성 |
|------|------|--------|
| `pnuh_protocol_generator.py` | 프로토콜 제출서 docx 생성 (범용) | python-docx |
| `docx_comment_injector.py` | 기존 docx에 Word 주석(Comment) 삽입 | lxml |
| `clinical_research_scaffold.py` | 임상연구 문서 3종 자동 생성 (예진지+Quick Sheet+CRF) | python-docx, lxml |
| `sap_generator.py` | 통계분석계획서(SAP) docx 자동 생성 — sample size 계산 + power curve + TransCelerate 구조 | python-docx, statsmodels, matplotlib, numpy, scipy |
| `protocol_crf_mapper.py` | Protocol ↔ CRF 일관성 검증 + SPIRIT 2025 체크리스트 (50항목) 자동 검증 | python-docx |

## 사용법

### 1. pnuh_protocol_generator.py

프로토콜 제출서를 config dict 기반으로 자동 생성합니다.

```python
from pnuh_protocol_generator import generate_protocol

config = {
    "title": "프로토콜 부제목",
    "header_title": "헤더에 표시할 프로토콜명",
    "department": "재활의학과",
    "sections": [
        {
            "label": "섹션 라벨",
            "rows": [
                ("소항목명", "BG색상코드", ["내용1", "내용2", ("굵은내용", True)]),
            ],
        },
    ],
    "notes": [("비고 텍스트", False, False)],  # (text, bold, red)
}

generate_protocol(config, "output.docx")
```

기본 색상 코드: `HDR`(진한파랑), `SEC`(연파랑), `SUB`(연녹색), `WARN`(연노랑), `CRIT`(연빨강), `NOTE`(연회색)

### 2. docx_comment_injector.py

기존 Word 파일에 실제 Comment(주석)를 프로그래밍 방식으로 삽입합니다.

**Python에서 사용:**

```python
from docx_comment_injector import inject_comments

comments = [
    {"id": 0, "anchor": "검색할 텍스트", "text": "주석 내용"},
    {"id": 1, "anchor": "다른 텍스트", "text": "주석 내용2"},
]

inject_comments("input.docx", "output.docx", comments,
                author="이재현 (PI)", date="2026-03-27T00:00:00Z")
```

**CLI에서 사용:**

```bash
python3 docx_comment_injector.py input.docx output.docx comments.json
```

comments.json 형식:
```json
{
  "author": "이재현 (PI)",
  "date": "2026-03-27T00:00:00Z",
  "comments": [
    {"id": 0, "anchor": "검색할 텍스트", "text": "주석 내용"},
    {"id": 1, "anchor": "다른 텍스트", "text": "주석 내용2"}
  ]
}
```

### 3. clinical_research_scaffold.py

"Practice-first, Research-seamless" 프레임워크 — config dict 하나로 연구별 맞춤 문서 3종을 생성합니다.

| 문서 | 역할 | 사용 장소 |
|------|------|----------|
| 예진지 | 환자 문진 + 검사자 측정 (baseline) | 외래 예진실 |
| Quick Sheet | 시술/검사실 전용 데이터만 | 시술실/검사실 |
| CRF | 마스터 문서 — 전체 통합 | 사후 정리용 |

**Python에서 사용:**

```python
from clinical_research_scaffold import generate_research_docs, EXAMPLE_CONFIG

# FSHD 예시 config로 3종 생성
generate_research_docs(EXAMPLE_CONFIG, output_dir="./output")

# 새 연구: config만 바꾸면 됨
my_config = {
    "study": {
        "short_name": "CRPS",
        "title_ko": "CRPS 환자의 경피 신경자극 효과",
        "department": "부산대학교병원 재활의학과",
        "pi": "이재현",
        "preexam_title": "CRPS 예진 문진표",
    },
    "baseline": {
        "sections": [
            {"name": "통증 평가", "items": [
                {"label": "VAS", "hint": "___ /100", "source": "patient"},
                {"label": "DN4", "hint": "___ /10", "source": "patient"},
            ]},
        ],
    },
    "procedure": {
        "sections": [
            {"name": "시술 기록", "items": [
                {"label": "자극 부위", "hint": "________", "source": "procedure"},
                {"label": "자극 강도", "hint": "___ mA", "source": "procedure"},
            ]},
        ],
    },
    # ... inclusion, exclusion, followup 등
}
generate_research_docs(my_config, output_dir="./CRPS_docs")
```

**CLI에서 사용:**

```bash
# 예시 config 출력
python3 clinical_research_scaffold.py --example > my_study.json

# config 수정 후 문서 생성
python3 clinical_research_scaffold.py my_study.json -o ./output

# YAML도 지원 (pip install pyyaml)
python3 clinical_research_scaffold.py my_study.yaml -o ./output
```

**워크플로우:**
```
1. --example로 config 템플릿 추출
2. 연구에 맞게 config 수정 (항목, 기준, 시술 프로토콜 등)
3. generate_research_docs() 실행 → 예진지 + Quick Sheet + CRF 생성
4. 생성된 docx를 실제 양식에 맞게 최종 수정
5. 연구 진행
```

**데이터 소스 색상:**
- 🟡 `patient` — 환자 직접 작성
- 🔵 `examiner`/`emr` — 검사자 측정 / EMR 후입력
- 🟢 `calc` — 자동 계산
- 🔴 `procedure` — 시술/검사 중 기록

### 4. sap_generator.py

연구 config dict → TransCelerate Common SAP Template 구조의 통계분석계획서(Statistical Analysis Plan)를 자동 생성합니다. Sample size 계산, sensitivity table, power curve가 내장되어 있습니다.

**Python에서 사용:**

```python
from sap_generator import generate_sap, calculate_sample_size, EXAMPLE_CONFIG

# Sample size만 계산
result = calculate_sample_size(EXAMPLE_CONFIG["primary_outcome"])
# → {'n_per_group': 39, 'n_per_group_adjusted': 49, 'total_n': 98, ...}

# SAP 문서 전체 생성
generate_sap(EXAMPLE_CONFIG, "My_Study_SAP.docx")
```

**CLI에서 사용:**

```bash
# FSHD 예시 config 출력
python3 sap_generator.py --example

# 예시 config로 SAP 즉시 생성
python3 sap_generator.py --example --run

# JSON config 파일로 생성
python3 sap_generator.py my_study_config.json --output My_SAP.docx
```

**Config 핵심 필드:**
- `study_title`, `design` (randomized_controlled_trial, cohort, etc.)
- `arms` (군별 이름 + 목표 n)
- `primary_outcome` (name, type, measure, mcid, sd, alpha, power, dropout_rate)
- `analysis` (primary_method, missing_data, multiplicity, subgroup, sensitivity)
- `randomization` (method, block_size, stratification)

**지원 검정 유형:** independent_t, paired_t, anova, chi_square

**SAP 출력 구조 (TransCelerate 9섹션):**
1. Administrative Information
2. Study Objectives & Hypotheses
3. Study Design Summary
4. Sample Size Justification (자동 계산 + sensitivity table + power curve)
5. Randomization & Blinding
6. Analysis Populations
7. Statistical Methods (primary/secondary/subgroup/sensitivity/missing data)
8. Tables & Figures Shells
9. SAP Amendment Log

### 5. protocol_crf_mapper.py

Protocol ↔ CRF 일관성 검증 + SPIRIT 2025 체크리스트 자동 검증 도구. scaffold config (CRF/예진지/Quick Sheet)와 SAP config를 cross-reference하여 변수 매핑 행렬과 gap report를 생성합니다.

**기능 2가지:**
1. **Mapping Matrix**: SAP에 정의된 모든 outcome 변수가 CRF에 수집되는지 확인
2. **SPIRIT 2025 Validator**: 프로토콜이 SPIRIT 2025 50개 항목을 충족하는지 자동 검증

**Python에서 사용:**

```python
from protocol_crf_mapper import (
    generate_mapping_matrix, validate_spirit_2025,
    generate_full_report, EXAMPLE_SCAFFOLD_CONFIG, EXAMPLE_SAP_CONFIG
)

# 1. Mapping matrix만 확인
matrix = generate_mapping_matrix(EXAMPLE_SCAFFOLD_CONFIG, EXAMPLE_SAP_CONFIG)
print(f"Coverage: {matrix['summary']['coverage_pct']}%")
print(f"Gaps: {len(matrix['gaps'])}")

# 2. SPIRIT 2025 validation
result = validate_spirit_2025(EXAMPLE_SCAFFOLD_CONFIG, EXAMPLE_SAP_CONFIG)
print(f"Score: {result['score']['addressed']}/{result['score']['auto_checkable']}")

# 3. Full docx report (mapping + SPIRIT + recommendations)
generate_full_report(EXAMPLE_SCAFFOLD_CONFIG, EXAMPLE_SAP_CONFIG, "report.docx")
```

**CLI에서 사용:**

```bash
# 예시 config로 즉시 실행
python3 protocol_crf_mapper.py --example --run

# JSON config 파일로 실행
python3 protocol_crf_mapper.py scaffold.json sap.json --output report.docx

# 텍스트 요약만 (docx 미생성)
python3 protocol_crf_mapper.py --example --run --summary
```

**Docx 보고서 구조:**
1. Protocol ↔ CRF Mapping Matrix (SAP 변수 → CRF 매칭 테이블)
2. Gap Analysis (CRF 누락 변수 목록)
3. SPIRIT 2025 Checklist (50항목 충족 현황)
4. Recommendations (자동 생성 개선 권고)

**변수 매칭 엔진:** 의학 용어 synonym dictionary + token overlap + substring matching

## 의존성 설치

```bash
pip install python-docx lxml pyyaml statsmodels matplotlib numpy scipy
# pyyaml: YAML config 사용 시만 필요
# statsmodels, matplotlib, numpy, scipy: sap_generator.py용
```

## 원본 스크립트

이 도구들은 다음 프로젝트별 스크립트에서 일반화되었습니다:

- `temporaryfiles/create_cpet_protocol.py` → `pnuh_protocol_generator.py`
- `temporaryfiles/create_icg_lymphography_protocol.py` → `pnuh_protocol_generator.py`
- `temporaryfiles/create_icg_protocol.py` → `pnuh_protocol_generator.py`
- `FSHD/build_commented_docx.py` → `docx_comment_injector.py`
- `FSHD/{FSHD_CRF, 통증예진_어깨, Ultrasound Quick Sheet}.docx` → `clinical_research_scaffold.py`
- TransCelerate Common SAP Template + statsmodels power analysis → `sap_generator.py`
- SPIRIT 2025 checklist (BMJ/JAMA/Lancet) + scaffold/SAP config 구조 → `protocol_crf_mapper.py`

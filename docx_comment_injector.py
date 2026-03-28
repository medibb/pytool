#!/usr/bin/env python3
"""
부산대병원 재활의학과 — Word(docx) 주석(Comment) 삽입 도구
기존 docx 파일에 실제 Word Comment를 프로그래밍 방식으로 주입.

사용법:
    1. 직접 실행 (CLI):
        python3 docx_comment_injector.py input.docx output.docx comments.json

    2. 스크립트에서 import:
        from docx_comment_injector import inject_comments

        comments = [
            {"id": 0, "anchor": "검색할 텍스트", "text": "주석 내용"},
            {"id": 1, "anchor": "다른 텍스트", "text": "주석 내용2"},
        ]
        inject_comments("input.docx", "output.docx", comments,
                        author="이재현 (PI)", date="2026-03-27T00:00:00Z")

동작 원리:
    python-docx는 Comment를 지원하지 않으므로, docx(ZIP) 내부의 XML을
    직접 조작하여 주석을 삽입합니다:
    - word/comments.xml: 주석 본문 생성
    - word/document.xml: commentRangeStart/End + commentReference 마커 삽입
    - word/_rels/document.xml.rels: comments 관계 추가
    - [Content_Types].xml: comments 콘텐츠 타입 등록
"""

import zipfile
import os
import json
import sys
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/comments"
)
COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.comments+xml"
)


def _escape_xml(text):
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


def _build_comments_xml(comments, author="Author", date="2026-01-01T00:00:00Z"):
    """Build word/comments.xml content from comment list."""
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]
    initials = "".join(w[0] for w in author.split() if w[0].isalpha())[:3] or "A"
    for c in comments:
        cid = c["id"]
        text = _escape_xml(c["text"])
        lines.append(
            f'<w:comment w:id="{cid}" w:author="{_escape_xml(author)}" '
            f'w:date="{date}" w:initials="{initials}">'
            f'<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
            f'<w:annotationRef/></w:r>'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
            f'</w:p></w:comment>'
        )
    lines.append('</w:comments>')
    return '\n'.join(lines)


def _strip_existing_comment_markers(doc_xml_bytes):
    """Remove any pre-existing commentRangeStart/End/Reference from document.xml."""
    tree = etree.fromstring(doc_xml_bytes)

    for tag in ['commentRangeStart', 'commentRangeEnd']:
        for el in tree.findall(f'.//{{{W_NS}}}{tag}'):
            el.getparent().remove(el)

    for ref in tree.findall(f'.//{{{W_NS}}}commentReference'):
        run = ref.getparent()
        if run is not None and run.tag == f'{{{W_NS}}}r':
            run.getparent().remove(run)
        else:
            ref.getparent().remove(ref)

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def _inject_comment_markers(doc_xml_bytes, comments):
    """Insert commentRangeStart/End + commentReference for each comment into document.xml."""
    tree = etree.fromstring(doc_xml_bytes)
    nsmap = {'w': W_NS}
    all_paras = tree.findall('.//w:p', nsmap)

    anchored = 0
    skipped = []

    for c in comments:
        cid = c["id"]
        anchor_lower = c["anchor"].lower()
        found = False

        for p in all_paras:
            texts = p.findall('.//w:t', nsmap)
            para_text = ''.join(t.text or '' for t in texts).lower()
            if anchor_lower in para_text:
                range_start = etree.SubElement(p, f'{{{W_NS}}}commentRangeStart')
                range_start.set(f'{{{W_NS}}}id', str(cid))
                p.insert(0, range_start)

                range_end = etree.SubElement(p, f'{{{W_NS}}}commentRangeEnd')
                range_end.set(f'{{{W_NS}}}id', str(cid))

                ref_run = etree.SubElement(p, f'{{{W_NS}}}r')
                ref_rpr = etree.SubElement(ref_run, f'{{{W_NS}}}rPr')
                ref_style = etree.SubElement(ref_rpr, f'{{{W_NS}}}rStyle')
                ref_style.set(f'{{{W_NS}}}val', 'CommentReference')
                ref_elem = etree.SubElement(ref_run, f'{{{W_NS}}}commentReference')
                ref_elem.set(f'{{{W_NS}}}id', str(cid))

                anchored += 1
                print(f"  ✓ Comment {cid}: '{c['anchor']}' found")
                found = True
                break

        if not found:
            skipped.append(c)
            print(f"  ✗ Comment {cid}: '{c['anchor']}' not found")

    return (
        etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True),
        anchored,
        skipped,
    )


def _clean_and_add_comments_rel(rels_xml_bytes):
    """Ensure exactly one comments relationship in document.xml.rels."""
    tree = etree.fromstring(rels_xml_bytes)
    ns = PKG_NS

    to_remove = [rel for rel in tree if 'comments' in (rel.get('Type') or '')]
    for el in to_remove:
        tree.remove(el)

    max_id = 0
    for rel in tree:
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                max_id = max(max_id, int(rid[3:]))
            except ValueError:
                pass

    new_rel = etree.SubElement(tree, f'{{{ns}}}Relationship')
    new_rel.set('Id', f'rId{max_id + 1}')
    new_rel.set('Type', COMMENTS_REL_TYPE)
    new_rel.set('Target', 'comments.xml')

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def _add_content_type(ct_xml_bytes):
    """Add comments content type override if missing."""
    tree = etree.fromstring(ct_xml_bytes)
    ns = CT_NS

    for child in tree:
        if 'comments.xml' in (child.get('PartName') or ''):
            return ct_xml_bytes

    override = etree.SubElement(tree, f'{{{ns}}}Override')
    override.set('PartName', '/word/comments.xml')
    override.set('ContentType', COMMENTS_CT)

    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def inject_comments(input_path, output_path, comments,
                    author="Author", date="2026-01-01T00:00:00Z"):
    """
    기존 docx 파일에 Word Comment를 삽입하여 새 파일로 저장.

    Args:
        input_path: 원본 docx 파일 경로
        output_path: 출력 docx 파일 경로
        comments: list of dict, 각 항목은 {"id": int, "anchor": str, "text": str}
            - id: 고유 코멘트 ID (0부터)
            - anchor: document.xml에서 검색할 텍스트 (대소문자 무시)
            - text: 코멘트 본문
        author: 코멘트 작성자 이름
        date: 코멘트 날짜 (ISO 8601)

    Returns:
        dict with keys: total, anchored, skipped (list of unmatched comments)
    """
    print(f"Injecting {len(comments)} comments into {input_path}...")

    comments_xml = _build_comments_xml(comments, author=author, date=date)

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

    anchored = 0
    skipped = []

    with zipfile.ZipFile(input_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/comments.xml':
                    continue  # skip existing — we write our own

                data = zin.read(item.filename)

                if item.filename == 'word/document.xml':
                    data = _strip_existing_comment_markers(data)
                    data, anchored, skipped = _inject_comment_markers(data, comments)

                elif item.filename == 'word/_rels/document.xml.rels':
                    data = _clean_and_add_comments_rel(data)

                elif item.filename == '[Content_Types].xml':
                    data = _add_content_type(data)

                zout.writestr(item, data)

            zout.writestr('word/comments.xml', comments_xml.encode('utf-8'))

    size = os.path.getsize(output_path)
    print(f"\n✓ Done: {output_path} ({size/1024:.1f} KB)")
    print(f"✓ {anchored}/{len(comments)} comments anchored, {len(skipped)} skipped")

    # Verify
    with zipfile.ZipFile(output_path, 'r') as z:
        names = z.namelist()
        assert 'word/comments.xml' in names, "comments.xml missing!"
        from collections import Counter
        dupes = {k: v for k, v in Counter(names).items() if v > 1}
        if dupes:
            print(f"  ⚠ Duplicate entries: {dupes}")
        else:
            print("  ✓ No duplicate entries")

    return {"total": len(comments), "anchored": anchored, "skipped": skipped}


# ══════════════════════════════════════════════════════════════
# CLI usage
# ══════════════════════════════════════════════════════════════

def _print_usage():
    print("Usage: python3 docx_comment_injector.py <input.docx> <output.docx> <comments.json>")
    print()
    print("comments.json format:")
    print('  {')
    print('    "author": "이재현 (PI)",')
    print('    "date": "2026-03-27T00:00:00Z",')
    print('    "comments": [')
    print('      {"id": 0, "anchor": "검색할 텍스트", "text": "주석 내용"},')
    print('      {"id": 1, "anchor": "다른 텍스트", "text": "주석 내용2"}')
    print('    ]')
    print('  }')


if __name__ == "__main__":
    if len(sys.argv) < 4:
        _print_usage()
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    json_path = sys.argv[3]

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    result = inject_comments(
        input_path,
        output_path,
        data["comments"],
        author=data.get("author", "Author"),
        date=data.get("date", "2026-01-01T00:00:00Z"),
    )

    if result["skipped"]:
        print(f"\n⚠ Skipped comments (anchor not found):")
        for c in result["skipped"]:
            print(f"  - [{c['id']}] '{c['anchor']}'")

#!/usr/bin/env python3
"""Create a tighter single-page variant of Nick_Myers_Resume_Base.docx.

Run from repo root:
  python resumes/build_resume_base_v2.py
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

RESUMES_DIR = Path(__file__).resolve().parent
SRC = RESUMES_DIR / "Nick_Myers_Resume_Base.docx"
OUT = RESUMES_DIR / "Nick_Myers_Resume_Base_V2.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Page margins (twips): 0.5" sides/top/bottom vs 1" in the base file
MARGIN_TWIPS = "720"

# Paragraph spacing (twips): ~single line, minimal gap between paragraphs
TIGHT_SPACING = '<w:spacing w:after="20" w:before="0" w:line="220" w:lineRule="auto"/>'
HEADER_SPACING = '<w:spacing w:after="0" w:before="0" w:line="200" w:lineRule="auto"/>'
SECTION_HEADING_SPACING = '<w:spacing w:after="40" w:before="120" w:line="220" w:lineRule="auto"/>'

SPACING_RE = re.compile(r"<w:spacing[^/]*/>")
PGMAR_RE = re.compile(r"<w:pgMar[^/]*/>")


def paragraph_text(p_xml: str) -> str:
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p_xml))


def is_empty_paragraph(p_xml: str) -> bool:
    if paragraph_text(p_xml).strip():
        return False
    # Keep paragraphs that only carry section properties
    if "<w:sectPr" in p_xml:
        return False
    return True


def tighten_pg_margins(pgmar_xml: str) -> str:
    for attr in ("top", "right", "bottom", "left"):
        pgmar_xml = re.sub(
            rf'(\bw:{attr}=")\d+(")', rf"\g<1>{MARGIN_TWIPS}\2", pgmar_xml
        )
    return pgmar_xml


def tighten_document_xml(xml: str) -> str:
    xml = PGMAR_RE.sub(lambda m: tighten_pg_margins(m.group(0)), xml)

    # Slightly tighter bullet hanging indent
    xml = xml.replace('w:left="720" w:hanging="360"', 'w:left="540" w:hanging="270"')

    # Body font 10pt -> 9pt; name 16pt -> 14pt
    xml = re.sub(r'(<w:sz w:val=")20(")', r"\g<1>18\2", xml)
    xml = re.sub(r'(<w:szCs w:val=")20(")', r"\g<1>18\2", xml)
    xml = re.sub(r'(<w:sz w:val=")32(")', r"\g<1>28\2", xml)
    xml = re.sub(r'(<w:szCs w:val=")32(")', r"\g<1>28\2", xml)

    parts = re.split(r"(<w:p\b[^>]*>)", xml)
    if len(parts) < 2:
        return apply_global_spacing(xml)

    out: list[str] = [parts[0]]
    i = 1
    while i < len(parts):
        p_open = parts[i]
        p_body = parts[i + 1] if i + 1 < len(parts) else ""
        i += 2

        if is_empty_paragraph(p_open + p_body):
            continue

        text = paragraph_text(p_open + p_body).strip()
        if text == "Nick Myers":
            p_body = apply_paragraph_spacing(p_body, HEADER_SPACING)
        elif text.isupper() and len(text) < 50:
            # SUMMARY, EXPERIENCE, TECHNICAL SKILLS, etc.
            p_body = apply_paragraph_spacing(p_body, SECTION_HEADING_SPACING)
        else:
            p_body = apply_paragraph_spacing(p_body, TIGHT_SPACING)

        out.append(p_open)
        out.append(p_body)

    return "".join(out)


def apply_paragraph_spacing(p_body: str, spacing_xml: str) -> str:
    if SPACING_RE.search(p_body):
        p_body = SPACING_RE.sub(spacing_xml, p_body, count=1)
    elif "<w:pPr>" in p_body:
        p_body = p_body.replace("<w:pPr>", f"<w:pPr>{spacing_xml}", 1)
    else:
        p_body = f"<w:pPr>{spacing_xml}</w:pPr>" + p_body
    return p_body


def apply_global_spacing(xml: str) -> str:
    xml = SPACING_RE.sub(TIGHT_SPACING, xml)
    return xml


def build_v2() -> None:
    if not SRC.exists():
        raise SystemExit(f"Missing source resume: {SRC}")

    shutil.copy2(SRC, OUT)

    with zipfile.ZipFile(OUT, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        numbering_xml = zin.read("word/numbering.xml").decode("utf-8")
        other = {
            name: zin.read(name)
            for name in zin.namelist()
            if name not in ("word/document.xml", "word/numbering.xml")
        }

    doc_xml = tighten_document_xml(doc_xml)
    numbering_xml = numbering_xml.replace(
        'w:left="720" w:hanging="360"', 'w:left="540" w:hanging="270"'
    )

    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in other.items():
            zout.writestr(name, data)
        zout.writestr("word/document.xml", doc_xml.encode("utf-8"))
        zout.writestr("word/numbering.xml", numbering_xml.encode("utf-8"))

    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build_v2()

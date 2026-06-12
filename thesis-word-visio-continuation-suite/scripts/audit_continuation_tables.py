# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path

import fitz
from lxml import etree


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
TABLE_RE = re.compile(r"表\s*(\d+(?:[-－.]\d+)?)\s*([^\n\r]*)")
CONT_TABLE_RE = re.compile(r"续表\s*(\d+(?:[-－.]\d+)?)\s*([^\n\r]*)")
CONT_RE = re.compile(r"(续表|（续）|\(续\))")


def extract_docx_table_captions(docx: Path) -> list[dict]:
    captions: list[dict] = []
    with zipfile.ZipFile(docx) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    for idx, p in enumerate(root.findall(".//w:p", NS)):
        text = "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()
        match = CONT_TABLE_RE.match(text) or TABLE_RE.match(text)
        if match:
            captions.append(
                {
                    "paragraph_index": idx,
                    "number": match.group(1),
                    "text": text,
                    "is_continuation": bool(CONT_RE.search(text)),
                }
            )
    return captions


def extract_pdf_table_mentions(pdf: Path) -> list[dict]:
    doc = fitz.open(pdf)
    mentions: list[dict] = []
    for page_index, page in enumerate(doc, start=1):
        text = page.get_text()
        for match in list(TABLE_RE.finditer(text)) + list(CONT_TABLE_RE.finditer(text)):
            line = text[match.start() : text.find("\n", match.start())]
            if not line:
                line = match.group(0)
            mentions.append(
                {
                    "page": page_index,
                    "number": match.group(1),
                    "text": line.strip(),
                    "is_continuation": bool(CONT_RE.search(line)),
                }
            )
    return mentions


def build_findings(docx_captions: list[dict], pdf_mentions: list[dict]) -> list[dict]:
    findings: list[dict] = []
    by_number: dict[str, list[dict]] = {}
    for item in pdf_mentions:
        by_number.setdefault(item["number"], []).append(item)

    for number, mentions in sorted(by_number.items()):
        pages = sorted({m["page"] for m in mentions})
        continuation_mentions = [m for m in mentions if m["is_continuation"]]
        if len(mentions) > 1:
            findings.append(
                {
                    "type": "repeated-table-number",
                    "number": number,
                    "pages": pages,
                    "has_continuation_title": bool(continuation_mentions),
                    "mentions": mentions,
                    "note": "Repeated table number may be a valid continuation or a duplicated caption; visually inspect pages.",
                }
            )

    continuation_docx = [c for c in docx_captions if c["is_continuation"]]
    for item in continuation_docx:
        matching_pdf = [m for m in pdf_mentions if m["number"] == item["number"] and m["is_continuation"]]
        findings.append(
            {
                "type": "docx-continuation-caption",
                "number": item["number"],
                "paragraph_index": item["paragraph_index"],
                "text": item["text"],
                "pdf_pages": sorted({m["page"] for m in matching_pdf}),
                "note": "Confirm this continuation caption is only used where the logical table crosses a page.",
            }
        )

    return findings


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()

    docx_captions = extract_docx_table_captions(args.docx)
    pdf_mentions = extract_pdf_table_mentions(args.pdf)
    result = {
        "docx": str(args.docx),
        "pdf": str(args.pdf),
        "docx_table_captions": docx_captions,
        "pdf_table_mentions": pdf_mentions,
        "findings": build_findings(docx_captions, pdf_mentions),
    }
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()

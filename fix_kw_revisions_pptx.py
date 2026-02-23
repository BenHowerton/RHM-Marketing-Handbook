#!/usr/bin/env python3
"""Revise a PPTX deck for KW formatting/content requirements.

This tool updates slide XML in a .pptx package to:
- remove em dashes,
- remove AI-type text tokens,
- enforce text autofit/wrap in text boxes,
- keep frame transforms within slide bounds.
"""

from __future__ import annotations

import argparse
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


@dataclass
class Stats:
    slides: int = 0
    text_replacements: int = 0
    autofit_updates: int = 0
    wrap_updates: int = 0
    transform_clamps: int = 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Revise KW PPTX formatting/text.")
    parser.add_argument("--input", required=True, help="Source PPTX path")
    parser.add_argument("--output", required=True, help="Output PPTX path")
    parser.add_argument("--report", default="pptx_revision_report.md", help="Markdown report path")
    return parser.parse_args()


def parse_slide_size(presentation_xml: bytes) -> tuple[int, int]:
    root = ET.fromstring(presentation_xml)
    sld_sz = root.find("p:sldSz", NS)
    if sld_sz is None:
        return 9144000, 5143500
    return int(sld_sz.get("cx", "9144000")), int(sld_sz.get("cy", "5143500"))


def clean_text(text: str) -> str:
    text = text.replace("—", " ")
    text = re.sub(r"\bA\.I\.\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bAI\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bartificial\s+intelligence\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s{2,}", " ", text)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    return text


def ensure_autofit_and_wrap(sp: ET.Element) -> tuple[bool, bool]:
    tx_body = sp.find("p:txBody", NS)
    if tx_body is None:
        return False, False

    body_pr = tx_body.find("a:bodyPr", NS)
    if body_pr is None:
        return False, False

    autofit_changed = False
    wrap_changed = False

    for tag in ("a:noAutofit", "a:normAutofit"):
        el = body_pr.find(tag, NS)
        if el is not None:
            body_pr.remove(el)
            autofit_changed = True

    if body_pr.find("a:spAutoFit", NS) is None:
        body_pr.append(ET.Element(f"{{{NS['a']}}}spAutoFit"))
        autofit_changed = True

    if body_pr.get("wrap") != "square":
        body_pr.set("wrap", "square")
        wrap_changed = True

    return autofit_changed, wrap_changed


def clamp_transform(xfrm: ET.Element, max_w: int, max_h: int) -> bool:
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return False

    x = int(off.get("x", "0"))
    y = int(off.get("y", "0"))
    w = int(ext.get("cx", "0"))
    h = int(ext.get("cy", "0"))
    changed = False

    if x < 0:
        x = 0
        changed = True
    if y < 0:
        y = 0
        changed = True
    if x + w > max_w:
        w = max(0, max_w - x)
        changed = True
    if y + h > max_h:
        h = max(0, max_h - y)
        changed = True

    if changed:
        off.set("x", str(x))
        off.set("y", str(y))
        ext.set("cx", str(w))
        ext.set("cy", str(h))

    return changed


def process_slide(xml_bytes: bytes, slide_w: int, slide_h: int, totals: Stats) -> bytes:
    root = ET.fromstring(xml_bytes)

    for t in root.findall(".//a:t", NS):
        original = t.text or ""
        cleaned = clean_text(original)
        if cleaned != original:
            totals.text_replacements += 1
            t.text = cleaned

    for sp in root.findall(".//p:sp", NS):
        autofit_changed, wrap_changed = ensure_autofit_and_wrap(sp)
        if autofit_changed:
            totals.autofit_updates += 1
        if wrap_changed:
            totals.wrap_updates += 1

    for path in (
        ".//p:sp/p:spPr/a:xfrm",
        ".//p:pic/p:spPr/a:xfrm",
        ".//p:graphicFrame/p:xfrm",
        ".//p:cxnSp/p:spPr/a:xfrm",
        ".//p:grpSp/p:grpSpPr/a:xfrm",
    ):
        for xfrm in root.findall(path, NS):
            if clamp_transform(xfrm, slide_w, slide_h):
                totals.transform_clamps += 1

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def run(input_path: str, output_path: str, report_path: str) -> Stats:
    totals = Stats()

    with zipfile.ZipFile(input_path, "r") as zin:
        slide_w, slide_h = parse_slide_size(zin.read("ppt/presentation.xml"))

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
                    totals.slides += 1
                    data = process_slide(data, slide_w, slide_h, totals)
                zout.writestr(item, data)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# PPTX Revision Report\n\n")
        f.write(f"- Source file: `{input_path}`\n")
        f.write(f"- Output file: `{output_path}`\n")
        f.write(f"- Slides processed: {totals.slides}\n")
        f.write(f"- Text nodes cleaned (Em dashes / AI text): {totals.text_replacements}\n")
        f.write(f"- Text boxes set to shape autofit: {totals.autofit_updates}\n")
        f.write(f"- Text boxes set to wrapped text: {totals.wrap_updates}\n")
        f.write(f"- Shape/chart/image frames clamped to slide bounds: {totals.transform_clamps}\n")

    return totals


def main() -> None:
    args = parse_args()
    print(
        run(
            input_path=args.input,
            output_path=args.output,
            report_path=args.report,
        )
    )


if __name__ == "__main__":
    main()

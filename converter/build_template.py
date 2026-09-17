"""Dev fixture generator for templates/everglades.pptx.

This is NOT Communications’ official brand template. It produces a stand-in
16:9 deck (python-pptx default layouts + Everglades theme colors) so the
rebrand engine can run locally. Chrome is drawn on the slide master — never
on individual slides during conversion. Replace templates/everglades.pptx
with the official TEMPLATE.pptx / .potx when Communications provides it.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

from brand import LAGOON, MARSH, WHITE, add_brand_chrome
from layouts import PPTX_LAYOUT_NAME, catalog_payload

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = ROOT / "templates" / "everglades.pptx"
CATALOG_PATH = Path(__file__).resolve().parent / "layout_catalog.json"

THEME_COLORS = {
    "dk1": "0B433F",
    "lt1": "FFFFFF",
    "dk2": "116680",
    "lt2": "E7F3C1",
    "accent1": "0893A6",
    "accent2": "C1EA40",
    "accent3": "0B433F",
    "accent4": "116680",
    "accent5": "0893A6",
    "accent6": "C1EA40",
    "hlink": "0893A6",
    "folHlink": "0B433F",
}


def _theme_part(prs: Presentation):
    master = prs.slide_masters[0]
    for rel in master.part.rels.values():
        if rel.reltype == RT.THEME:
            return rel.target_part
    raise RuntimeError("Theme part not found on slide master.")


def _set_srgb(parent, hex_color: str) -> None:
    for child in list(parent):
        parent.remove(child)
    srgb = etree.SubElement(parent, f"{{{A_NS}}}srgbClr")
    srgb.set("val", hex_color.upper())


def patch_theme(prs: Presentation) -> None:
    part = _theme_part(prs)
    root = etree.fromstring(part.blob)
    scheme = root.find(f".//{{{A_NS}}}clrScheme")
    if scheme is None:
        raise RuntimeError("Color scheme missing from theme.")
    for name, hex_color in THEME_COLORS.items():
        node = scheme.find(f"{{{A_NS}}}{name}")
        if node is not None:
            _set_srgb(node, hex_color)

    for latin in root.findall(f".//{{{A_NS}}}latin"):
        latin.set("typeface", "Poppins")
    for ea in root.findall(f".//{{{A_NS}}}ea"):
        ea.set("typeface", "Poppins")
    part.blob = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _fill_run(shape, text: str, *, size: int, color, bold: bool = True) -> None:
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Poppins"


def build() -> Path:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    patch_theme(prs)

    by_name = {layout.name: layout for layout in prs.slide_layouts}
    missing = [name for name in PPTX_LAYOUT_NAME.values() if name not in by_name]
    if missing:
        available = ", ".join(by_name)
        raise RuntimeError(f"Missing layouts {missing}. Available: {available}")

    # Sample slides are deleted before conversion. python-pptx masters cannot
    # add_shape, so stand-in chrome stays on these samples only — not on output.
    samples = [
        ("title", "America's Everglades", "Restoration is worth it."),
        ("section", "Science", ""),
        ("content", "What we protect", "Water\nWildlife\nCommunities"),
        ("two_column", "Two stories", "Left column"),
        ("image", "Field work", "Caption"),
        ("quote", "A thriving Everglades means a thriving economy.", ""),
        ("chart", "Funding over time", ""),
    ]
    for layout_id, title, subtitle in samples:
        layout = by_name[PPTX_LAYOUT_NAME[layout_id]]
        slide = prs.slides.add_slide(layout)
        add_brand_chrome(slide, prs, dark=layout_id == "title")
        title_set = False
        for shape in slide.placeholders:
            name = (shape.name or "").lower()
            ph_type = str(shape.placeholder_format.type).lower()
            if not title_set and ("title" in name or "title" in ph_type):
                _fill_run(
                    shape,
                    title,
                    size=32 if layout_id != "title" else 40,
                    color=WHITE if layout_id == "title" else MARSH,
                )
                title_set = True
            elif subtitle and (
                "subtitle" in name
                or "subtitle" in ph_type
                or "content" in name
                or "text" in name
                or "body" in ph_type
            ):
                _fill_run(
                    shape,
                    subtitle.replace("\n", "\n"),
                    size=18,
                    color=WHITE if layout_id == "title" else LAGOON,
                    bold=False,
                )

    TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    prs.save(TEMPLATE_PATH)
    CATALOG_PATH.write_text(
        json.dumps(catalog_payload(), indent=2) + "\n",
        encoding="utf-8",
    )
    return TEMPLATE_PATH


if __name__ == "__main__":
    path = build()
    prs = Presentation(path)
    print(f"Wrote {path}")
    print("Layouts:")
    for i, layout in enumerate(prs.slide_layouts):
        ph = [
            f"{shape.name}:{shape.placeholder_format.type}"
            for shape in layout.placeholders
        ]
        print(f"  [{i}] {layout.name} — {', '.join(ph)}")
    sys.exit(0)

"""Dev-only layout names for converter/build_template.py (stand-in generator).

Conversion uses rebrand/inspect_template.py → brand_tokens.json["layout_map"].
Staff-facing quote/chart ids are retired.
"""

from __future__ import annotations

# Staff-facing ids used by the preview UI and /api/convert.
LAYOUT_IDS = (
    "title",
    "section",
    "content",
    "two_column",
    "image",
    "quote",
    "chart",
)

# python-pptx default template layout names (widescreen 16:9 branded copy).
PPTX_LAYOUT_NAME: dict[str, str] = {
    "title": "Title Slide",
    "section": "Section Header",
    "content": "Title and Content",
    "two_column": "Two Content",
    "image": "Picture with Caption",
    "quote": "Title Only",
    "chart": "Title Only",
}

PLACEHOLDERS: dict[str, tuple[str, ...]] = {
    "title": ("center title", "subtitle"),
    "section": ("title", "text"),
    "content": ("title", "body"),
    "two_column": ("title", "body", "body"),
    "image": ("title", "picture", "caption"),
    "quote": ("title",),
    "chart": ("title", "embedded chart + Excel workbook"),
}

LABELS: dict[str, str] = {
    "title": "Title",
    "section": "Section",
    "content": "Content",
    "two_column": "Two column",
    "image": "Image",
    "quote": "Quote",
    "chart": "Chart",
}

BRAND = {
    "marsh": "0B433F",
    "lagoon": "0893A6",
    "sawgrass": "C1EA40",
    "ink": "1C1C1C",
    "white": "FFFFFF",
    "accent_dark": "116680",
}


def catalog_payload() -> dict[str, object]:
    return {
        "ids": list(LAYOUT_IDS),
        "pptxLayoutName": PPTX_LAYOUT_NAME,
        "placeholders": {key: list(value) for key, value in PLACEHOLDERS.items()},
        "labels": LABELS,
        "brand": BRAND,
        "notes": (
            "Official Communications template should keep these layout ids. "
            "Chart layout must have a title placeholder and room for a live Office chart."
        ),
    }

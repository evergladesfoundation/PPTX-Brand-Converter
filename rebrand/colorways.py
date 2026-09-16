"""Green / blue colorways for the 2023 Everglades Foundation template.

The official deck ships as sample slides (one dummy DEFAULT layout) with a
documented lime + teal palette. Colorways are theme-token sets plus an sRGB
swap so inspect/plan/build/QA all share one selected map.
"""

from __future__ import annotations

import io
from copy import deepcopy
from pathlib import Path
from typing import Any

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from helpers import (
    A_NS,
    P_NS,
    R_NS,
    apply_picture_crop,
    classify_layout_role,
    contain_fit,
    content_safe_box,
    delete_shape,
    enum_name,
    hex_color,
    is_edge_logo,
    iter_shapes_abs,
    maybe_recompress,
    placeholder_type,
    resolve_layout_map,
    theme_from_part_xml,
)

LIME = "C1D451"
TEAL = "00ACBF"
MANGROVE = "82C769"
SAND = "F1F3E4"
NAVY = "0E0340"
MUTED = "5C5675"
DEEP_TEAL = "00808F"
CREAM = "FDFDFD"
SHELL_LINE = "DFE2CE"

EF_WHITELIST = [
    LIME, TEAL, MANGROVE, SAND, NAVY, MUTED, DEEP_TEAL, CREAM, SHELL_LINE,
    "FFFFFF", "000000", "3B3159", "D6F2F5",
]

SKIP_PROTOTYPE_HINTS = ("notes for staff", "how to use this template", "internal reference")


def _base_theme(major: str, minor: str, accent1: str, accent2: str) -> dict[str, Any]:
    return {
        "colors": {
            "dk1": NAVY,
            "lt1": CREAM,
            "dk2": MUTED,
            "lt2": SAND,
            "accent1": accent1,
            "accent2": accent2,
            "accent3": MANGROVE,
            "accent4": DEEP_TEAL,
            "accent5": TEAL,
            "accent6": LIME,
            "hlink": TEAL,
            "folHlink": NAVY,
        },
        "fonts": {
            "major": major or "Questrial",
            "minor": minor or "Arial",
            "major_ea": "",
            "minor_ea": "",
            "major_cs": "",
            "minor_cs": "",
        },
    }


def _sample_fonts(prs) -> tuple[str, str]:
    major = minor = ""
    for slide in list(prs.slides)[:4]:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    name = run.font.name or ""
                    if not name:
                        continue
                    if name.lower() in {"questrial", "poppins"} and not major:
                        major = name
                    if name.lower() in {"arial", "arial nova"} and not minor:
                        minor = name
    return major or "Questrial", minor or "Arial"


def _prototype_role(index: int, title: str, notes: str, total: int) -> str | None:
    blob = f"{title} {notes}".lower()
    if any(hint in blob for hint in SKIP_PROTOTYPE_HINTS):
        return None
    if "cover" in notes.lower() or index == 0:
        return "TITLE"
    if "section divider" in notes.lower() or (title.strip().isdigit() and "section" in blob):
        return "SECTION"
    if "closing" in notes.lower() or "thank you" in title.lower():
        return "CLOSING"
    if "full-bleed" in notes.lower() or "gallery" in notes.lower() or "photograph" in notes.lower() and "caption" in blob:
        return "PICTURE"
    if "two-point" in notes.lower() or "two points" in blob or "three parallel" in notes.lower() or "three parts" in blob:
        return "TWO_CONTENT"
    if "contents" in title.lower() or "standard content" in notes.lower() or "timeline" in notes.lower() or "people" in notes.lower():
        return "TITLE_BODY"
    if "statement" in notes.lower() or "figures" in notes.lower() or "data table" in notes.lower() or "quote" in notes.lower():
        return "TITLE_ONLY"
    if index == total - 1:
        return "CLOSING"
    return classify_layout_role(title, [])


def discover_prototypes(prs) -> list[dict[str, Any]]:
    prototypes = []
    total = len(prs.slides)
    for index, slide in enumerate(prs.slides):
        notes = ""
        if slide.has_notes_slide:
            notes = (slide.notes_slide.notes_text_frame.text or "").strip()
        title = ""
        texts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                t = (shape.text_frame.text or "").strip()
                if t:
                    texts.append(t.split("\n")[0])
        title = texts[1] if len(texts) > 1 else (texts[0] if texts else f"Slide {index + 1}")
        if texts and texts[0].isupper() and len(texts) > 1:
            name = texts[0].title()
        else:
            name = title[:60] or f"Slide {index + 1}"
        role = _prototype_role(index, " ".join(texts[:3]), notes, total)
        if role is None:
            continue
        prototypes.append(
            {
                "index": index,
                "name": name,
                "role": role,
                "notes": notes.split("\n")[0][:160],
            }
        )
    return prototypes


def layout_map_from_prototypes(prototypes: list[dict[str, Any]]) -> dict[str, int]:
    fake_layouts = [{"index": p["index"], "name": p["name"], "role": p["role"], "placeholders": [{"type_name": "TITLE"}]} for p in prototypes]
    # BLANK: use a content-like prototype if no blank exists
    mapping = resolve_layout_map(fake_layouts)
    return mapping


def detect_colorways(prs, layouts: list[dict[str, Any]], theme: dict[str, Any], prototypes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return green and blue colorway records.

    If the package actually has distinct masters/themes, each becomes a colorway
    with its own layout_map. Otherwise both colorways share sample-slide
    prototypes and differ by theme tokens + sRGB remap (lime ↔ teal).
    """
    major, minor = _sample_fonts(prs)
    proto_map = layout_map_from_prototypes(prototypes) if prototypes else resolve_layout_map(layouts)
    layout_names = {}
    for proto in prototypes:
        layout_names[str(proto["index"])] = proto["name"]
    for layout in layouts:
        layout_names.setdefault(str(layout["index"]), layout.get("name"))

    green_theme = _base_theme(major, minor, LIME, TEAL)
    blue_theme = _base_theme(major, minor, TEAL, LIME)
    # Official EF slides hardcode Questrial / Arial; keep those over a generic theme part.

    shared = {
        "layout_map": proto_map,
        "layout_names": {role: layout_names.get(str(idx), role) for role, idx in proto_map.items()},
        "prototype_indices": [p["index"] for p in prototypes],
        "build_mode": "prototypes" if prototypes else "layouts",
    }
    return [
        {
            "id": "green",
            "label": "Green",
            "description": "Sawgrass lime accent (C1D451)",
            "theme": green_theme,
            "srgb_map": {},
            **shared,
        },
        {
            "id": "blue",
            "label": "Blue",
            "description": "Water teal accent (00ACBF)",
            "theme": blue_theme,
            "srgb_map": {LIME: TEAL, TEAL: LIME, "82C769": DEEP_TEAL},
            **shared,
        },
    ]


def apply_colorway(tokens: dict[str, Any], colorway_id: str | None) -> dict[str, Any]:
    tokens = dict(tokens)
    colorways = tokens.get("colorways") or []
    chosen_id = (colorway_id or tokens.get("default_colorway") or "green").lower()
    match = next((c for c in colorways if c.get("id") == chosen_id), None)
    if match is None and colorways:
        match = colorways[0]
        chosen_id = match["id"]
    if match is None:
        tokens["colorway"] = chosen_id
        return tokens
    tokens["colorway"] = match["id"]
    tokens["theme"] = match.get("theme") or tokens.get("theme")
    tokens["layout_map"] = match.get("layout_map") or tokens.get("layout_map")
    tokens["layout_names"] = match.get("layout_names") or tokens.get("layout_names")
    tokens["srgb_map"] = match.get("srgb_map") or {}
    tokens["build_mode"] = match.get("build_mode") or tokens.get("build_mode") or "layouts"
    tokens["colorway_label"] = match.get("label") or match["id"]
    brand = dict(tokens.get("brand_md") or {})
    brand["color_whitelist"] = sorted(set((brand.get("color_whitelist") or []) + EF_WHITELIST))
    extra_fonts = [
        (match.get("theme") or {}).get("fonts", {}).get("major"),
        (match.get("theme") or {}).get("fonts", {}).get("minor"),
        "Questrial",
        "Arial",
        "Arial Nova",
        "Poppins",
    ]
    brand["fonts"] = sorted({f for f in list(brand.get("fonts") or []) + extra_fonts if f})
    tokens["brand_md"] = brand
    return tokens


def remap_srgb_in_slide(slide, srgb_map: dict[str, str]) -> None:
    if not srgb_map:
        return
    mapping = {hex_color(k): hex_color(v) for k, v in srgb_map.items()}
    for node in slide._element.findall(f".//{{{A_NS}}}srgbClr"):
        val = hex_color(node.get("val") or "")
        if val in mapping:
            node.set("val", mapping[val])


def patch_theme_part(prs, theme: dict[str, Any]) -> None:
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    colors = (theme or {}).get("colors") or {}
    fonts = (theme or {}).get("fonts") or {}
    master = prs.slide_masters[0]
    part = None
    for rel in master.part.rels.values():
        if rel.reltype == RT.THEME:
            part = rel.target_part
            break
    if part is None:
        return
    root = etree.fromstring(part.blob)
    scheme = root.find(f".//{{{A_NS}}}clrScheme")
    if scheme is not None:
        for name, hex_value in colors.items():
            node = scheme.find(f"{{{A_NS}}}{name}")
            if node is None:
                continue
            for child in list(node):
                node.remove(child)
            etree.SubElement(node, f"{{{A_NS}}}srgbClr").set("val", hex_color(hex_value))
    if fonts.get("major") or fonts.get("minor"):
        major = root.find(f".//{{{A_NS}}}majorFont")
        minor = root.find(f".//{{{A_NS}}}minorFont")
        if major is not None and fonts.get("major"):
            latin = major.find(f"{{{A_NS}}}latin")
            if latin is not None:
                latin.set("typeface", fonts["major"])
        if minor is not None and fonts.get("minor"):
            latin = minor.find(f"{{{A_NS}}}latin")
            if latin is not None:
                latin.set("typeface", fonts["minor"])
    part.blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def duplicate_slide(prs, index: int):
    """Clone a template sample slide inside the same package (not from SOURCE)."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)
    src_cSld = source._element.find(qn("p:cSld"))
    dest_cSld = dest._element.find(qn("p:cSld"))
    if src_cSld is not None and dest_cSld is not None:
        src_bg = src_cSld.find(qn("p:bg"))
        dest_bg = dest_cSld.find(qn("p:bg"))
        if src_bg is not None:
            if dest_bg is not None:
                dest_cSld.remove(dest_bg)
            copied = deepcopy(src_bg)
            dest_cSld.insert(0, copied)
            _rewire_blips(copied, source.part, dest)
    src_ovr = source._element.find(qn("p:clrMapOvr"))
    dest_ovr = dest._element.find(qn("p:clrMapOvr"))
    if src_ovr is not None:
        copied_ovr = deepcopy(src_ovr)
        if dest_ovr is not None:
            dest._element.replace(dest_ovr, copied_ovr)
        else:
            dest._element.append(copied_ovr)
    for shape in list(dest.shapes):
        el = shape._element
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    source_part = source.part
    for shape in source.shapes:
        new_el = deepcopy(shape._element)
        _rewire_blips(new_el, source_part, dest)
        dest.shapes._spTree.append(new_el)
    return dest


def _rewire_blips(element, source_part, dest_slide) -> None:
    embed_attr = f"{{{R_NS}}}embed"
    for blip in element.findall(f".//{{{A_NS}}}blip"):
        rid = blip.get(embed_attr)
        if not rid:
            continue
        try:
            image_part = source_part.related_part(rid)
            blob = image_part.blob
        except Exception:
            continue
        try:
            _image_part, new_rid = dest_slide.part.get_or_add_image_part(io.BytesIO(blob))
            blip.set(embed_attr, new_rid)
        except Exception:
            continue


def prototype_text_slots(slide) -> list[dict[str, Any]]:
    slots = []
    for shape, left, top in iter_shapes_abs(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if not getattr(shape, "has_text_frame", False):
            continue
        text = (shape.text_frame.text or "").strip()
        if not text:
            continue
        slots.append(
            {
                "shape": shape,
                "left": left,
                "top": top,
                "width": int(shape.width),
                "height": int(shape.height),
                "text": text,
            }
        )
    slots.sort(key=lambda s: (int(s["top"]), int(s["left"])))
    return slots


def replace_shape_text(shape, text: str) -> None:
    if not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.word_wrap = True
    first = tf.paragraphs[0]
    if first.runs:
        first.runs[0].text = text
        for run in first.runs[1:]:
            run.text = ""
    else:
        first.text = text
    for para in tf.paragraphs[1:]:
        for run in para.runs:
            run.text = ""


def fill_text_slots(slide, title: str, subtitle: str, body_paras: list[str]) -> None:
    slots = prototype_text_slots(slide)
    if not slots:
        return
    # Kicker = short uppercase first line; title = next; rest = body.
    kicker = None
    rest = list(slots)
    first = rest[0]
    if len(first["text"]) <= 24 and first["text"] == first["text"].upper() and not first["text"].isdigit():
        kicker = rest.pop(0)
    title_slot = rest.pop(0) if rest else None
    if title_slot is None and kicker is not None:
        title_slot = kicker
        kicker = None
    if title_slot is not None and title:
        replace_shape_text(title_slot["shape"], title)
    elif title_slot is not None:
        replace_shape_text(title_slot["shape"], title_slot["text"] if not title else title)
    body_texts = []
    if subtitle:
        body_texts.append(subtitle)
    body_texts.extend([p for p in body_paras if p])
    # TWO_CONTENT: split remaining slots by x midpoint
    if len(rest) >= 2 and len(body_texts) >= 2:
        mid = (min(s["left"] for s in rest) + max(s["left"] + s["width"] for s in rest)) / 2
        left_slots = [s for s in rest if s["left"] + s["width"] / 2 < mid]
        right_slots = [s for s in rest if s not in left_slots]
        half = max(1, len(body_texts) // 2)
        _fill_slot_group(left_slots, body_texts[:half])
        _fill_slot_group(right_slots, body_texts[half:])
        return
    _fill_slot_group(rest, body_texts)


def _fill_slot_group(slots: list[dict[str, Any]], texts: list[str]) -> None:
    if not slots:
        return
    if not texts:
        # Leave sample copy rather than emptying chrome labels like 01 / COLUMN ONE
        return
    for i, slot in enumerate(slots):
        if i < len(texts):
            value = texts[i]
            if i == len(slots) - 1 and len(texts) > len(slots):
                value = "\n".join(texts[i:])
            replace_shape_text(slot["shape"], value)
        # extra decorative slots keep template copy


def strip_pic_locks(element) -> None:
    """Remove picLocks (noChangeAspect etc.) so staff can crop, resize, and Change Picture."""
    for locks in list(element.findall(f".//{{{A_NS}}}picLocks")):
        parent = locks.getparent()
        if parent is not None:
            parent.remove(locks)


def _is_picture_placeholder(shape) -> bool:
    try:
        if getattr(shape, "is_placeholder", False):
            name = (enum_name(placeholder_type(shape)) or "").upper()
            if name in {"PICTURE", "MEDIA_CLIP"}:
                return True
    except Exception:
        pass
    nv_pr = shape._element.find(f".//{{{P_NS}}}nvPr")
    if nv_pr is None:
        return False
    ph = nv_pr.find(qn("p:ph"))
    if ph is None:
        return False
    return (ph.get("type") or "").lower() in {"pic", "clipart", "media"}


def cover_src_rect(img_w: int, img_h: int, box_w: int, box_h: int) -> dict[str, int]:
    if not img_w or not img_h or not box_w or not box_h:
        return {"left": 0, "right": 0, "top": 0, "bottom": 0}
    img_aspect = img_w / img_h
    box_aspect = box_w / box_h
    if img_aspect > box_aspect:
        visible = box_aspect / img_aspect
        cut = max(0.0, (1 - visible) / 2)
        units = int(cut * 100000)
        return {"left": units, "right": units, "top": 0, "bottom": 0}
    visible = img_aspect / box_aspect
    cut = max(0.0, (1 - visible) / 2)
    units = int(cut * 100000)
    return {"left": 0, "right": 0, "top": units, "bottom": units}


def _clear_src_rect(shape) -> None:
    for src in list(shape._element.findall(f".//{{{A_NS}}}srcRect")):
        parent = src.getparent()
        if parent is not None:
            parent.remove(src)


def _swap_blip(shape, blob: bytes) -> None:
    _image_part, rId = shape.part.get_or_add_image_part(io.BytesIO(blob))
    embed_attr = f"{{{R_NS}}}embed"
    link_attr = f"{{{R_NS}}}link"
    for blip in shape._element.findall(f".//{{{A_NS}}}blip"):
        blip.set(embed_attr, rId)
        if link_attr in blip.attrib:
            del blip.attrib[link_attr]


def replace_picture_in_place(shape, blob: bytes, img_w: int | None, img_h: int | None) -> bool:
    """Put a source blob into an existing p:pic / picture placeholder. Keep the frame."""
    if _is_picture_placeholder(shape) and hasattr(shape, "insert_picture"):
        try:
            inserted = shape.insert_picture(io.BytesIO(blob))
            shape = inserted
        except Exception:
            _swap_blip(shape, blob)
    else:
        _swap_blip(shape, blob)
    strip_pic_locks(shape._element)
    _clear_src_rect(shape)
    if img_w and img_h:
        apply_picture_crop(shape, cover_src_rect(int(img_w), int(img_h), int(shape.width or 1), int(shape.height or 1)))
    return True


def content_photo_frames(slide, slide_w: int, slide_h: int) -> list[dict[str, Any]]:
    """Template content photo frames (not corner logos / master chrome)."""
    recs: list[dict[str, Any]] = []
    slide_area = max(1, int(slide_w) * int(slide_h))
    for shape in slide.shapes:
        if shape.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            continue
        left = int(shape.left or 0)
        top = int(shape.top or 0)
        width = int(shape.width or 0)
        height = int(shape.height or 0)
        area = width * height
        recs.append(
            {
                "shape": shape,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "area": area,
                "placeholder": _is_picture_placeholder(shape),
            }
        )
    frames: list[dict[str, Any]] = []
    for rec in recs:
        siblings = [other for other in recs if abs(other["area"] - rec["area"]) / max(rec["area"], 1) <= 0.08]
        if len(siblings) >= 2:
            frames.append(rec)
            continue
        logo = is_edge_logo(rec["area"], slide_area, rec["left"], rec["top"], rec["width"], rec["height"], slide_w, slide_h)
        if logo:
            continue
        if rec["area"] >= slide_area * 0.05:
            frames.append(rec)
    frames.sort(
        key=lambda item: (
            0 if item["placeholder"] else 1,
            -int(item["area"]),
            int(item["left"]),
            int(item["top"]),
        )
    )
    return frames


def place_source_photos(
    slide,
    placements: list[dict[str, Any]],
    tokens: dict[str, Any],
    flags: list[str],
) -> None:
    """Replace content photo frames in place. Never stack a floating pic over stock art."""
    sources: list[dict[str, Any]] = []
    for placement in placements:
        image = placement.get("image") or {}
        path = image.get("path")
        if path and Path(path).exists():
            sources.append(placement)
    slide_w = int(tokens["slide_width"])
    slide_h = int(tokens["slide_height"])
    frames = content_photo_frames(slide, slide_w, slide_h)
    used: set[int] = set()
    for index, placement in enumerate(sources):
        image = placement.get("image") or {}
        blob = Path(image["path"]).read_bytes()
        blob, _ext = maybe_recompress(blob, image.get("ext") or "png")
        px = image.get("px") or [None, None]
        img_w, img_h = px[0], px[1]
        if not img_w or not img_h:
            try:
                from PIL import Image

                with Image.open(io.BytesIO(blob)) as im:
                    img_w, img_h = im.size
            except Exception:
                img_w, img_h = None, None
        if index < len(frames):
            shape = frames[index]["shape"]
            replace_picture_in_place(shape, blob, img_w, img_h)
            used.add(id(shape))
            continue
        # Extra source photo: unlocked p:pic in the content box, stock frames already consumed.
        box = content_safe_box(tokens, tokens.get("_layout_index"))
        if img_w and img_h:
            fit = contain_fit(int(img_w), int(img_h), box)
        else:
            fit = box
        picture = slide.shapes.add_picture(io.BytesIO(blob), int(fit["left"]), int(fit["top"]), int(fit["width"]), int(fit["height"]))
        strip_pic_locks(picture._element)
        flags.append("PICTURE_OVERFLOW")
    if sources:
        for frame in frames:
            if id(frame["shape"]) not in used:
                delete_shape(frame["shape"])


def unlock_slide_pictures(slide) -> None:
    """Staff can Change Picture / crop / resize every native pic on the slide."""
    for shape, _left, _top in iter_shapes_abs(slide.shapes):
        if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE}:
            strip_pic_locks(shape._element)


def replace_largest_picture(slide, image_path: str, slide_w: int, slide_h: int) -> bool:
    """Back-compat: put one image into the largest content photo frame."""
    frames = content_photo_frames(slide, slide_w, slide_h)
    if not frames:
        return False
    blob = Path(image_path).read_bytes()
    img_w = img_h = None
    try:
        from PIL import Image

        with Image.open(io.BytesIO(blob)) as im:
            img_w, img_h = im.size
    except Exception:
        pass
    replace_picture_in_place(frames[0]["shape"], blob, img_w, img_h)
    for frame in frames[1:]:
        delete_shape(frame["shape"])
    return True

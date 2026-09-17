#!/usr/bin/env python3
"""JSON CLI used by the Astro staff app. Does not go through converter/convert.py."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
sys.path.insert(0, str(ROOT))

from helpers import ROLE_LABELS, ROLES, load_json  # noqa: E402
from inspect_source import inspect_source  # noqa: E402
from inspect_template import inspect_template  # noqa: E402
from plan import build_plan  # noqa: E402
from helpers import dump_json  # noqa: E402
from colorways import apply_colorway  # noqa: E402


def parse_payload(tokens: dict[str, Any], manifest: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    layouts = [
        {
            "index": int(layout["index"]),
            "name": layout.get("name"),
            "role": layout.get("role"),
        }
        for layout in tokens.get("layouts") or []
    ]
    layout_by_index = {int(item["index"]): item for item in layouts}
    slides_out = []
    for slide in manifest.get("slides") or []:
        kinds = sorted({s.get("kind") for s in slide.get("shapes") or [] if s.get("kind") and s.get("kind") != "group"})
        flags = []
        if slide.get("has_transitions_or_animations"):
            flags.append("ANIMATIONS_DROPPED")
        for shape in slide.get("shapes") or []:
            if shape.get("drop_logo"):
                flags.append("LOGO_DROPPED")
            if shape.get("kind") == "smartart":
                flags.append("SMARTART_FLATTENED")
            if shape.get("kind") == "video":
                flags.append("MEDIA_MISSING")
            if shape.get("kind") == "chart":
                chart = shape.get("chart") or {}
                if chart.get("external") or not chart.get("supported"):
                    flags.append("NEEDS_MANUAL_REBUILD")
        entry = next((e for e in plan.get("entries") or [] if int(e["source_index"]) == int(slide["index"])), None)
        preview_parts = []
        for shape in slide.get("shapes") or []:
            if shape.get("drop_footer") or shape.get("drop_logo"):
                continue
            text = (shape.get("text") or "").strip()
            if text and text != (slide.get("title_text") or "").strip():
                preview_parts.append(text)
        slides_out.append(
            {
                "index": int(slide["index"]),
                "title": slide.get("title_text") or "",
                "subtitle": slide.get("subtitle_text") or "",
                "bodyPreview": "\n".join(preview_parts)[:800],
                "role": (entry or {}).get("role") or slide.get("role"),
                "templateLayoutIndex": (entry or {}).get("template_layout_index"),
                "templateLayoutName": (entry or {}).get("template_layout_name"),
                "shapeKinds": kinds,
                "hasImage": "picture" in kinds,
                "imageCount": sum(1 for s in slide.get("shapes") or [] if s.get("kind") == "picture" and not s.get("drop_logo")),
                "hasChart": "chart" in kinds,
                "hasTable": "table" in kinds,
                "hasSmartArt": "smartart" in kinds,
                "notes": slide.get("notes_text") or "",
                "hidden": bool(slide.get("hidden")),
                "flags": sorted(set(flags)),
                "warnings": [],
            }
        )
    layout_names = dict(tokens.get("layout_names") or {})
    for colorway in tokens.get("colorways") or []:
        if colorway.get("id") == tokens.get("colorway"):
            layout_names = dict(colorway.get("layout_names") or layout_names)
            break
    if not layout_names:
        layout_names = {
            str(k): layout_by_index.get(int(v), {}).get("name")
            for k, v in (tokens.get("layout_map") or {}).items()
            if str(v).isdigit() and int(v) in layout_by_index
        }
    colorways = [
        {
            "id": item.get("id"),
            "label": item.get("label") or item.get("id"),
            "description": item.get("description") or "",
        }
        for item in tokens.get("colorways") or []
        if item.get("id")
    ]
    if not colorways:
        colorways = [
            {"id": "green", "label": "Green", "description": "Sawgrass lime accent (C1D451)"},
            {"id": "blue", "label": "Blue", "description": "Water teal accent (00ACBF)"},
        ]
    return {
        "fileName": Path(manifest.get("source_path") or "upload.pptx").name,
        "slideCount": int(manifest.get("slide_count") or 0),
        "roles": [{"id": role, "label": ROLE_LABELS[role]} for role in ROLES],
        "layouts": layouts,
        "layoutMap": tokens.get("layout_map") or {},
        "layoutNames": layout_names,
        "slides": slides_out,
        "plan": plan,
        "colorways": colorways,
        "defaultColorway": tokens.get("default_colorway") or "green",
        "selectedColorway": tokens.get("colorway") or tokens.get("default_colorway") or "green",
        "templateNotes": tokens.get("notes") or [],
    }


def cmd_parse(args: argparse.Namespace) -> dict[str, Any]:
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    brand_md = Path(args.brand_md) if args.brand_md else None
    tokens = inspect_template(Path(args.template), out_dir, brand_md, args.colorway)
    tokens = apply_colorway(tokens, args.colorway)
    dump_json(out_dir / "brand_tokens.json", tokens)
    manifest = inspect_source(Path(args.source), out_dir)
    plan = build_plan(tokens, manifest, None, args.colorway)
    dump_json(out_dir / "plan.json", plan)
    return parse_payload(tokens, manifest, plan)


def cmd_convert(args: argparse.Namespace) -> dict[str, Any]:
    from build import build_presentation
    from qa import run_qa

    out_dir = Path(args.out_dir)
    brand_md = Path(args.brand_md) if args.brand_md else None
    if (out_dir / "brand_tokens.json").exists():
        tokens = load_json(out_dir / "brand_tokens.json")
    else:
        tokens = inspect_template(Path(args.template), out_dir, brand_md, args.colorway)
    tokens = apply_colorway(tokens, args.colorway)
    dump_json(out_dir / "brand_tokens.json", tokens)
    manifest = load_json(out_dir / "source_manifest.json") if (out_dir / "source_manifest.json").exists() else inspect_source(Path(args.source), out_dir)
    overrides = None
    if args.overrides:
        raw = Path(args.overrides)
        overrides = json.loads(raw.read_text(encoding="utf-8") if raw.exists() else args.overrides)
        if isinstance(overrides, dict) and "entries" in overrides:
            overrides = overrides["entries"]
    plan = build_plan(tokens, manifest, overrides, args.colorway)
    dump_json(out_dir / "plan.json", plan)
    meta = build_presentation(
        template_path=Path(args.template),
        source_path=Path(args.source),
        out_dir=out_dir,
        tokens=tokens,
        manifest=manifest,
        plan=plan,
        brand_md=brand_md,
        colorway=args.colorway,
    )
    summary = run_qa(out_dir, Path(args.source), Path(meta["output"]))
    report_text = Path(summary["report"]).read_text(encoding="utf-8") if Path(summary["report"]).exists() else ""
    return {
        "output": meta["output"],
        "slideCount": meta["slide_count"],
        "flags": summary.get("flags") or meta.get("flags") or [],
        "warnings": meta.get("warnings") or [],
        "checks": summary.get("checks") or {},
        "reportMarkdown": report_text,
        "parityDiffs": (summary.get("checks") or {}).get("text_parity", {}).get("diffs") or [],
        "plan": plan,
        "colorway": tokens.get("colorway"),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Staff-app JSON bridge for the rebrand engine")
    sub = parser.add_subparsers(dest="cmd", required=True)
    parse_cmd = sub.add_parser("parse")
    parse_cmd.add_argument("source")
    parse_cmd.add_argument("--template", required=True)
    parse_cmd.add_argument("--out-dir", required=True)
    parse_cmd.add_argument("--brand-md", default=None)
    parse_cmd.add_argument("--colorway", default=None)

    convert_cmd = sub.add_parser("convert")
    convert_cmd.add_argument("source")
    convert_cmd.add_argument("--template", required=True)
    convert_cmd.add_argument("--out-dir", required=True)
    convert_cmd.add_argument("--overrides", default=None)
    convert_cmd.add_argument("--brand-md", default=None)
    convert_cmd.add_argument("--colorway", default=None)

    args = parser.parse_args(argv)
    try:
        if args.cmd == "parse":
            result = cmd_parse(args)
        else:
            result = cmd_convert(args)
        json.dump(result, sys.stdout, ensure_ascii=False)
        sys.stdout.write("\n")
        return 0
    except Exception as exc:
        json.dump({"error": str(exc)}, sys.stderr)
        sys.stderr.write("\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())

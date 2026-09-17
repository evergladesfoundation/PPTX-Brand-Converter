# PPTX-Brand-Converter

Internal Everglades Foundation staff tool: upload any PowerPoint, inspect its content, rebuild every slide onto the Foundation template, and download a branded deck plus a QA report.

## Run locally

Requires **Node.js 22.12** or newer and **Python 3.11+**. For visual QA, install **LibreOffice** and **Poppler** (`pdftoppm`).

```sh
python -m venv converter/.venv
# macOS / Linux
converter/.venv/bin/python -m pip install -r converter/requirements.txt
# Windows
converter/.venv\Scripts\python -m pip install -r converter/requirements.txt
npm install
npm run dev
```

Dev server: `http://127.0.0.1:4321`

Optional staff password: copy `.env.example` to `.env` and set `INTERNAL_PASSWORD`.

Place Communications’ official 2023 template at `templates/everglades.pptx`. If you only have a `.potx`, the engine rewrites it to `.pptx` on inspect. Green and blue in that file are **palette colors on the sample slides** (Sawgrass Lime `C1D451` and Water Teal `00ACBF`), not a converter overlay and not Design-tab theme variants. Convert clones those layouts as drawn. A staff picker appears only if a later template actually encodes two selectable design options.

## Commands

| Command            | Action |
| :----------------- | :----- |
| `npm install`      | Install dependencies |
| `npm run dev`      | Dev server at `localhost:4321` |
| `npm run build`    | Production build to `./dist/` |
| `npm run preview`  | Preview the production build locally |
| `npm run template` | Rebuild a python-pptx starter deck (dev only; does not replace the official Communications file) |
| `npm run fixture`  | Rebuild `fixtures/sample.pptx` |
| `rebrand/run.sh`   | Inspect → plan → build → QA on a source deck |

CLI (from repo root, using the converter venv):

```sh
python rebrand/inspect_template.py templates/everglades.pptx --out-dir .tmp/rebrand
python rebrand/inspect_source.py fixtures/sample.pptx --out-dir .tmp/rebrand
python rebrand/plan.py --out-dir .tmp/rebrand
python rebrand/build.py --out-dir .tmp/rebrand
python rebrand/qa.py --out-dir .tmp/rebrand
```

Or: `rebrand/run.sh templates/everglades.pptx fixtures/sample.pptx .tmp/rebrand`

Optional `COLORWAY=` / `--colorway` is only used when the template itself exposes named design options (extra color schemes, extra masters, or labeled prototype series). The 2023 EF file does not.

## Project structure

```text
rebrand/                 inspect → plan → rebuild → QA engine
converter/               template/fixture generators (not the conversion core)
templates/everglades.pptx
src/pages/index.astro    Staff upload UI
src/pages/api/           parse + convert (plan.json + QA flags)
src/layouts/Layout.astro
src/styles/global.css    Everglades color tokens
```

## Tailwind notes

This project uses the official Vite plugin (`@tailwindcss/vite`). Import `src/styles/global.css` from the layout — never from every page.

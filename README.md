# PPTX-Brand-Converter

Internal Everglades Foundation staff tool: upload any PowerPoint, extract titles, body, images, notes, and live Office charts, rebuild them onto the Foundation template, and download a branded deck.

## Run locally

Requires **Node.js 22.12** or newer and **Python 3.11+**.

```sh
python -m venv converter/.venv
converter/.venv/Scripts/python -m pip install -r converter/requirements.txt
npm install
npm run dev
```

On macOS or Linux, use `converter/.venv/bin/python` instead of `Scripts/python`.

Dev server: `http://127.0.0.1:4321`

Optional staff password: copy `.env.example` to `.env` and set `INTERNAL_PASSWORD`.

Place Communications’ official template at `templates/everglades.pptx` (a branded starter is generated with `npm run template`).

## Commands

| Command           | Action                                      |
| :---------------- | :------------------------------------------ |
| `npm install`     | Install dependencies                        |
| `npm run dev`     | Dev server at `localhost:4321`              |
| `npm run build`   | Production build to `./dist/`               |
| `npm run preview` | Preview the production build locally        |
| `npm run template`| Rebuild `templates/everglades.pptx`         |
| `npm run fixture` | Rebuild `fixtures/sample.pptx`              |

## Project structure

```text
converter/               Python parse / rebuild / live-chart copy
templates/everglades.pptx
src/pages/index.astro    Staff upload UI
src/pages/api/           parse + convert
src/layouts/Layout.astro
src/styles/global.css    Everglades color tokens
```

## Tailwind notes

This project uses the official Vite plugin (`@tailwindcss/vite`). Import `src/styles/global.css` from the layout — never from every page.

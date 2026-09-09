# PPTX-Brand-Converter

Internal Everglades Foundation staff tool: upload any PowerPoint, extract titles, body, images, notes, and live Office charts, rebuild them onto the Foundation template, and download a branded deck.

This repo starts from an Astro 7 + Tailwind CSS v4 site. File upload and conversion will run on a Node server with a Python converter.

## Run locally

Requires **Node.js 22.12** or newer.

```sh
npm install
npm run dev
```

Dev server: `localhost:4321`

## Commands

| Command           | Action                                      |
| :---------------- | :------------------------------------------ |
| `npm install`     | Install dependencies                        |
| `npm run dev`     | Dev server at `localhost:4321`              |
| `npm run build`   | Production build to `./dist/`               |
| `npm run preview` | Preview the production build locally        |

## Project structure

```text
src/
  consts.ts              Site title, description, nav
  layouts/Layout.astro   HTML shell, SEO, fonts, theme boot
  components/            Header, Footer, Button, theme toggle
  pages/                 File-based routes
  styles/global.css      Tailwind + @theme tokens
public/                  Favicon, robots.txt
```

## Tailwind notes

This project uses the official Vite plugin (`@tailwindcss/vite`), not the deprecated `@astrojs/tailwind` integration. Import `src/styles/global.css` from the layout — never from every page.

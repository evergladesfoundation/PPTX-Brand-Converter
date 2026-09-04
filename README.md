# Folio — Astro + Tailwind template

A small, reusable starter for static sites. Astro 7, Tailwind CSS v4, CSS-first design tokens, and a base layout with dark mode.

## Use this repo as a template

After you push it to GitHub:

```sh
npm create astro@latest my-site -- --template YOUR_GITHUB_USER/YOUR_REPO
```

Or clone this folder and run:

```sh
npm install
npm run dev
```

Requires **Node.js 22.12** or newer.

## Customize

1. Site name, description, and nav — `src/consts.ts`
2. Colors and fonts — `src/styles/global.css` (`:root` and `.dark`)
3. Homepage — `src/pages/index.astro`
4. Canonical URLs — set `site` in `astro.config.mjs`

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

## Commands

| Command           | Action                                      |
| :---------------- | :------------------------------------------ |
| `npm install`     | Install dependencies                        |
| `npm run dev`     | Dev server at `localhost:4321`              |
| `npm run build`   | Production build to `./dist/`               |
| `npm run preview` | Preview the production build locally        |

## Tailwind notes

This project uses the official Vite plugin (`@tailwindcss/vite`), not the deprecated `@astrojs/tailwind` integration. Import `src/styles/global.css` from the layout — never from every page.

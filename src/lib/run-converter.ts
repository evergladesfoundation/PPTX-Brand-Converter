import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

export const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;
export const SPEC_ROLES = [
  "TITLE",
  "SECTION",
  "TITLE_BODY",
  "TWO_CONTENT",
  "TITLE_ONLY",
  "PICTURE",
  "BLANK",
  "CLOSING",
] as const;

function normalizeColorway(value: string | undefined | null): string | undefined {
  if (typeof value !== "string") return undefined;
  const id = value.trim();
  return id || undefined;
}

const API_PY = join("rebrand", "api.py");
const TEMPLATE_REL = join("templates", "everglades.pptx");

function findRepoRoot(): string {
  const starts = [process.cwd(), fileURLToPath(new URL(".", import.meta.url))];
  for (const start of starts) {
    let dir = start;
    for (let i = 0; i < 10; i++) {
      if (existsSync(join(dir, API_PY))) return dir;
      const parent = join(dir, "..");
      if (parent === dir) break;
      dir = parent;
    }
  }
  throw new Error("Could not find rebrand/api.py from the server process.");
}

const ROOT = findRepoRoot();
const API_PATH = join(ROOT, API_PY);
export const TEMPLATE_PATH = join(ROOT, TEMPLATE_REL);

export function pythonExecutable(): string {
  const windows = join(ROOT, "converter", ".venv", "Scripts", "python.exe");
  const unix = join(ROOT, "converter", ".venv", "bin", "python");
  if (existsSync(windows)) return windows;
  if (existsSync(unix)) return unix;
  return process.platform === "win32" ? "python" : "python3";
}

function runPython(args: string[], timeoutMs: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn(pythonExecutable(), args, {
      cwd: ROOT,
      windowsHide: true,
    });
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("Conversion timed out."));
    }, timeoutMs);
    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) {
        resolve(stdout);
        return;
      }
      try {
        const parsed = JSON.parse(stderr) as { error?: string };
        reject(new Error(parsed.error || stderr.trim() || "Converter failed."));
      } catch {
        reject(new Error(stderr.trim() || stdout.trim() || "Converter failed."));
      }
    });
  });
}

async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), "ef-pptx-"));
  try {
    return await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

function templatePath(): string {
  if (!existsSync(TEMPLATE_PATH)) {
    throw new Error(
      "Missing templates/everglades.pptx. Place Communications’ official 2023 TEMPLATE.pptx there.",
    );
  }
  return TEMPLATE_PATH;
}

export async function parsePptx(bytes: Uint8Array, fileName: string, colorway?: string) {
  const chosen = normalizeColorway(colorway);
  return withTempDir(async (dir) => {
    const input = join(dir, fileName.endsWith(".pptx") ? fileName : "upload.pptx");
    await writeFile(input, bytes);
    const args = [API_PATH, "parse", input, "--template", templatePath(), "--out-dir", dir];
    if (chosen) args.push("--colorway", chosen);
    const stdout = await runPython(args, 120_000);
    return JSON.parse(stdout) as ParseResult;
  });
}

export async function convertPptx(
  bytes: Uint8Array,
  fileName: string,
  overrides: PlanOverride[],
  colorway?: string,
) {
  const chosen = normalizeColorway(colorway);
  return withTempDir(async (dir) => {
    const input = join(dir, "upload.pptx");
    const overridesPath = join(dir, "overrides.json");
    await writeFile(input, bytes);
    await writeFile(overridesPath, JSON.stringify(overrides));
    const args = [
      API_PATH,
      "convert",
      input,
      "--template",
      templatePath(),
      "--out-dir",
      dir,
      "--overrides",
      overridesPath,
    ];
    if (chosen) args.push("--colorway", chosen);
    const stdout = await runPython(args, 300_000);
    const meta = JSON.parse(stdout) as ConvertMeta;
    const file = await readFile(join(dir, "OUTPUT.pptx"));
    const base = fileName.replace(/\.pptx$/i, "") || "presentation";
    const selected = meta.colorway || chosen;
    return {
      bytes: file,
      downloadName: selected ? `${base}-everglades-${selected}.pptx` : `${base}-everglades.pptx`,
      flags: meta.flags ?? [],
      warnings: meta.warnings ?? [],
      checks: meta.checks ?? {},
      reportMarkdown: meta.reportMarkdown ?? "",
      parityDiffs: meta.parityDiffs ?? [],
      slideCount: meta.slideCount ?? 0,
      colorway: selected ?? null,
    };
  });
}

export type PlanOverride = {
  source_index: number;
  role?: string;
  template_layout_index?: number;
};

export type ParseResult = {
  fileName: string;
  slideCount: number;
  roles: { id: string; label: string }[];
  layouts: { index: number; name: string; role: string }[];
  layoutMap: Record<string, number>;
  layoutNames: Record<string, string>;
  slides: {
    index: number;
    title: string;
    subtitle: string;
    bodyPreview: string;
    role: string;
    templateLayoutIndex: number;
    templateLayoutName: string;
    shapeKinds: string[];
    hasImage: boolean;
    imageCount: number;
    hasChart: boolean;
    hasTable: boolean;
    hasSmartArt: boolean;
    notes: string;
    hidden: boolean;
    flags: string[];
    warnings: string[];
  }[];
  plan: { entries: unknown[]; warnings?: string[]; colorway?: string | null };
  colorways?: { id: string; label: string; description?: string }[];
  defaultColorway?: string | null;
  selectedColorway?: string | null;
  palette?: { label: string; hex: string }[];
  templateNotes?: string[];
};

type ConvertMeta = {
  flags?: string[];
  warnings?: string[];
  checks?: Record<string, { pass?: boolean; justified?: boolean; error?: string }>;
  reportMarkdown?: string;
  parityDiffs?: { source_index: number; source: string; output: string }[];
  slideCount?: number;
  colorway?: string | null;
};

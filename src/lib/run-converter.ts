import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

export const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;

const CONVERT_NAME = join("converter", "convert.py");

function findRepoRoot(): string {
  const starts = [process.cwd(), fileURLToPath(new URL(".", import.meta.url))];
  for (const start of starts) {
    let dir = start;
    for (let i = 0; i < 10; i++) {
      if (existsSync(join(dir, CONVERT_NAME))) return dir;
      const parent = join(dir, "..");
      if (parent === dir) break;
      dir = parent;
    }
  }
  throw new Error("Could not find converter/convert.py from the server process.");
}

const ROOT = findRepoRoot();
const CONVERT_PY = join(ROOT, CONVERT_NAME);

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

export async function parsePptx(bytes: Uint8Array, fileName: string) {
  return withTempDir(async (dir) => {
    const input = join(dir, fileName.endsWith(".pptx") ? fileName : "upload.pptx");
    await writeFile(input, bytes);
    const stdout = await runPython([CONVERT_PY, "parse", input], 60_000);
    return JSON.parse(stdout) as ParseResult;
  });
}

export async function convertPptx(
  bytes: Uint8Array,
  fileName: string,
  layouts: string[],
) {
  return withTempDir(async (dir) => {
    const input = join(dir, "upload.pptx");
    const output = join(dir, "everglades.pptx");
    const layoutsPath = join(dir, "layouts.json");
    await writeFile(input, bytes);
    await writeFile(layoutsPath, JSON.stringify(layouts));
    const stdout = await runPython(
      [CONVERT_PY, "convert", input, output, "--layouts", layoutsPath],
      120_000,
    );
    const meta = JSON.parse(stdout) as { warnings?: string[] };
    const file = await readFile(output);
    const base = fileName.replace(/\.pptx$/i, "") || "presentation";
    return {
      bytes: file,
      downloadName: `${base}-everglades.pptx`,
      warnings: meta.warnings ?? [],
    };
  });
}

export type ParseResult = {
  fileName: string;
  slideCount: number;
  layouts: { id: string; label: string }[];
  slides: {
    index: number;
    title: string;
    body: string[];
    bodyPreview: string;
    hasImage: boolean;
    imageCount: number;
    hasChart: boolean;
    externalChart: boolean;
    hasTable: boolean;
    notes: string;
    suggestedLayout: string;
    warnings: string[];
  }[];
};

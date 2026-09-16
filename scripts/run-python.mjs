#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(fileURLToPath(new URL(".", import.meta.url)), "..");
const win = process.platform === "win32";
const venv = win
  ? join(root, "converter", ".venv", "Scripts", "python.exe")
  : join(root, "converter", ".venv", "bin", "python");
const python = existsSync(venv) ? venv : win ? "python" : "python3";
const result = spawnSync(python, process.argv.slice(2), {
  cwd: root,
  stdio: "inherit",
  windowsHide: true,
});
process.exit(result.status ?? 1);

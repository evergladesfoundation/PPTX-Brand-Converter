import type { APIRoute } from "astro";
import {
  convertPptx,
  MAX_UPLOAD_BYTES,
  SPEC_ROLES,
  type PlanOverride,
} from "../../lib/run-converter";

export const prerender = false;

const ROLES = new Set<string>(SPEC_ROLES);

export const POST: APIRoute = async ({ request }) => {
  const form = await request.formData();
  const file = form.get("file");
  const planRaw = form.get("plan") ?? form.get("overrides");
  if (!(file instanceof File)) {
    return jsonError("Choose a .pptx file.", 400);
  }
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    return jsonError("Only .pptx files are supported for the source deck.", 400);
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    return jsonError("File is over the 50 MB limit.", 413);
  }
  if (typeof planRaw !== "string") {
    return jsonError("Missing plan overrides.", 400);
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(planRaw);
  } catch {
    return jsonError("Plan overrides must be JSON.", 400);
  }
  const entries = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === "object" && Array.isArray((parsed as { entries?: unknown }).entries)
      ? (parsed as { entries: unknown[] }).entries
      : null;
  if (!entries) {
    return jsonError("Plan must be an array of {source_index, role} entries.", 400);
  }
  const overrides: PlanOverride[] = [];
  for (const item of entries) {
    if (!item || typeof item !== "object") {
      return jsonError("One or more plan entries are not valid.", 400);
    }
    const rec = item as Record<string, unknown>;
    const sourceIndex = rec.source_index ?? rec.sourceIndex;
    if (typeof sourceIndex !== "number") {
      return jsonError("Each plan entry needs a source_index.", 400);
    }
    const role = rec.role;
    if (role != null && (typeof role !== "string" || !ROLES.has(role))) {
      return jsonError("One or more layout roles are not valid.", 400);
    }
    const layoutIndex = rec.template_layout_index ?? rec.templateLayoutIndex;
    overrides.push({
      source_index: sourceIndex,
      role: typeof role === "string" ? role : undefined,
      template_layout_index: typeof layoutIndex === "number" ? layoutIndex : undefined,
    });
  }

  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const result = await convertPptx(bytes, file.name, overrides);
    return new Response(
      JSON.stringify({
        downloadName: result.downloadName,
        pptxBase64: Buffer.from(result.bytes).toString("base64"),
        flags: result.flags,
        warnings: result.warnings,
        checks: result.checks,
        reportMarkdown: result.reportMarkdown,
        parityDiffs: result.parityDiffs,
        slideCount: result.slideCount,
      }),
      {
        headers: {
          "Content-Type": "application/json",
          "X-Conversion-Warnings": encodeURIComponent(JSON.stringify(result.warnings)),
          "X-QA-Flags": encodeURIComponent(JSON.stringify(result.flags)),
        },
      },
    );
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Could not convert this PowerPoint.";
    return jsonError(message, 500);
  }
};

function jsonError(error: string, status: number) {
  return new Response(JSON.stringify({ error }), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

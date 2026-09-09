import type { APIRoute } from "astro";
import { convertPptx, MAX_UPLOAD_BYTES } from "../../lib/run-converter";

export const prerender = false;

const LAYOUT_IDS = new Set([
  "title",
  "section",
  "content",
  "two_column",
  "image",
  "quote",
  "chart",
]);

export const POST: APIRoute = async ({ request }) => {
  const form = await request.formData();
  const file = form.get("file");
  const layoutsRaw = form.get("layouts");
  if (!(file instanceof File)) {
    return jsonError("Choose a .pptx file.", 400);
  }
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    return jsonError("Only .pptx files are supported.", 400);
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    return jsonError("File is over the 50 MB limit.", 413);
  }
  if (typeof layoutsRaw !== "string") {
    return jsonError("Missing layout choices.", 400);
  }

  let layouts: unknown;
  try {
    layouts = JSON.parse(layoutsRaw);
  } catch {
    return jsonError("Layout choices must be JSON.", 400);
  }
  if (
    !Array.isArray(layouts) ||
    layouts.some((id) => typeof id !== "string" || !LAYOUT_IDS.has(id))
  ) {
    return jsonError("One or more layout ids are not valid.", 400);
  }

  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const result = await convertPptx(bytes, file.name, layouts as string[]);
    return new Response(new Uint8Array(result.bytes), {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${result.downloadName}"`,
        "X-Conversion-Warnings": encodeURIComponent(JSON.stringify(result.warnings)),
      },
    });
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

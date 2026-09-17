import type { APIRoute } from "astro";
import { MAX_UPLOAD_BYTES, parsePptx } from "../../lib/run-converter";

export const prerender = false;

function asError(error: unknown): string {
  return error instanceof Error ? error.message : "Could not read this PowerPoint.";
}

export const POST: APIRoute = async ({ request }) => {
  const form = await request.formData();
  const file = form.get("file");
  if (!(file instanceof File)) {
    return jsonError("Choose a .pptx file.", 400);
  }
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    return jsonError("Only .pptx files are supported for the source deck.", 400);
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    return jsonError("File is over the 50 MB limit.", 413);
  }
  const colorwayRaw = form.get("colorway");
  const colorway = typeof colorwayRaw === "string" && colorwayRaw.trim() ? colorwayRaw.trim() : undefined;

  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const result = await parsePptx(bytes, file.name, colorway);
    return new Response(JSON.stringify(result), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (error) {
    return jsonError(asError(error), 500);
  }
};

function jsonError(error: string, status: number) {
  return new Response(JSON.stringify({ error }), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

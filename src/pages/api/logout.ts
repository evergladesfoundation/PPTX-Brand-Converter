import type { APIRoute } from "astro";
import { cookieName } from "../../lib/auth";

export const prerender = false;

export const POST: APIRoute = async ({ cookies }) => {
  cookies.delete(cookieName(), { path: "/" });
  return new Response(JSON.stringify({ ok: true }), {
    headers: { "Content-Type": "application/json" },
  });
};

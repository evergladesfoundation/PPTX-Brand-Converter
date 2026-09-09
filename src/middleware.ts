import { defineMiddleware } from "astro:middleware";
import { cookieName, isAuthed, passwordConfigured } from "./lib/auth";

export const onRequest = defineMiddleware(async (context, next) => {
  const required = passwordConfigured();
  const authed = isAuthed(context.cookies.get(cookieName())?.value);
  context.locals.authRequired = required;
  context.locals.authed = authed;

  const path = context.url.pathname;
  const openApi = path === "/api/login";
  if (required && !authed && path.startsWith("/api/") && !openApi) {
    return new Response(JSON.stringify({ error: "Sign in required." }), {
      status: 401,
      headers: { "Content-Type": "application/json" },
    });
  }

  return next();
});

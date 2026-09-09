import { createHmac, timingSafeEqual } from "node:crypto";

const COOKIE = "ef_staff";

function secret(): string | undefined {
  const value = process.env.INTERNAL_PASSWORD?.trim();
  return value ? value : undefined;
}

export function passwordConfigured(): boolean {
  return Boolean(secret());
}

export function signPassword(): string | undefined {
  const value = secret();
  if (!value) return undefined;
  return createHmac("sha256", value).update("everglades-pptx-staff").digest("hex");
}

export function isAuthed(cookieValue: string | undefined): boolean {
  const expected = signPassword();
  if (!expected) return true;
  if (!cookieValue) return false;
  const a = Buffer.from(cookieValue);
  const b = Buffer.from(expected);
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

export function cookieName(): string {
  return COOKIE;
}

export function checkPassword(candidate: string): boolean {
  const value = secret();
  if (!value) return true;
  const a = Buffer.from(candidate);
  const b = Buffer.from(value);
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

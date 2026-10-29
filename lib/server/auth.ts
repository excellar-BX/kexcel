// Server-only: passwords, sessions and the signed-in user.
import { createHash, randomBytes, scrypt as scryptCb, timingSafeEqual } from "node:crypto";
import { promisify } from "node:util";
import { cookies } from "next/headers";
import { db } from "./db";

const scrypt = promisify(scryptCb) as (password: string, salt: Buffer, keylen: number) => Promise<Buffer>;

export const SESSION_COOKIE = "kexcel_session";
const SESSION_DAYS = 30;

export interface User {
  id: string;
  email: string;
  name: string;
}

export async function hashPassword(password: string): Promise<string> {
  const salt = randomBytes(16);
  const hash = await scrypt(password, salt, 64);
  return `scrypt$${salt.toString("base64")}$${hash.toString("base64")}`;
}

export async function verifyPassword(password: string, stored: string): Promise<boolean> {
  const [scheme, salt, hash] = stored.split("$");
  if (scheme !== "scrypt" || !salt || !hash) return false;
  const expected = Buffer.from(hash, "base64");
  const actual = await scrypt(password, Buffer.from(salt, "base64"), expected.length);
  return actual.length === expected.length && timingSafeEqual(actual, expected);
}

/** Only a hash of the session token is stored, so a leaked database can't be used to sign in. */
const hashToken = (token: string) => createHash("sha256").update(token).digest("hex");

export const newId = (bytes = 9) => randomBytes(bytes).toString("base64url");

export async function startSession(userId: string) {
  const token = randomBytes(32).toString("base64url");
  const now = Date.now();
  const expires = now + SESSION_DAYS * 86_400_000;
  db().prepare("insert into sessions (token_hash, user_id, expires_at, created_at) values (?, ?, ?, ?)").run(
    hashToken(token),
    userId,
    expires,
    now,
  );
  (await cookies()).set(SESSION_COOKIE, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    expires: new Date(expires),
  });
}

export async function endSession() {
  const store = await cookies();
  const token = store.get(SESSION_COOKIE)?.value;
  if (token) db().prepare("delete from sessions where token_hash = ?").run(hashToken(token));
  store.delete(SESSION_COOKIE);
}

/** The signed-in user, or null. Expired sessions are cleaned up as they're seen. */
export async function currentUser(): Promise<User | null> {
  const token = (await cookies()).get(SESSION_COOKIE)?.value;
  if (!token) return null;
  const row = db()
    .prepare(
      "select u.id, u.email, u.name, s.expires_at from sessions s join users u on u.id = s.user_id where s.token_hash = ?",
    )
    .get(hashToken(token)) as { id: string; email: string; name: string; expires_at: number } | undefined;
  if (!row) return null;
  if (row.expires_at < Date.now()) {
    db().prepare("delete from sessions where token_hash = ?").run(hashToken(token));
    return null;
  }
  return { id: row.id, email: row.email, name: row.name };
}

/**
 * Rejects cross-site form posts: requests that change data must come from
 * this site's own pages (same Origin) and send JSON.
 */
export function sameOrigin(request: Request): boolean {
  const origin = request.headers.get("origin");
  const host = request.headers.get("x-forwarded-host") ?? request.headers.get("host");
  if (origin && host && new URL(origin).host !== host) return false;
  return (request.headers.get("content-type") ?? "").includes("application/json");
}

// Slows down password guessing: a few attempts per address and email per window.
const attempts = new Map<string, number[]>();
export function tooManyAttempts(key: string, max = 8, windowMs = 15 * 60_000): boolean {
  const now = Date.now();
  const recent = (attempts.get(key) ?? []).filter((t) => now - t < windowMs);
  recent.push(now);
  attempts.set(key, recent);
  return recent.length > max;
}

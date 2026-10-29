import { hashPassword, newId, sameOrigin, startSession, tooManyAttempts } from "@/lib/server/auth";
import { db } from "@/lib/server/db";
import { clientAddress, fail, json, readJson } from "@/lib/server/http";

const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export async function POST(request: Request) {
  if (!sameOrigin(request)) return fail("That request came from somewhere else.", 403);
  const body = await readJson<{ email?: string; password?: string; name?: string }>(request, 10_000);
  const email = body?.email?.trim().toLowerCase() ?? "";
  const password = body?.password ?? "";
  const name = body?.name?.trim().slice(0, 80) || email.split("@")[0];
  if (!EMAIL.test(email) || email.length > 200) return fail("Enter a valid email address.", 400);
  if (password.length < 8) return fail("Use a password of at least 8 characters.", 400);
  if (password.length > 200) return fail("That password is too long.", 400);
  if (tooManyAttempts(`signup:${clientAddress(request)}`, 10, 60 * 60_000)) {
    return fail("Too many new accounts from here. Try again later.", 429);
  }

  if (db().prepare("select 1 from users where email = ?").get(email)) {
    return fail("There's already an account with that email. Sign in instead.", 409);
  }
  const id = newId();
  db()
    .prepare("insert into users (id, email, name, password_hash, created_at) values (?, ?, ?, ?, ?)")
    .run(id, email, name, await hashPassword(password), Date.now());
  await startSession(id);
  return json({ user: { id, email, name } }, 201);
}

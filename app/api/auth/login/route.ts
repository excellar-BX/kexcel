import { sameOrigin, startSession, tooManyAttempts, verifyPassword } from "@/lib/server/auth";
import { db } from "@/lib/server/db";
import { clientAddress, fail, json, readJson } from "@/lib/server/http";

export async function POST(request: Request) {
  if (!sameOrigin(request)) return fail("That request came from somewhere else.", 403);
  const body = await readJson<{ email?: string; password?: string }>(request, 10_000);
  const email = body?.email?.trim().toLowerCase() ?? "";
  const password = body?.password ?? "";
  if (tooManyAttempts(`login:${clientAddress(request)}:${email}`)) {
    return fail("Too many tries. Wait a few minutes and try again.", 429);
  }
  const user = db().prepare("select id, email, name, password_hash from users where email = ?").get(email) as
    | { id: string; email: string; name: string; password_hash: string }
    | undefined;
  // The same message either way, so the form doesn't reveal which emails have accounts.
  if (!user || !(await verifyPassword(password, user.password_hash))) {
    return fail("That email and password don't match.", 401);
  }
  await startSession(user.id);
  return json({ user: { id: user.id, email: user.email, name: user.name } });
}

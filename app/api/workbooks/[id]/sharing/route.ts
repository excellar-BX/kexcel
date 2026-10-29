import { currentUser, sameOrigin } from "@/lib/server/auth";
import { fail, json, readJson } from "@/lib/server/http";
import { setLinkAccess, shareWith, sharing, unshare, type ShareResult } from "@/lib/server/workbooks";

type Context = { params: Promise<{ id: string }> };

const done = (r: ShareResult) => ("ok" in r ? json({ ok: true }) : fail(r.error, r.status));

/** Who the spreadsheet is shared with, and whether its link opens it. Owner only. */
export async function GET(_request: Request, { params }: Context) {
  const user = await currentUser();
  if (!user) return fail("Sign in first.", 401);
  const result = sharing(user.id, (await params).id);
  return result ? json(result) : fail("Only the owner can see who this is shared with.", 403);
}

/** Changes sharing: { email, role } adds or updates a person, { remove: userId } removes one, { linkAccess } sets link viewing. */
export async function POST(request: Request, { params }: Context) {
  if (!sameOrigin(request)) return fail("That request came from somewhere else.", 403);
  const user = await currentUser();
  if (!user) return fail("Sign in first.", 401);
  const { id } = await params;
  const body = await readJson<{ email?: string; role?: string; remove?: string; linkAccess?: string }>(request, 5_000);
  if (!body) return fail("That request wasn't valid.", 400);
  if (body.remove) return done(unshare(user.id, id, body.remove));
  if (body.linkAccess === "none" || body.linkAccess === "view") return done(setLinkAccess(user.id, id, body.linkAccess));
  if (body.email && (body.role === "view" || body.role === "edit")) return done(shareWith(user.id, id, body.email, body.role));
  return fail("Say who to share with and whether they can view or edit.", 400);
}

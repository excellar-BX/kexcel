import { endSession, sameOrigin } from "@/lib/server/auth";
import { fail, json } from "@/lib/server/http";

export async function POST(request: Request) {
  if (!sameOrigin(request)) return fail("That request came from somewhere else.", 403);
  await endSession();
  return json({ ok: true });
}

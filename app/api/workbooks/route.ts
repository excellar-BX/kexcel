import { currentUser } from "@/lib/server/auth";
import { fail, json } from "@/lib/server/http";
import { listWorkbooks } from "@/lib/server/workbooks";

/** The signed-in user's spreadsheets, newest first. */
export async function GET() {
  const user = await currentUser();
  if (!user) return fail("Sign in to see your spreadsheets.", 401);
  return json({ workbooks: listWorkbooks(user.id) });
}

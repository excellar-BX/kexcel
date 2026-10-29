import { currentUser } from "@/lib/server/auth";
import { json } from "@/lib/server/http";

export async function GET() {
  return json({ user: await currentUser() });
}

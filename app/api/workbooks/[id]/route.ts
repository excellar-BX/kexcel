import type { WorkbookJSON } from "@/lib/engine";
import { currentUser, sameOrigin } from "@/lib/server/auth";
import { fail, json, readJson } from "@/lib/server/http";
import { deleteWorkbook, getWorkbook, MAX_WORKBOOK_BYTES, saveWorkbook, workbookVersion } from "@/lib/server/workbooks";

type Context = { params: Promise<{ id: string }> };

const VALID_ID = /^[A-Za-z0-9_-]{4,64}$/;

/** The spreadsheet, for its owner, people it's shared with, or anyone if link viewing is on. ?version=1 returns just the version. */
export async function GET(request: Request, { params }: Context) {
  const user = await currentUser();
  const { id } = await params;
  if (new URL(request.url).searchParams.get("version")) {
    const version = workbookVersion(user?.id ?? null, id);
    return version === null ? fail("Not found.", 404) : json({ version });
  }
  const result = getWorkbook(user?.id ?? null, id);
  if ("missing" in result) return fail("There's no spreadsheet with that link.", 404);
  if ("forbidden" in result) {
    return user ? fail("You don't have access to this spreadsheet. Ask its owner to share it with you.", 403) : fail("Sign in to open this spreadsheet.", 401);
  }
  return json(result.found);
}

/** Creates or updates; `baseVersion` is the version the change was made on (null for a new one). */
export async function PUT(request: Request, { params }: Context) {
  if (!sameOrigin(request)) return fail("That request came from somewhere else.", 403);
  const user = await currentUser();
  if (!user) return fail("Sign in to save to your account.", 401);
  const { id } = await params;
  if (!VALID_ID.test(id)) return fail("That isn't a valid spreadsheet id.", 400);
  const body = await readJson<{ title?: string; workbook?: WorkbookJSON; activeSheet?: string; baseVersion?: number | null }>(
    request,
    MAX_WORKBOOK_BYTES + 10_000,
  );
  if (!body || !body.workbook || !Array.isArray(body.workbook.sheets)) {
    return fail("That spreadsheet is too large or isn't valid.", 400);
  }
  const title = (body.title ?? "").trim().slice(0, 200) || "Untitled spreadsheet";
  const result = saveWorkbook(
    user.id,
    id,
    title,
    { workbook: body.workbook, activeSheet: body.activeSheet },
    body.baseVersion ?? null,
  );
  if ("forbidden" in result) return fail("You can view this spreadsheet but not change it.", 403);
  if ("tooLarge" in result) return fail("This spreadsheet is too large to save to your account.", 413);
  if ("conflict" in result) return json({ conflict: result.conflict }, 409);
  return json(result.saved);
}

export async function DELETE(request: Request, { params }: Context) {
  if (!sameOrigin(request)) return fail("That request came from somewhere else.", 403);
  const user = await currentUser();
  if (!user) return fail("Sign in first.", 401);
  const { id } = await params;
  return deleteWorkbook(user.id, id) ? json({ ok: true }) : fail("Only the owner can delete this spreadsheet.", 403);
}

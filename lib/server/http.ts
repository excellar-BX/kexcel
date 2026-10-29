// Server-only: small helpers shared by the API routes.

export const json = (body: unknown, status = 200) => Response.json(body, { status });
export const fail = (error: string, status: number) => json({ error }, status);

/** Parses a JSON body up to a size limit; null when it's too big or not JSON. */
export async function readJson<T>(request: Request, maxBytes: number): Promise<T | null> {
  const text = await request.text();
  if (text.length > maxBytes) return null;
  try {
    return JSON.parse(text) as T;
  } catch {
    return null;
  }
}

/** The caller's address, for rate limits. */
export const clientAddress = (request: Request) =>
  request.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || "local";

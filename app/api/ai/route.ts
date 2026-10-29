import { AIError, aiModel, aiProvider, askSheet, checkSheet, explainFormula, planActions, writeFormula } from "@/lib/ai/ai.server";
import type { AIRequest, AIResponse } from "@/lib/ai/schemas";

const MAX_BODY_BYTES = 200_000;
const MAX_TEXT = 2_000;

// A simple per-address limit so one client can't spend the whole API budget.
// It lives in memory, so it resets on restart and isn't shared between servers;
// put real authentication and limits in front of this before going public.
const WINDOW_MS = 60_000;
const MAX_PER_WINDOW = 20;
const recent = new Map<string, number[]>();

function rateLimited(key: string): boolean {
  const now = Date.now();
  const hits = (recent.get(key) ?? []).filter((t) => now - t < WINDOW_MS);
  hits.push(now);
  recent.set(key, hits);
  return hits.length > MAX_PER_WINDOW;
}

const reply = (body: AIResponse, status = 200) => Response.json(body, { status });
const fail = (error: string, status: number) => reply({ ok: false, error }, status);

const text = (v: unknown, max = MAX_TEXT) => (typeof v === "string" && v.trim() && v.length <= max ? v.trim() : null);

export async function POST(request: Request) {
  const raw = await request.text();
  if (raw.length > MAX_BODY_BYTES) return fail("That request is too large.", 413);

  let body: AIRequest;
  try {
    body = JSON.parse(raw) as AIRequest;
  } catch {
    return fail("The request wasn't valid JSON.", 400);
  }

  if (body.action === "status") {
    const provider = aiProvider();
    return reply({ ok: true, action: "status", enabled: provider !== null, provider, model: provider ? aiModel() : "" });
  }

  const who = request.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || "local";
  if (rateLimited(who)) return fail("That's a lot of AI requests in a minute. Wait a moment and try again.", 429);

  const context = typeof body.context === "string" ? body.context : "";
  try {
    switch (body.action) {
      case "formula": {
        const prompt = text(body.prompt);
        if (!prompt) return fail("Describe the formula you want, in up to 2,000 characters.", 400);
        return reply({ ok: true, action: "formula", result: await writeFormula(context, prompt) });
      }
      case "explain": {
        const cell = text(body.cell, 20);
        const formula = text(body.formula, 8_000);
        if (!cell || !formula) return fail("Pick a cell with a formula to explain.", 400);
        return reply({ ok: true, action: "explain", result: await explainFormula(context, cell, formula) });
      }
      case "ask": {
        const question = text(body.question);
        if (!question) return fail("Type a question, in up to 2,000 characters.", 400);
        return reply({ ok: true, action: "ask", result: await askSheet(context, question) });
      }
      case "check":
        return reply({ ok: true, action: "check", result: await checkSheet(context) });
      case "plan": {
        const request = text(body.request);
        if (!request) return fail("Say what you'd like done, in up to 2,000 characters.", 400);
        return reply({ ok: true, action: "plan", result: await planActions(context, request) });
      }
      default:
        return fail("Unknown AI action.", 400);
    }
  } catch (e) {
    if (e instanceof AIError) return fail(e.message, e.status);
    console.error("AI route failed", e);
    return fail("Something went wrong with the AI request.", 500);
  }
}

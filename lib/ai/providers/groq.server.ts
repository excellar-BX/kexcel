// Server-only adapter for Groq's OpenAI-compatible chat completions API.
import { z } from "zod";
import { AIError } from "../errors";

const ENDPOINT = "https://api.groq.com/openai/v1/chat/completions";
const DEFAULT_MODEL = "openai/gpt-oss-120b";
const TIMEOUT_MS = 60_000;

export const groqConfigured = () => Boolean(process.env.GROQ_API_KEY);
export const groqModel = () => process.env.GROQ_MODEL?.trim() || DEFAULT_MODEL;

interface ChatCompletion {
  choices?: { finish_reason?: string; message?: { content?: string | null } }[];
  error?: { message?: string; code?: string };
}

/** Pulls a JSON object out of a reply, tolerating a surrounding code fence. */
function extractJSON(text: string): unknown {
  const trimmed = text.trim().replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "");
  return JSON.parse(trimmed);
}

/** Reasoning models (gpt-oss, qwen3) take an effort setting; others ignore it. */
const isReasoningModel = (model: string) => /gpt-oss|qwen3/i.test(model);

async function complete(
  model: string,
  effort: "low" | "medium",
  messages: { role: string; content: string }[],
): Promise<ChatCompletion> {
  let res: Response;
  try {
    res = await fetch(ENDPOINT, {
      method: "POST",
      headers: { authorization: `Bearer ${process.env.GROQ_API_KEY}`, "content-type": "application/json" },
      body: JSON.stringify({
        model,
        messages,
        temperature: 0.2,
        // Reasoning models spend part of this on thinking before the answer.
        max_completion_tokens: 8192,
        ...(isReasoningModel(model) ? { reasoning_effort: effort } : {}),
        // JSON mode: the reply is a single JSON object. The schema itself goes in the prompt.
        response_format: { type: "json_object" },
      }),
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
  } catch (e) {
    if (e instanceof DOMException && e.name === "TimeoutError") {
      throw new AIError("The AI took too long to answer. Please try again.", 504);
    }
    throw new AIError("Couldn't reach Groq. Check your internet connection.", 502);
  }

  const data = (await res.json().catch(() => ({}))) as ChatCompletion;
  if (res.ok) return data;
  const detail = data.error?.message ?? "";
  if (res.status === 401) throw new AIError("Groq rejected the API key. Check GROQ_API_KEY in .env.local.", 401);
  if (res.status === 429) throw new AIError("Groq's rate limit was reached. Wait a moment and try again.", 429);
  if (res.status === 404 || data.error?.code === "model_not_found") {
    throw new AIError(`Groq doesn't have the model "${model}". Set GROQ_MODEL in .env.local to one from console.groq.com.`, 404);
  }
  // Groq answers 400 when JSON mode output didn't come out as valid JSON.
  if (res.status === 400 && data.error?.code === "json_validate_failed") return { choices: [] };
  console.error("Groq API error", res.status, detail);
  throw new AIError(`Groq returned an error (${res.status}). Please try again.`, 502);
}

export async function callGroq<S extends z.ZodType>(
  schema: S,
  effort: "low" | "medium",
  system: string,
  user: string,
): Promise<z.infer<S>> {
  const model = groqModel();
  const shape = JSON.stringify(z.toJSONSchema(schema));
  const messages = [
    {
      role: "system",
      content: `${system}\n\nReply with one JSON object only, no other text, matching this JSON Schema:\n${shape}`,
    },
    { role: "user", content: user },
  ];

  // One retry if the model's JSON doesn't fit the schema, telling it what was wrong.
  for (let attempt = 0; attempt < 2; attempt++) {
    const data = await complete(model, effort, messages);
    const choice = data.choices?.[0];
    if (choice?.finish_reason === "length") throw new AIError("The AI's answer was too long. Please try again.", 502);
    const content = choice?.message?.content ?? "";
    let problem: string;
    try {
      const parsed = schema.safeParse(extractJSON(content));
      if (parsed.success) return parsed.data;
      problem = parsed.error.issues.map((i) => `${i.path.join(".") || "root"}: ${i.message}`).join("; ");
    } catch {
      problem = "the reply was not valid JSON";
    }
    messages.push(
      { role: "assistant", content },
      { role: "user", content: `That reply didn't match the schema (${problem}). Reply again with only the corrected JSON object.` },
    );
  }
  throw new AIError("The AI's answer came back in the wrong shape. Please try again.", 502);
}

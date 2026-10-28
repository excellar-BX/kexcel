// Server-only adapter for Claude through the official Anthropic SDK.
import Anthropic from "@anthropic-ai/sdk";
import { betaZodOutputFormat } from "@anthropic-ai/sdk/helpers/beta/zod";
import type { z } from "zod";
import { AIError } from "../errors";

const DEFAULT_MODEL = "claude-opus-5-5";

/** Models that accept server-side refusal fallbacks in "default" mode on the Claude API. */
const FALLBACK_MODELS = new Set(["claude-opus-5-5", "claude-opus-5", "claude-fable-5-1", "claude-sonnet-5-5"]);

export const anthropicConfigured = () => Boolean(process.env.ANTHROPIC_API_KEY || process.env.ANTHROPIC_AUTH_TOKEN);
export const anthropicModel = () => process.env.AI_MODEL?.trim() || DEFAULT_MODEL;

export async function callAnthropic<S extends z.ZodType>(
  schema: S,
  effort: "low" | "medium",
  system: string,
  user: string,
): Promise<z.infer<S>> {
  const model = anthropicModel();
  // Fallbacks are a Claude API feature; a custom base URL may not accept them.
  const fallback = !process.env.ANTHROPIC_BASE_URL && FALLBACK_MODELS.has(model);
  const client = new Anthropic({ maxRetries: 2, timeout: 120_000 });

  try {
    const response = await client.beta.messages.parse({
      model,
      max_tokens: 16000,
      system,
      messages: [{ role: "user", content: user }],
      output_config: { effort, format: betaZodOutputFormat(schema) },
      ...(fallback ? { betas: ["server-side-fallback-2026-07-01"], fallbacks: "default" as const } : {}),
    });
    if (response.stop_reason === "refusal") {
      throw new AIError("The AI declined this request. Try rephrasing it.", 422);
    }
    if (response.stop_reason === "max_tokens" || !response.parsed_output) {
      throw new AIError("The AI's answer was incomplete. Please try again.", 502);
    }
    return response.parsed_output as z.infer<S>;
  } catch (e) {
    if (e instanceof AIError) throw e;
    if (e instanceof Anthropic.AuthenticationError) {
      throw new AIError("The AI service rejected the API key. Check ANTHROPIC_API_KEY in .env.local.", 401);
    }
    if (e instanceof Anthropic.PermissionDeniedError) {
      throw new AIError("This API key isn't allowed to use that model.", 403);
    }
    if (e instanceof Anthropic.NotFoundError) {
      throw new AIError(`The model "${model}" wasn't found. Check AI_MODEL in .env.local.`, 404);
    }
    if (e instanceof Anthropic.RateLimitError) {
      throw new AIError("The AI is busy right now. Wait a moment and try again.", 429);
    }
    if (e instanceof Anthropic.APIConnectionError) {
      throw new AIError("Couldn't reach the AI service. Check your internet connection.", 502);
    }
    if (e instanceof Anthropic.APIError) {
      console.error("Claude API error", e.status, e.message);
      throw new AIError(`The AI service returned an error (${e.status ?? "unknown"}). Please try again.`, 502);
    }
    throw e;
  }
}

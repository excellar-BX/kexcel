import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FormulaResult } from "../schemas";
import { callGroq } from "./groq.server";

const reply = (content: string, status = 200, extra: object = {}) =>
  new Response(JSON.stringify(status === 200 ? { choices: [{ finish_reason: "stop", message: { content } }] } : extra), {
    status,
  });

describe("callGroq", () => {
  beforeEach(() => vi.stubEnv("GROQ_API_KEY", "test-key"));
  afterEach(() => {
    vi.unstubAllEnvs();
    vi.unstubAllGlobals();
  });

  it("returns schema-valid JSON and sends the key and JSON mode", async () => {
    const fetchMock = vi.fn().mockResolvedValue(reply('{"formula":"=SUM(A1:A3)","explanation":"Adds them.","cell":null}'));
    vi.stubGlobal("fetch", fetchMock);
    const out = await callGroq(FormulaResult, "system", "user");
    expect(out).toEqual({ formula: "=SUM(A1:A3)", explanation: "Adds them.", cell: null });
    const [, init] = fetchMock.mock.calls[0];
    expect(init.headers.authorization).toBe("Bearer test-key");
    const body = JSON.parse(init.body);
    expect(body.response_format).toEqual({ type: "json_object" });
    expect(body.messages[0].content).toContain("JSON Schema");
  });

  it("retries once with the problem when the shape is wrong", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(reply('{"formula":"=SUM(A1:A3)"}'))
      .mockResolvedValueOnce(reply('```json\n{"formula":"=SUM(A1:A3)","explanation":"x","cell":"B4"}\n```'));
    vi.stubGlobal("fetch", fetchMock);
    const out = await callGroq(FormulaResult, "system", "user");
    expect(out.cell).toBe("B4");
    const retry = JSON.parse(fetchMock.mock.calls[1][1].body).messages.at(-1).content;
    expect(retry).toContain("didn't match the schema");
  });

  it("turns HTTP errors into readable messages", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(reply("", 401, { error: { message: "bad key" } })));
    await expect(callGroq(FormulaResult, "s", "u")).rejects.toMatchObject({ status: 401, message: expect.stringContaining("GROQ_API_KEY") });
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(reply("", 404, { error: { code: "model_not_found" } })));
    await expect(callGroq(FormulaResult, "s", "u")).rejects.toMatchObject({ status: 404 });
  });
});

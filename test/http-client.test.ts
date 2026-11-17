import { describe, expect, it } from "vitest";
import {
  ApiError,
  computeDelay,
  HttpClient,
  parseRetryAfter,
  type FetchLike,
} from "../src/http/client.js";

interface Recorded {
  url: string;
  method: string;
  authorization: string | undefined;
}

function harness(responses: Array<() => Response>) {
  const calls: Recorded[] = [];
  const delays: number[] = [];
  let index = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    calls.push({
      url,
      method: init?.method ?? "GET",
      authorization: (init?.headers as Record<string, string> | undefined)?.["Authorization"],
    });
    const factory = responses[Math.min(index, responses.length - 1)];
    index += 1;
    if (!factory) throw new Error("no response queued");
    return factory();
  };
  const client = new HttpClient({
    baseUrl: "https://outlook.office.com/api/v2.0",
    tokenProvider: async () => ({ token: "secret-token-value" }),
    fetchImpl,
    sleep: async (ms) => {
      delays.push(ms);
    },
    random: () => 0.5,
    maxAttempts: 4,
    baseDelayMs: 100,
    maxDelayMs: 1000,
  });
  return { client, calls, delays };
}

const json = (body: unknown, status = 200, headers: Record<string, string> = {}): (() => Response) => {
  return () =>
    new Response(JSON.stringify(body), {
      status,
      headers: { "content-type": "application/json", ...headers },
    });
};

describe("HttpClient auth + parsing", () => {
  it("injects the bearer token and returns parsed JSON", async () => {
    const { client, calls } = harness([json({ Id: "msg-1" })]);
    const me = await client.get<{ Id: string }>("me/messages");
    expect(me.Id).toBe("msg-1");
    expect(calls[0]?.authorization).toBe("Bearer secret-token-value");
    expect(calls[0]?.url).toBe("https://outlook.office.com/api/v2.0/me/messages");
  });

  it("treats 204 as an empty success", async () => {
    const { client } = harness([() => new Response(null, { status: 204 })]);
    await expect(client.delete("me/events/1")).resolves.toBeUndefined();
  });

  it("rejects a non-JSON success body with invalid_response", async () => {
    const { client } = harness([() => new Response("<html>oops</html>", { status: 200 })]);
    await expect(client.get("me")).rejects.toMatchObject({
      name: "ApiError",
      code: "invalid_response",
    });
  });

  it("only owns URLs under its base path", () => {
    const { client } = harness([]);
    expect(client.ownsUrl("https://outlook.office.com/api/v2.0/me/messages?$top=1")).toBe(true);
    expect(client.ownsUrl("https://evil.example.com/steal")).toBe(false);
  });
});

describe("HttpClient retry/backoff", () => {
  it("retries 429 then succeeds, honouring Retry-After seconds", async () => {
    const { client, calls, delays } = harness([
      json({ error: { code: "TooManyRequests", message: "slow down" } }, 429, { "retry-after": "1" }),
      json({ Id: "msg-1" }),
    ]);
    const me = await client.get<{ Id: string }>("me/messages");
    expect(me.Id).toBe("msg-1");
    expect(calls).toHaveLength(2);
    expect(delays).toEqual([1000]);
  });

  it("backs off exponentially on 503", async () => {
    const { client, calls, delays } = harness([
      json({ error: { code: "serviceUnavailable", message: "later" } }, 503),
      json({ error: { code: "serviceUnavailable", message: "later" } }, 503),
      json({ Id: "ok" }),
    ]);
    await expect(client.get("me")).resolves.toEqual({ Id: "ok" });
    expect(calls).toHaveLength(3);
    expect(delays[0]).toBe(100);
    expect(delays[1]).toBe(200);
  });

  it("stops after maxAttempts and surfaces the last error", async () => {
    const { client, calls } = harness([json({ error: { code: "TooManyRequests", message: "no" } }, 429)]);
    await expect(client.get("me")).rejects.toMatchObject({
      code: "TooManyRequests",
      status: 429,
      retryable: true,
    });
    expect(calls).toHaveLength(4);
  });

  it("does not retry ordinary 4xx errors", async () => {
    const { client, calls } = harness([
      json({ error: { code: "ErrorItemNotFound", message: "no such message" } }, 404),
    ]);
    await expect(client.get("me/messages/1")).rejects.toMatchObject({
      code: "ErrorItemNotFound",
      status: 404,
      retryable: false,
    });
    expect(calls).toHaveLength(1);
  });

  it("does not retry 405 (delta not supported)", async () => {
    const { client, calls } = harness([
      json({ error: { code: "ErrorMethodNotAllowed", message: "method not allowed" } }, 405),
    ]);
    await expect(client.get("me/calendar/delta")).rejects.toMatchObject({ status: 405 });
    expect(calls).toHaveLength(1);
  });

  it("retries transport failures", async () => {
    let attempt = 0;
    const delays: number[] = [];
    const client = new HttpClient({
      baseUrl: "https://outlook.office.com/api/v2.0",
      tokenProvider: async () => ({ token: "t" }),
      fetchImpl: async () => {
        attempt += 1;
        if (attempt === 1) throw new TypeError("fetch failed");
        return new Response(JSON.stringify({ ok: true }), { status: 200 });
      },
      sleep: async (ms) => {
        delays.push(ms);
      },
      random: () => 0.5,
      baseDelayMs: 50,
      maxAttempts: 2,
    });
    await expect(client.get("me")).resolves.toEqual({ ok: true });
    expect(delays).toEqual([50]);
  });
});

describe("HttpClient error shape", () => {
  it("parses the Outlook error envelope plus request id", async () => {
    const { client } = harness([
      () =>
        new Response(
          JSON.stringify({
            error: {
              code: "InvalidAuthenticationToken",
              message: "Access token has expired.",
              innerError: { "request-id": "abc-123" },
            },
          }),
          { status: 401, headers: { "content-type": "application/json" } },
        ),
    ]);
    const error = await client.get("me").catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ApiError);
    const apiError = error as ApiError;
    expect(apiError.code).toBe("InvalidAuthenticationToken");
    expect(apiError.message).toBe("Access token has expired.");
    expect(apiError.requestId).toBe("abc-123");
    expect(apiError.status).toBe(401);
    expect(apiError.hint).toContain("auth login");
    expect(apiError.toLine()).toContain("InvalidAuthenticationToken (HTTP 401)");
  });

  it("never leaks the bearer token in error text", async () => {
    const { client } = harness([
      json({ error: { code: "BadRequest", message: "Authorization: Bearer secret-token-value rejected" } }, 400),
    ]);
    const error = (await client.get("me").catch((e: unknown) => e)) as ApiError;
    expect(error.message).not.toContain("secret-token-value");
    expect(error.message).toContain("[redacted]");
  });

  it("redacts a token echoed in a transport failure", async () => {
    const client = new HttpClient({
      baseUrl: "https://outlook.office.com/api/v2.0",
      tokenProvider: async () => ({ token: "secret-token-value" }),
      fetchImpl: async () => {
        throw new Error("Bearer secret-token-value refused");
      },
      maxAttempts: 1,
    });
    const error = (await client.get("me").catch((e: unknown) => e)) as ApiError;
    expect(error.code).toBe("network_error");
    expect(error.message).not.toContain("secret-token-value");
  });
});

describe("backoff helpers", () => {
  it("parses Retry-After seconds and HTTP dates", () => {
    expect(parseRetryAfter("3")).toBe(3000);
    expect(parseRetryAfter(null)).toBeUndefined();
    expect(parseRetryAfter("garbage")).toBeUndefined();
    const now = Date.UTC(2026, 0, 1, 0, 0, 0);
    expect(parseRetryAfter(new Date(now + 5000).toUTCString(), now)).toBe(5000);
  });

  it("caps delays at maxDelayMs and prefers Retry-After", () => {
    expect(computeDelay(10, { baseDelayMs: 100, maxDelayMs: 2000, random: () => 0.5 })).toBe(2000);
    expect(
      computeDelay(0, { baseDelayMs: 100, maxDelayMs: 2000, retryAfterMs: 9000, random: () => 0.5 }),
    ).toBe(2000);
  });
});

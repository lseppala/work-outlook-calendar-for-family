import { describe, expect, it } from "vitest";

import {
  type Env,
  handleRequest,
  type Runtime,
} from "../src/index";

const SOURCE_CALENDAR = [
  "BEGIN:VCALENDAR",
  "VERSION:2.0",
  "BEGIN:VEVENT",
  "UID:meeting",
  "DTSTAMP:20260825T160000Z",
  "DTSTART:20260826T160000Z",
  "DTEND:20260826T170000Z",
  "X-MICROSOFT-CDO-BUSYSTATUS:BUSY",
  "SUMMARY:Private title",
  "END:VEVENT",
  "END:VCALENDAR",
  "",
].join("\r\n");

const ENV: Env = {
  PERSON_NAME: "Taylor",
  OUTLOOK_CALENDAR_URL: "https://outlook.example.com/private/calendar.ics",
  ACCESS_SECRET: "correct-secret",
};

interface TestContext {
  context: Pick<ExecutionContext, "waitUntil">;
  pending: Promise<unknown>[];
}

function makeContext(): TestContext {
  const pending: Promise<unknown>[] = [];
  return {
    context: {
      waitUntil(promise) {
        pending.push(promise);
      },
    },
    pending,
  };
}

function makeCache(): Runtime["cache"] & {
  entries: Map<string, Response>;
  matchedUrls: string[];
  storedUrls: string[];
} {
  const entries = new Map<string, Response>();
  const matchedUrls: string[] = [];
  const storedUrls: string[] = [];

  return {
    entries,
    matchedUrls,
    storedUrls,
    async match(request) {
      const url = requestUrl(request);
      matchedUrls.push(url);
      return entries.get(url)?.clone();
    },
    async put(request, response) {
      const url = requestUrl(request);
      storedUrls.push(url);
      entries.set(url, response.clone());
    },
  };
}

function requestUrl(request: RequestInfo | URL): string {
  if (typeof request === "string") {
    return request;
  }
  if (request instanceof URL) {
    return request.toString();
  }
  return request.url;
}

function makeRuntime(
  cache = makeCache(),
  sourceResponse = new Response(SOURCE_CALENDAR),
): Runtime & { cache: ReturnType<typeof makeCache>; fetchCalls: string[] } {
  const fetchCalls: string[] = [];
  async function fetcher(input: RequestInfo | URL): Promise<Response> {
    fetchCalls.push(requestUrl(input));
    return sourceResponse.clone();
  }

  return {
    cache,
    fetch: fetcher,
    fetchCalls,
  };
}

function authorizedRequest(method = "GET"): Request {
  return new Request(
    "https://calendar.example.com/calendar.ics?key=correct-secret",
    { method },
  );
}

describe("handleRequest", () => {
  it.each([
    ["wrong secret", "/calendar.ics?key=wrong", "GET"],
    ["missing secret", "/calendar.ics", "GET"],
    [
      "duplicate secret",
      "/calendar.ics?key=correct-secret&key=correct-secret",
      "GET",
    ],
    ["wrong path", "/?key=correct-secret", "GET"],
    ["unsupported method", "/calendar.ics?key=correct-secret", "POST"],
  ])("redirects %s without fetching or checking the cache", async (_, path, method) => {
    const runtime = makeRuntime();
    const { context } = makeContext();

    const response = await handleRequest(
      new Request(`https://calendar.example.com${path}`, { method }),
      ENV,
      context,
      runtime,
    );

    expect(response.status).toBe(302);
    expect(response.headers.get("location")).toBe(
      "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
    );
    expect(runtime.fetchCalls).toEqual([]);
    expect(runtime.cache.matchedUrls).toEqual([]);
  });

  it("fetches, transforms, and caches an authorized calendar", async () => {
    const runtime = makeRuntime();
    const { context, pending } = makeContext();

    const response = await handleRequest(
      authorizedRequest(),
      ENV,
      context,
      runtime,
    );
    const body = await response.text();
    await Promise.all(pending);

    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toBe(
      "text/calendar; charset=utf-8",
    );
    expect(response.headers.get("cache-control")).toBe("public, max-age=300");
    expect(body).toContain("SUMMARY:Taylor in meeting");
    expect(body).not.toContain("Private title");
    expect(runtime.fetchCalls).toEqual([ENV.OUTLOOK_CALENDAR_URL]);
    expect(runtime.cache.storedUrls).toEqual([
      "https://calendar.example.com/calendar.ics",
    ]);
    expect(runtime.cache.storedUrls[0]).not.toContain("correct-secret");
    expect(runtime.cache.storedUrls[0]).not.toContain("outlook");
  });

  it("serves a cache hit without fetching Outlook", async () => {
    const cache = makeCache();
    cache.entries.set(
      "https://calendar.example.com/calendar.ics",
      new Response("cached calendar", {
        headers: { "Content-Type": "text/calendar; charset=utf-8" },
      }),
    );
    const runtime = makeRuntime(cache);
    const { context } = makeContext();

    const response = await handleRequest(
      authorizedRequest(),
      ENV,
      context,
      runtime,
    );

    expect(await response.text()).toBe("cached calendar");
    expect(runtime.fetchCalls).toEqual([]);
  });

  it("returns bodyless HEAD responses from both cache hits and misses", async () => {
    const hitCache = makeCache();
    hitCache.entries.set(
      "https://calendar.example.com/calendar.ics",
      new Response("cached calendar", {
        headers: { "Content-Type": "text/calendar; charset=utf-8" },
      }),
    );
    const hitRuntime = makeRuntime(hitCache);
    const hitContext = makeContext();

    const hit = await handleRequest(
      authorizedRequest("HEAD"),
      ENV,
      hitContext.context,
      hitRuntime,
    );

    const missRuntime = makeRuntime();
    const missContext = makeContext();
    const miss = await handleRequest(
      authorizedRequest("HEAD"),
      ENV,
      missContext.context,
      missRuntime,
    );
    await Promise.all(missContext.pending);

    expect(await hit.text()).toBe("");
    expect(await miss.text()).toBe("");
    expect(hit.headers.get("content-type")).toContain("text/calendar");
    expect(miss.headers.get("content-type")).toContain("text/calendar");
    expect(hitRuntime.fetchCalls).toEqual([]);
    expect(missRuntime.fetchCalls).toEqual([ENV.OUTLOOK_CALENDAR_URL]);
    expect(missRuntime.cache.entries.get(
      "https://calendar.example.com/calendar.ics",
    )).toBeDefined();
  });

  it.each([
    ["missing name", { ...ENV, PERSON_NAME: "" }],
    ["missing URL", { ...ENV, OUTLOOK_CALENDAR_URL: "" }],
    ["non-HTTPS URL", { ...ENV, OUTLOOK_CALENDAR_URL: "http://example.com" }],
    ["malformed URL", { ...ENV, OUTLOOK_CALENDAR_URL: "not a url" }],
  ])("returns a sanitized 500 for %s", async (_, env) => {
    const runtime = makeRuntime();
    const { context } = makeContext();

    const response = await handleRequest(
      authorizedRequest(),
      env,
      context,
      runtime,
    );

    expect(response.status).toBe(500);
    expect(response.headers.get("cache-control")).toBe("no-store");
    expect(await response.text()).not.toContain(ENV.OUTLOOK_CALENDAR_URL);
    expect(runtime.fetchCalls).toEqual([]);
  });

  it.each([
    ["upstream error", new Response("private error", { status: 503 })],
    ["malformed calendar", new Response("private malformed content")],
  ])("returns an uncached sanitized 502 for %s", async (_, sourceResponse) => {
    const runtime = makeRuntime(makeCache(), sourceResponse);
    const { context, pending } = makeContext();

    const response = await handleRequest(
      authorizedRequest(),
      ENV,
      context,
      runtime,
    );
    await Promise.all(pending);
    const body = await response.text();

    expect(response.status).toBe(502);
    expect(response.headers.get("cache-control")).toBe("no-store");
    expect(body).not.toContain("private");
    expect(runtime.cache.storedUrls).toEqual([]);
  });

  it("returns a sanitized 502 when the source fetch rejects", async () => {
    const cache = makeCache();
    async function failingFetch(): Promise<Response> {
      throw new Error(`Could not fetch ${ENV.OUTLOOK_CALENDAR_URL}`);
    }
    const runtime: Runtime = { cache, fetch: failingFetch };
    const { context } = makeContext();

    const response = await handleRequest(
      authorizedRequest(),
      ENV,
      context,
      runtime,
    );

    expect(response.status).toBe(502);
    expect(await response.text()).toBe("Calendar source is unavailable");
    expect(cache.storedUrls).toEqual([]);
  });

  it("aborts a source fetch after the timeout", async () => {
    const cache = makeCache();
    let receivedSignal: AbortSignal | undefined;
    function hangingFetch(
      _input: RequestInfo | URL,
      init?: RequestInit,
    ): Promise<Response> {
      receivedSignal = init?.signal ?? undefined;
      return new Promise((_, reject) => {
        receivedSignal?.addEventListener("abort", () => {
          reject(new DOMException("Aborted", "AbortError"));
        });
      });
    }
    const runtime: Runtime = {
      cache,
      fetch: hangingFetch,
      upstreamTimeoutMs: 1,
    };
    const { context } = makeContext();

    const response = await handleRequest(
      authorizedRequest(),
      ENV,
      context,
      runtime,
    );

    expect(receivedSignal?.aborted).toBe(true);
    expect(response.status).toBe(502);
    expect(cache.storedUrls).toEqual([]);
  });
});

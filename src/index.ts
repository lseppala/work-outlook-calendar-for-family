import {
  CalendarTransformError,
  transformCalendar,
} from "./calendar";

const CALENDAR_PATH = "/calendar.ics";
const CACHE_TTL_SECONDS = 300;
const DECOY_URL = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";
const UPSTREAM_TIMEOUT_MS = 10_000;

export interface Env {
  PERSON_NAME: string;
  OUTLOOK_CALENDAR_URL: string;
  ACCESS_SECRET: string;
}

export interface Runtime {
  fetch: typeof fetch;
  cache: Pick<Cache, "match" | "put">;
  upstreamTimeoutMs?: number;
}

export default {
  fetch(
    request: Request,
    env: Env,
    context: ExecutionContext,
  ): Promise<Response> {
    return handleRequest(request, env, context, {
      fetch,
      cache: caches.default,
    });
  },
} satisfies ExportedHandler<Env>;

export async function handleRequest(
  request: Request,
  env: Env,
  context: Pick<ExecutionContext, "waitUntil">,
  runtime: Runtime,
): Promise<Response> {
  const url = new URL(request.url);
  const methodIsAllowed = request.method === "GET" || request.method === "HEAD";
  const keys = url.searchParams.getAll("key");

  if (
    !methodIsAllowed ||
    url.pathname !== CALENDAR_PATH ||
    keys.length !== 1 ||
    !env.ACCESS_SECRET ||
    !(await secretsMatch(keys[0], env.ACCESS_SECRET))
  ) {
    return Response.redirect(DECOY_URL, 302);
  }

  const configurationError = validateConfiguration(env);
  if (configurationError !== undefined) {
    return textResponse(configurationError, 500);
  }

  const cacheRequest = buildCacheRequest(url);
  const cached = await runtime.cache.match(cacheRequest);
  if (cached !== undefined) {
    return request.method === "HEAD" ? withoutBody(cached) : cached;
  }

  const sourceResponse = await fetchSourceCalendar(
    env.OUTLOOK_CALENDAR_URL,
    runtime.fetch,
    runtime.upstreamTimeoutMs ?? UPSTREAM_TIMEOUT_MS,
  );
  if (sourceResponse === undefined || !sourceResponse.ok) {
    return textResponse("Calendar source is unavailable", 502);
  }

  let transformed: string;
  try {
    transformed = transformCalendar(
      await sourceResponse.text(),
      env.PERSON_NAME,
    );
  } catch (error) {
    if (error instanceof CalendarTransformError) {
      return textResponse("Calendar source is invalid", 502);
    }
    throw error;
  }

  const response = new Response(transformed, {
    headers: {
      "Cache-Control": `public, max-age=${CACHE_TTL_SECONDS}`,
      "Content-Type": "text/calendar; charset=utf-8",
      "X-Content-Type-Options": "nosniff",
    },
  });

  context.waitUntil(runtime.cache.put(cacheRequest, response.clone()));
  return request.method === "HEAD" ? withoutBody(response) : response;
}

async function secretsMatch(provided: string, expected: string): Promise<boolean> {
  const encoder = new TextEncoder();
  const [providedHash, expectedHash] = await Promise.all([
    crypto.subtle.digest("SHA-256", encoder.encode(provided)),
    crypto.subtle.digest("SHA-256", encoder.encode(expected)),
  ]);
  const providedBytes = new Uint8Array(providedHash);
  const expectedBytes = new Uint8Array(expectedHash);
  let difference = 0;

  for (let index = 0; index < providedBytes.length; index += 1) {
    difference |= providedBytes[index] ^ expectedBytes[index];
  }

  return difference === 0;
}

function validateConfiguration(env: Env): string | undefined {
  if (!env.PERSON_NAME?.trim() || !env.OUTLOOK_CALENDAR_URL?.trim()) {
    return "Worker configuration is incomplete";
  }

  try {
    const outlookUrl = new URL(env.OUTLOOK_CALENDAR_URL);
    if (outlookUrl.protocol !== "https:") {
      return "Worker configuration is invalid";
    }
  } catch {
    return "Worker configuration is invalid";
  }

  return undefined;
}

function buildCacheRequest(requestUrl: URL): Request {
  const cacheUrl = new URL(requestUrl);
  cacheUrl.search = "";
  cacheUrl.hash = "";
  return new Request(cacheUrl, { method: "GET" });
}

async function fetchSourceCalendar(
  url: string,
  fetcher: typeof fetch,
  timeoutMs: number,
): Promise<Response | undefined> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetcher(url, {
      headers: {
        Accept: "text/calendar",
      },
      signal: controller.signal,
    });
  } catch {
    return undefined;
  } finally {
    clearTimeout(timeout);
  }
}

function withoutBody(response: Response): Response {
  return new Response(null, {
    headers: response.headers,
    status: response.status,
    statusText: response.statusText,
  });
}

function textResponse(message: string, status: number): Response {
  return new Response(message, {
    status,
    headers: {
      "Cache-Control": "no-store",
      "Content-Type": "text/plain; charset=utf-8",
      "X-Content-Type-Options": "nosniff",
    },
  });
}

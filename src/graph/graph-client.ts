import type { TokenProvider } from '../auth/token-provider.js';

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
    /** Seconds the server asked us to wait (Retry-After), when it named one. */
    readonly retryAfterSeconds?: number,
  ) {
    super(message);
    this.name = 'GraphError';
  }

  /**
   * The known blocker: a licence-less account gets 403 "Failed to get license information" from
   * every /chats endpoint. Worth calling out by name so nobody spends an afternoon on the token.
   */
  get isLicenceProblem(): boolean {
    return this.status === 403 && /license/i.test(this.message);
  }
}

export interface GraphClientOptions {
  tokenProvider: TokenProvider;
  fetchFn?: typeof fetch;
  baseUrl?: string;
  sleepFn?: (ms: number) => Promise<void>;
  /** Extra attempts after the first for throttled/unavailable READS. Writes never auto-retry. Default 1. */
  readRetries?: number;
  nowFn?: () => number;
}

/** When a 429 names no Retry-After, close the gate for this long rather than for nothing. */
export const DEFAULT_THROTTLE_WINDOW_MS = 30_000;
/**
 * The gate honours the FULL Retry-After (sanity-capped at an hour): resuming inside the window
 * Graph named is what escalates a throttle. What IS capped is how long one call will sleep
 * waiting for a retry — a tool call must not hang for minutes; beyond this the call fails fast
 * with the 429 and the gate keeps everyone else honest until the window really passes.
 */
const MAX_THROTTLE_WINDOW_MS = 60 * 60_000;
export const MAX_RETRY_SLEEP_MS = 60_000;

const RETRYABLE_READ_STATUSES = new Set([429, 503, 504]);

/** One reading of Retry-After for the whole client: a positive finite number of seconds, else nothing. */
function retryAfterSecondsOf(response: Response): number | undefined {
  const parsed = Number(response.headers.get('retry-after'));
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

export class GraphClient {
  private readonly fetchFn: typeof fetch;
  private readonly baseUrl: string;
  private readonly sleepFn: (ms: number) => Promise<void>;
  private readonly readRetries: number;
  private readonly nowFn: () => number;
  /**
   * The throttle gate. Graph ESCALATES its penalty window when a caller keeps sending while
   * throttled — and on 2026-08-25 this client's own retries, stacked under a poller and a
   * readback loop, kept an account throttled for hours. One 429 now closes the gate for the
   * whole client: every request until the window passes fails fast, locally, without touching
   * the network. Silence is the only thing that ends a throttle.
   */
  private throttledUntil = 0;

  constructor(private readonly options: GraphClientOptions) {
    this.fetchFn = options.fetchFn ?? fetch;
    this.baseUrl = options.baseUrl ?? GRAPH_BASE_URL;
    this.sleepFn = options.sleepFn ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.readRetries = options.readRetries ?? 1;
    this.nowFn = options.nowFn ?? (() => Date.now());
  }

  /** Milliseconds until the gate reopens; 0 when open. Callers pacing themselves read this. */
  throttledForMs(): number {
    return Math.max(0, this.throttledUntil - this.nowFn());
  }

  private assertGateOpen(path: string): void {
    const remaining = this.throttledForMs();
    if (remaining > 0) {
      throw new GraphError(
        `Locally throttled: a recent 429 closed this client for another ${Math.ceil(remaining / 1000)}s (${path} not sent)`,
        429,
        'LocallyThrottled',
        Math.ceil(remaining / 1000),
      );
    }
  }

  private noteThrottle(response: Response): void {
    if (response.status !== 429) {
      return;
    }
    const named = retryAfterSecondsOf(response);
    const windowMs = named ? Math.min(named * 1000, MAX_THROTTLE_WINDOW_MS) : DEFAULT_THROTTLE_WINDOW_MS;
    this.throttledUntil = Math.max(this.throttledUntil, this.nowFn() + windowMs);
  }

  private url(path: string): string {
    return path.startsWith('http') ? path : `${this.baseUrl}${path}`;
  }

  private async authorized(path: string, init: RequestInit): Promise<Response> {
    this.assertGateOpen(path);
    const token = await this.options.tokenProvider.getAccessToken();
    const response = await this.fetchFn(this.url(path), {
      ...init,
      headers: {
        ...(init.headers as Record<string, string> | undefined),
        authorization: `Bearer ${token}`,
      },
    });
    this.noteThrottle(response);
    return response;
  }

  private async fail(response: Response): Promise<never> {
    const body = (await response.json().catch(() => undefined)) as
      | { error?: { code?: string; message?: string } }
      | undefined;
    throw new GraphError(
      body?.error?.message ?? `Graph request failed with HTTP ${response.status}`,
      response.status,
      body?.error?.code,
      retryAfterSecondsOf(response),
    );
  }

  async get<T>(path: string): Promise<T> {
    // Reads are idempotent, so a throttled or briefly unavailable GET may be retried after the
    // wait the server names (capped — an aggressive Retry-After must not park the caller for
    // minutes), falling back to a short exponential pause.
    for (let attempt = 0; ; attempt += 1) {
      const response = await this.authorized(path, { method: 'GET' });
      if (response.ok) {
        if (response.status === 204) {
          return undefined as T;
        }
        return (await response.json()) as T;
      }
      if (attempt >= this.readRetries || !RETRYABLE_READ_STATUSES.has(response.status)) {
        await this.fail(response);
      }
      await response.body?.cancel().catch(() => undefined);
      // The wait must outlast the gate the 429 just closed, or the retry fails fast locally —
      // and if that wait would exceed what one call may hang for, there is no honest retry:
      // fail now with the 429 rather than sleep a minute and then fail locally anyway.
      const waitMs = Math.max(this.throttledForMs(), 2 ** attempt * 1000);
      if (waitMs > MAX_RETRY_SLEEP_MS) {
        await this.fail(response);
      }
      await this.sleepFn(waitMs);
    }
  }

  async post<T>(path: string, body: unknown): Promise<T> {
    // Deliberately NO automatic retry here: a POST is a write, and a failure report says nothing
    // about whether the write landed. Retry policy for sends lives in ReliableTeamsChats, which
    // reads the chat back before ever sending again.
    const response = await this.authorized(path, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      await this.fail(response);
    }
    // Creates MUST return the created resource — an empty success here is an unknown outcome,
    // and defaulting it to an empty object would hand callers a message with no id. Actions
    // that legitimately answer 204 use postNoContent instead.
    const raw = await response.text();
    if (response.status === 204 || raw.trim() === '') {
      throw new GraphError(
        `Graph answered ${response.status} with an empty body to a JSON POST on ${path} — ` +
          'the write may have landed; treat the outcome as unknown, not as failed.',
        response.status,
        'EmptyCreateResponse',
      );
    }
    return JSON.parse(raw) as T;
  }

  /**
   * POST for OData actions that take a JSON body and answer 204 No Content (setReaction and
   * friends). Parsing that success as JSON used to throw, and that lie — a landed write
   * reported as failed — is one half of how a broadcast became eleven on 2026-08-24.
   */
  async postNoContent(path: string, body: unknown): Promise<void> {
    const response = await this.authorized(path, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      await this.fail(response);
    }
    await response.body?.cancel().catch(() => undefined);
  }

  /** PATCH with a JSON body. Graph answers 204 No Content on success, so nothing is returned. */
  async patch(path: string, body: unknown): Promise<void> {
    const response = await this.authorized(path, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      await this.fail(response);
    }
  }

  /** POST to an OData action that takes no body and answers 204 No Content. */
  async postAction(path: string): Promise<void> {
    const response = await this.authorized(path, { method: 'POST' });
    if (!response.ok) {
      await this.fail(response);
    }
  }

  async putBinary<T>(path: string, bytes: Uint8Array, contentType: string): Promise<T> {
    const response = await this.authorized(path, {
      method: 'PUT',
      headers: { 'content-type': contentType },
      body: bytes,
    });
    if (!response.ok) {
      await this.fail(response);
    }
    return (await response.json()) as T;
  }

  async getBinary(path: string): Promise<{ bytes: Uint8Array; contentType: string }> {
    const response = await this.authorized(path, { method: 'GET' });
    if (!response.ok) {
      await this.fail(response);
    }
    return {
      bytes: new Uint8Array(await response.arrayBuffer()),
      contentType: response.headers.get('content-type') ?? 'application/octet-stream',
    };
  }

  /** Follows @odata.nextLink until the pages run out or `max` items have been collected. */
  async getAll<T>(path: string, max = 200): Promise<T[]> {
    const collected: T[] = [];
    let next: string | undefined = path;
    while (next && collected.length < max) {
      const page: { value?: T[]; '@odata.nextLink'?: string } = await this.get(next);
      collected.push(...(page.value ?? []));
      next = page['@odata.nextLink'];
    }
    return collected.slice(0, max);
  }
}

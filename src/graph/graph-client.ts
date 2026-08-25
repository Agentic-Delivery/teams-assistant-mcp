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
/** Graph collections whose next path segment is an id — the shape of a resource family. Only the
 *  collections this chats-only server actually touches; a speculative list only hides gaps. */
const GATE_COLLECTIONS = new Set(['chats', 'messages', 'shares', 'hostedContents']);
/** Error codes Graph uses for APPLICATION-wide throttling: these close every family at once. */
const GLOBAL_THROTTLE_CODES = new Set(['ApplicationThrottled', 'ActivityLimitReached', 'ServiceUnavailable']);
const GLOBAL_GATE_KEY = '*';

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
   * The throttle gates. Graph ESCALATES its penalty window when a caller keeps sending while
   * throttled — and on 2026-08-25 this client's own retries, stacked under a poller and a
   * readback loop, kept an account throttled for hours. A 429 closes a gate: every request on
   * that gate until the window passes fails fast, locally, without touching the network.
   * Silence is the only thing that ends a throttle.
   *
   * Gates are keyed per resource FAMILY (see gateKeyFor) because Graph throttles per resource —
   * the same day showed the single-message GET refused for hours while the message list and
   * plain posts on the same chat were healthy. An APPLICATION-wide throttle (Graph says so by
   * error code, e.g. ApplicationThrottled) closes the global gate, which every request checks.
   */
  private readonly throttledUntil = new Map<string, number>();

  constructor(private readonly options: GraphClientOptions) {
    this.fetchFn = options.fetchFn ?? fetch;
    this.baseUrl = options.baseUrl ?? GRAPH_BASE_URL;
    this.sleepFn = options.sleepFn ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.readRetries = options.readRetries ?? 1;
    this.nowFn = options.nowFn ?? (() => Date.now());
  }

  /**
   * Graph throttles per resource family, and 2026-08-25 showed the families really are
   * separate: the single-message GET (`/chats/{id}/messages/{id}`) stayed 429 for hours while
   * the message LIST and plain posts on the very same chat answered 200/201. One gate for the
   * whole client would have kept the healthy families locked out, so the gate is keyed by the
   * path with its ids blanked — `/chats/{id}/messages/{id}` and `/chats/{id}/messages` are two gates.
   */
  static gateKeyFor(path: string): string {
    const withoutHost = path.replace(/^https?:\/\/[^/]+/, '').replace(/^\/v1\.0/, '');
    const pathname = withoutHost.split('?')[0] as string;
    // OneDrive path syntax (`/me/drive/root:/folder/file.pdf:/content`) carries the file name in
    // the path itself; everything from `root:` on is one family, whatever the file or folder.
    const driveCut = pathname.indexOf('/root:');
    const shaped = driveCut >= 0 ? `${pathname.slice(0, driveCut)}/root:*` : pathname;
    const segments = shaped.split('/');
    // Structural, not length-based: the segment AFTER a collection name is an id.
    return segments
      .map((segment, index) => (index > 0 && GATE_COLLECTIONS.has(segments[index - 1] ?? '') ? '*' : segment))
      .join('/');
  }

  /** Milliseconds until this path's gate (its family's or the global one) reopens; 0 when open. */
  throttledForMs(path: string): number {
    const now = this.nowFn();
    const family = this.throttledUntil.get(GraphClient.gateKeyFor(path)) ?? 0;
    const global = this.throttledUntil.get(GLOBAL_GATE_KEY) ?? 0;
    return Math.max(0, Math.max(family, global) - now);
  }

  private assertGateOpen(path: string): void {
    const remaining = this.throttledForMs(path);
    if (remaining > 0) {
      throw new GraphError(
        `Locally throttled: a recent 429 closed this client for another ${Math.ceil(remaining / 1000)}s (${path} not sent)`,
        429,
        'LocallyThrottled',
        Math.ceil(remaining / 1000),
      );
    }
  }

  private async noteThrottle(path: string, response: Response): Promise<void> {
    if (response.status !== 429) {
      return;
    }
    const named = retryAfterSecondsOf(response);
    const windowMs = named ? Math.min(named * 1000, MAX_THROTTLE_WINDOW_MS) : DEFAULT_THROTTLE_WINDOW_MS;
    // Peek at the error code without consuming the body the caller still needs.
    const code = await response
      .clone()
      .json()
      .then((body) => (body as { error?: { code?: string } })?.error?.code)
      .catch(() => undefined);
    const key = code && GLOBAL_THROTTLE_CODES.has(code) ? GLOBAL_GATE_KEY : GraphClient.gateKeyFor(path);
    this.throttledUntil.set(key, Math.max(this.throttledUntil.get(key) ?? 0, this.nowFn() + windowMs));
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
    await this.noteThrottle(path, response);
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

  async get<T>(path: string, options: { readRetries?: number } = {}): Promise<T> {
    // Reads are idempotent, so a throttled or briefly unavailable GET may be retried after the
    // wait the server names (capped — an aggressive Retry-After must not park the caller for
    // minutes), falling back to a short exponential pause. A caller holding a cheaper fallback
    // (a list scan instead of a throttled single fetch) passes readRetries: 0 and takes it now.
    const readRetries = options.readRetries ?? this.readRetries;
    for (let attempt = 0; ; attempt += 1) {
      const response = await this.authorized(path, { method: 'GET' });
      if (response.ok) {
        if (response.status === 204) {
          return undefined as T;
        }
        return (await response.json()) as T;
      }
      if (attempt >= readRetries || !RETRYABLE_READ_STATUSES.has(response.status)) {
        await this.fail(response);
      }
      // The wait honours a named Retry-After on ANY retryable status (a 503 that asks for
      // room gets it), must outlast the gate a 429 just closed (or the retry fails fast
      // locally), and never shrinks below a short exponential pause. If that wait would exceed
      // what one call may hang for, there is no honest retry: fail now — with Graph's own
      // message and code intact — rather than sleep a minute and then fail locally anyway.
      const named = retryAfterSecondsOf(response);
      const waitMs = Math.max(named ? named * 1000 : 0, this.throttledForMs(path), 2 ** attempt * 1000);
      if (waitMs > MAX_RETRY_SLEEP_MS) {
        await this.fail(response);
      }
      await response.body?.cancel().catch(() => undefined);
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

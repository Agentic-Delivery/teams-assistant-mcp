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
  /** Extra attempts after the first for throttled/unavailable READS. Writes never auto-retry. */
  readRetries?: number;
}

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

  constructor(private readonly options: GraphClientOptions) {
    this.fetchFn = options.fetchFn ?? fetch;
    this.baseUrl = options.baseUrl ?? GRAPH_BASE_URL;
    this.sleepFn = options.sleepFn ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.readRetries = options.readRetries ?? 2;
  }

  private url(path: string): string {
    return path.startsWith('http') ? path : `${this.baseUrl}${path}`;
  }

  private async authorized(path: string, init: RequestInit): Promise<Response> {
    const token = await this.options.tokenProvider.getAccessToken();
    return this.fetchFn(this.url(path), {
      ...init,
      headers: {
        ...(init.headers as Record<string, string> | undefined),
        authorization: `Bearer ${token}`,
      },
    });
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
      const retryAfter = retryAfterSecondsOf(response);
      const waitMs = retryAfter ? Math.min(retryAfter, 60) * 1000 : 2 ** attempt * 1000;
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

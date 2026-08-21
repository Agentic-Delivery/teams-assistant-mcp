import type { TokenProvider } from '../auth/token-provider.js';

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
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
}

export class GraphClient {
  private readonly fetchFn: typeof fetch;
  private readonly baseUrl: string;

  constructor(private readonly options: GraphClientOptions) {
    this.fetchFn = options.fetchFn ?? fetch;
    this.baseUrl = options.baseUrl ?? GRAPH_BASE_URL;
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
    );
  }

  async get<T>(path: string): Promise<T> {
    const response = await this.authorized(path, { method: 'GET' });
    if (!response.ok) {
      await this.fail(response);
    }
    return (await response.json()) as T;
  }

  async post<T>(path: string, body: unknown): Promise<T> {
    const response = await this.authorized(path, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!response.ok) {
      await this.fail(response);
    }
    return (await response.json()) as T;
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

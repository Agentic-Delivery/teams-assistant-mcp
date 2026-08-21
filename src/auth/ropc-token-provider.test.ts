import { describe, expect, it, vi } from 'vitest';
import { RopcTokenProvider } from './ropc-token-provider.js';
import { MemoryTokenCache } from './token-cache.js';
import { AuthenticationError, type TokenProvider } from './token-provider.js';

function tokenResponse(body: Record<string, unknown>, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

function provider(fetchFn: typeof fetch, now = () => 1_000_000, cache = new MemoryTokenCache()) {
  return new RopcTokenProvider({
    tenantId: 'tenant',
    clientId: 'client',
    username: 'assistant@example.com',
    password: 'pw',
    fetchFn,
    now,
    cache,
  });
}

describe('ROPC token provider', () => {
  it('exchanges username and password for an access token', async () => {
    const fetchFn = vi.fn(async () => tokenResponse({ access_token: 'tok', expires_in: 3600 }));

    expect(await provider(fetchFn as unknown as typeof fetch).getAccessToken()).toBe('tok');

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('https://login.microsoftonline.com/tenant/oauth2/v2.0/token');
    const body = new URLSearchParams(init.body as string);
    expect(body.get('grant_type')).toBe('password');
    expect(body.get('scope')).toContain('https://graph.microsoft.com/.default');
    expect(body.get('scope')).toContain('offline_access');
  });

  it('reuses a cached token instead of signing in again', async () => {
    const fetchFn = vi.fn(async () => tokenResponse({ access_token: 'tok', expires_in: 3600 }));
    const subject = provider(fetchFn as unknown as typeof fetch);

    await subject.getAccessToken();
    await subject.getAccessToken();

    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it('refreshes before the token actually expires', async () => {
    let clock = 0;
    const fetchFn = vi.fn(async () => tokenResponse({ access_token: `tok-${clock}`, expires_in: 60 }));
    const subject = provider(fetchFn as unknown as typeof fetch, () => clock);

    expect(await subject.getAccessToken()).toBe('tok-0');
    // 30s in: still 30s of life left, but inside the 60s skew, so it must not be handed out.
    clock = 30_000;
    expect(await subject.getAccessToken()).toBe('tok-30000');
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it('uses the refresh token when it has one, and falls back to the password if that fails', async () => {
    const grants: string[] = [];
    const fetchFn = vi.fn(async (_url: string, init: RequestInit) => {
      const grant = new URLSearchParams(init.body as string).get('grant_type') ?? '';
      grants.push(grant);
      if (grant === 'password') {
        return tokenResponse({ access_token: 'tok', expires_in: 60, refresh_token: 'rt' });
      }
      return tokenResponse({ error: 'invalid_grant', error_description: 'AADSTS700082 expired' }, 400);
    });

    let clock = 0;
    const subject = provider(fetchFn as unknown as typeof fetch, () => clock);
    await subject.getAccessToken();
    clock = 3_600_000;
    await subject.getAccessToken();

    expect(grants).toEqual(['password', 'refresh_token', 'password']);
  });

  it('does not fire two password grants for two concurrent callers', async () => {
    const fetchFn = vi.fn(async () => tokenResponse({ access_token: 'tok', expires_in: 3600 }));
    const subject = provider(fetchFn as unknown as typeof fetch);

    await Promise.all([subject.getAccessToken(), subject.getAccessToken()]);

    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it('surfaces the AADSTS code so a blocked sign-in is diagnosable', async () => {
    const fetchFn = vi.fn(async () =>
      tokenResponse(
        {
          error: 'invalid_grant',
          error_description: 'AADSTS50076: multi-factor authentication is required',
        },
        400,
      ),
    );

    await expect(provider(fetchFn as unknown as typeof fetch).getAccessToken()).rejects.toThrow(
      AuthenticationError,
    );
    await expect(
      provider(fetchFn as unknown as typeof fetch).getAccessToken(),
    ).rejects.toThrow(/AADSTS50076/);
  });

  it('never puts the password in the error text', async () => {
    const fetchFn = vi.fn(async () =>
      tokenResponse({ error: 'invalid_grant', error_description: 'AADSTS50126: bad password' }, 400),
    );

    await expect(provider(fetchFn as unknown as typeof fetch).getAccessToken()).rejects.not.toThrow(
      /pw/,
    );
  });

  it('is substitutable — anything satisfying TokenProvider works in its place', async () => {
    const deviceCodeLookalike: TokenProvider = {
      kind: 'device-code',
      getAccessToken: async () => 'from-device-code',
    };

    expect(await deviceCodeLookalike.getAccessToken()).toBe('from-device-code');
  });
});

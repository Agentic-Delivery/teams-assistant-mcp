import { AuthenticationError, EXPIRY_SKEW_MS, type TokenProvider } from './token-provider.js';
import { MemoryTokenCache, type TokenCache } from './token-cache.js';

export interface RopcOptions {
  tenantId: string;
  clientId: string;
  username: string;
  password: string;
  scope?: string;
  cache?: TokenCache;
  fetchFn?: typeof fetch;
  now?: () => number;
  authorityHost?: string;
}

interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
}

const DEFAULT_SCOPE = 'https://graph.microsoft.com/.default offline_access';
const DEFAULT_AUTHORITY = 'https://login.microsoftonline.com';

/**
 * Resource Owner Password Credentials against Entra ID.
 *
 * This works today with a first-party Teams client id and needs no app registration and no admin
 * consent — that is the whole reason it was chosen. It also cannot be used with an account that
 * has MFA or an MFA-requiring Conditional Access policy, and Microsoft has said the grant will go
 * away. When it does, write another TokenProvider (device code is the obvious successor: same
 * client id, same scopes, one interactive sign-in per refresh-token lifetime) and swap it in
 * src/index.ts. Nothing else in this server needs to change.
 */
export class RopcTokenProvider implements TokenProvider {
  readonly kind = 'ropc';

  private readonly cache: TokenCache;
  private readonly fetchFn: typeof fetch;
  private readonly now: () => number;
  private readonly scope: string;
  private readonly tokenEndpoint: string;
  private inFlight: Promise<string> | undefined;
  private forceReauth = false;

  constructor(private readonly options: RopcOptions) {
    this.cache = options.cache ?? new MemoryTokenCache();
    this.fetchFn = options.fetchFn ?? fetch;
    this.now = options.now ?? Date.now;
    this.scope = options.scope ?? DEFAULT_SCOPE;
    const authority = options.authorityHost ?? DEFAULT_AUTHORITY;
    this.tokenEndpoint = `${authority}/${options.tenantId}/oauth2/v2.0/token`;
  }

  async getAccessToken(): Promise<string> {
    const cached = this.forceReauth ? undefined : this.cache.read();
    if (cached && cached.expiresAt - EXPIRY_SKEW_MS > this.now()) {
      return cached.accessToken;
    }
    // Concurrent tool calls must not each fire their own password grant.
    this.inFlight ??= this.acquire(cached?.refreshToken).finally(() => {
      this.inFlight = undefined;
      this.forceReauth = false;
    });
    return this.inFlight;
  }

  /** See TokenProvider.invalidate's doc comment. Dropping the cached token here (rather than just
   *  the refresh token) means the forced re-auth is a full password grant — "re-run ROPC" — not
   *  a refresh_token exchange against a possibly equally-dead refresh token. */
  invalidate(): void {
    this.forceReauth = true;
  }

  private async acquire(refreshToken: string | undefined): Promise<string> {
    if (refreshToken) {
      const refreshed = await this.request({
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
      }).catch(() => undefined);
      if (refreshed) {
        return refreshed;
      }
    }
    return this.request({
      grant_type: 'password',
      username: this.options.username,
      password: this.options.password,
    });
  }

  private async request(grant: Record<string, string>): Promise<string> {
    const body = new URLSearchParams({
      client_id: this.options.clientId,
      scope: this.scope,
      ...grant,
    });

    const response = await this.fetchFn(this.tokenEndpoint, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });

    const payload = (await response.json().catch(() => ({}))) as TokenResponse;

    if (!response.ok || !payload.access_token) {
      // error_description carries the AADSTS code, which is the only useful diagnostic here.
      // It never contains the password, but the request body does, so the body is never logged.
      throw new AuthenticationError(
        payload.error_description ?? `Token request failed with HTTP ${response.status}`,
        payload.error,
      );
    }

    const expiresIn = payload.expires_in ?? 3600;
    this.cache.write({
      accessToken: payload.access_token,
      expiresAt: this.now() + expiresIn * 1000,
      ...(payload.refresh_token ? { refreshToken: payload.refresh_token } : {}),
    });
    return payload.access_token;
  }
}

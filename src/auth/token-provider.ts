/**
 * The only thing the rest of the server knows about authentication.
 *
 * ROPC (username + password password-grant) is on Microsoft's deprecation path and does not
 * survive MFA or most Conditional Access policies. Everything above this interface — the Graph
 * client, the tools, the allowlist — must stay ignorant of how the token was obtained, so that
 * replacing RopcTokenProvider with a device-code or auth-code provider is a wiring change in
 * src/index.ts and nothing else.
 */
export interface TokenProvider {
  /** Short name for logs and the auth probe. Never include the token or the password. */
  readonly kind: string;

  /**
   * A bearer token for Microsoft Graph. Implementations are expected to cache and only hit the
   * identity provider when the cached token is expired or missing.
   */
  getAccessToken(): Promise<string>;
}

export class AuthenticationError extends Error {
  constructor(
    message: string,
    readonly errorCode?: string,
  ) {
    super(message);
    this.name = 'AuthenticationError';
  }
}

/** Treat a token as stale this many milliseconds before its real expiry. */
export const EXPIRY_SKEW_MS = 60_000;

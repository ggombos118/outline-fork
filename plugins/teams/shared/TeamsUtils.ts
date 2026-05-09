import env from "@shared/env";
import type { IntegrationType } from "@shared/types";
import { integrationSettingsPath } from "@shared/utils/routeHelpers";

export const TeamsOAuthNonceCookie = "teamsOAuthNonce";

export type OAuthState = {
  teamId: string;
  type: IntegrationType;
  nonce: string;
};

export class TeamsUtils {
  private static get tenantId() {
    return env.MICROSOFT_TENANT_ID ?? "common";
  }

  private static get msOAuthBaseUrl() {
    return `https://login.microsoftonline.com/${TeamsUtils.tenantId}/oauth2/v2.0/authorize`;
  }

  /**
   * Create a state string for use in OAuth flows.
   *
   * @param teamId - the Outline team ID.
   * @param type - the integration type.
   * @param data - additional data to include in the state.
   * @returns a JSON state string.
   */
  static createState(
    teamId: string,
    type: IntegrationType,
    data?: Record<string, unknown>
  ) {
    return JSON.stringify({ type, teamId, ...data });
  }

  /**
   * Parse a state string from an OAuth flow.
   *
   * @param state - the state string.
   * @returns the parsed OAuthState, or undefined if parsing fails.
   */
  static parseState(state: string): OAuthState | undefined {
    try {
      return JSON.parse(state);
    } catch {
      return undefined;
    }
  }

  /**
   * @returns the URL to the Teams integration settings page.
   */
  static get url() {
    return integrationSettingsPath("teams");
  }

  /**
   * @param err - the error to include in the URL.
   * @returns the URL to redirect to on error.
   */
  static errorUrl(err: string) {
    return integrationSettingsPath(`teams?error=${encodeURIComponent(err)}`);
  }

  /**
   * @returns the OAuth callback URL for Teams LinkedAccount.
   */
  static callbackUrl(
    { baseUrl }: { baseUrl?: string } = { baseUrl: env.URL }
  ) {
    return `${baseUrl ?? env.URL}/auth/teams.callback`;
  }

  /**
   * @returns the URL for the OAuth connect route.
   */
  static connectUrl(
    {
      baseUrl,
      params,
    }: { baseUrl?: string; params?: string } = {
      baseUrl: env.URL,
    }
  ) {
    const base = `${baseUrl ?? env.URL}/auth/teams.connect`;
    return params ? `${base}?${params}` : base;
  }

  /**
   * Build the Microsoft OAuth authorization URL.
   *
   * @param state - the state parameter for CSRF protection.
   * @param redirectUri - the redirect URI registered in Azure AD.
   * @returns the full Microsoft OAuth authorization URL.
   */
  static authUrl(state: string, redirectUri = TeamsUtils.callbackUrl()): string {
    const params = new URLSearchParams({
      client_id: env.MICROSOFT_TEAMS_CLIENT_ID ?? "",
      response_type: "code",
      redirect_uri: redirectUri,
      scope: "openid profile email",
      state,
      response_mode: "query",
      prompt: "select_account",
    });
    return `${TeamsUtils.msOAuthBaseUrl}?${params.toString()}`;
  }
}

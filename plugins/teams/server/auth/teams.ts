import jwt from "jsonwebtoken";
import Router from "koa-router";
import { IntegrationService, IntegrationType } from "@shared/types";
import { ValidationError } from "@server/errors";
import auth from "@server/middlewares/authentication";
import { Integration } from "@server/models";
import { authorize } from "@server/policies";
import type { APIContext } from "@server/types";
import { verifyOAuthStateNonce } from "@server/utils/oauth";
import env from "../env";
import { TeamsUtils, TeamsOAuthNonceCookie } from "../../shared/TeamsUtils";

const router = new Router();

if (env.MICROSOFT_TEAMS_CLIENT_ID && env.MICROSOFT_TEAMS_CLIENT_SECRET) {
  /**
   * Handles the OAuth callback from Microsoft.
   * Exchanges the code for tokens, extracts user identity, and creates
   * a LinkedAccount integration.
   */
  router.get(
    "teams.callback",
    auth({ optional: true }),
    async (ctx: APIContext) => {
      const { code, error, state, error_description: errorDescription } = ctx.query as Record<string, string>;
      const { user } = ctx.state.auth;

      if (error) {
        ctx.redirect(TeamsUtils.errorUrl(error));
        return;
      }

      if (!user) {
        ctx.redirect(TeamsUtils.errorUrl("unauthenticated"));
        return;
      }

      if (!code) {
        ctx.redirect(TeamsUtils.errorUrl("missing_code"));
        return;
      }

      const parsedState = TeamsUtils.parseState(state ?? "");
      if (!parsedState) {
        throw ValidationError("Invalid state");
      }

      verifyOAuthStateNonce(ctx, TeamsOAuthNonceCookie, parsedState.nonce);

      authorize(user, "read", user);

      const tokenResponse = await exchangeCode(
        code,
        TeamsUtils.callbackUrl()
      );

      if (tokenResponse.error) {
        const errorMsg =
          typeof tokenResponse.error === "string"
            ? tokenResponse.error
            : "token_error";
        ctx.redirect(TeamsUtils.errorUrl(errorMsg));
        return;
      }

      const profile = jwt.decode(
        String(tokenResponse.id_token)
      ) as jwt.JwtPayload;

      if (!profile) {
        ctx.redirect(TeamsUtils.errorUrl("invalid_token"));
        return;
      }

      const serviceTeamId: string = String(profile.tid ?? "");
      const serviceUserId: string = String(profile.oid ?? "");

      await Integration.create<Integration<IntegrationType.LinkedAccount>>({
        service: IntegrationService.MicrosoftTeams,
        type: IntegrationType.LinkedAccount,
        userId: user.id,
        teamId: user.teamId,
        settings: {
          teams: {
            serviceTeamId,
            serviceUserId,
          },
        },
      });

      ctx.redirect(TeamsUtils.url);
    }
  );
}

/**
 * Exchange an authorization code for tokens using the Microsoft OAuth2 token endpoint.
 *
 * @param code - the authorization code from the OAuth callback.
 * @param redirectUri - the redirect URI registered in Azure AD.
 * @returns the token response from Microsoft.
 */
async function exchangeCode(
  code: string,
  redirectUri: string
): Promise<Record<string, unknown>> {
  const body = new URLSearchParams({
    grant_type: "authorization_code",
    client_id: env.MICROSOFT_TEAMS_CLIENT_ID!,
    client_secret: env.MICROSOFT_TEAMS_CLIENT_SECRET!,
    code,
    redirect_uri: redirectUri,
    scope: "openid profile email",
  });

  const tenantId = env.MICROSOFT_TENANT_ID ?? "common";
  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  return response.json();
}

export default router;

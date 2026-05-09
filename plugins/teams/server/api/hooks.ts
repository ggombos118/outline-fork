import crypto from "crypto";
import { escapeRegExp } from "es-toolkit/compat";
import Router from "koa-router";
import { IntegrationService, IntegrationType } from "@shared/types";
import { AuthenticationError, NotFoundError } from "@server/errors";
import Logger from "@server/logging/Logger";
import auth from "@server/middlewares/authentication";
import validate from "@server/middlewares/validate";
import {
  Integration,
  Collection,
  SearchQuery,
  Team,
  User,
} from "@server/models";
import { authorize } from "@server/policies";
import { presentIntegration } from "@server/presenters";
import SearchProviderManager from "@server/utils/SearchProviderManager";
import { safeEqual } from "@server/utils/crypto";
import type { APIContext } from "@server/types";
import env from "../env";
import { presentSearchCard } from "../presenters/messageCard";
import * as T from "./schema";

const router = new Router();

/**
 * Creates a new Teams Incoming Webhook integration for a collection.
 * The user must be authenticated and have permission to update their team.
 */
router.post(
  "teams.post",
  auth(),
  validate(T.TeamsPostSchema),
  async (ctx: APIContext<T.TeamsPostReq>) => {
    const { collectionId, url, channelName } = ctx.input.body;
    const { user } = ctx.state.auth;

    const collection = await Collection.findByPk(collectionId, {
      userId: user.id,
    });

    if (!collection) {
      throw NotFoundError("Collection not found");
    }

    authorize(user, "read", collection);
    authorize(user, "update", user.team);

    const integration =
      await Integration.create<Integration<IntegrationType.Post>>({
        service: IntegrationService.MicrosoftTeams,
        type: IntegrationType.Post,
        userId: user.id,
        teamId: user.teamId,
        collectionId,
        events: ["documents.update", "documents.publish"],
        settings: {
          url,
          channel: channelName,
        },
      });

    ctx.body = {
      data: presentIntegration(integration),
    };
  }
);

/**
 * Handles incoming requests from Teams Outgoing Webhooks.
 * Verifies the HMAC signature and processes the search query.
 */
router.post(
  "teams.hook",
  validate(T.TeamsHookSchema),
  async (ctx: APIContext<T.TeamsHookReq>) => {
    verifyTeamsHmac(ctx);

    const { text, from, tenantId, conversation } = ctx.input.body;

    // Strip @mention tags, all remaining HTML tags, and decode common HTML entities
    const query = text
      .replace(/<at>[^<]*<\/at>/gi, "")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/gi, " ")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"')
      .trim();

    if (!query || query.toLowerCase() === "help") {
      ctx.body = {
        type: "message",
        text: "Use `@Outline <search term>` to search your workspace.",
      };
      return;
    }

    // tenantId may be null at the top level; fall back to conversation.tenantId
    const serviceTeamId = tenantId ?? conversation?.tenantId ?? from?.id ?? "";
    const serviceUserId = from?.aadObjectId ?? from?.id ?? "";

    const user = await findUserForRequest(serviceTeamId, serviceUserId);

    if (!user) {
      Logger.debug("plugins", "No Teams user found for request", {
        serviceTeamId,
        serviceUserId,
      });

      const team = await findTeamForRequest(serviceTeamId);
      const linkBase = team?.url ?? env.URL;
      const linkPath = `/settings/integrations/teams`;

      ctx.body = {
        type: "message",
        text: `It looks like you haven't linked your account yet. [Link your account](${linkBase}${linkPath}) to search from Teams.`,
      };
      return;
    }

    const { results, total } =
      await SearchProviderManager.getProvider().searchForUser(user, {
        query,
        limit: 5,
      });

    await SearchQuery.create({
      userId: user.id,
      teamId: user.teamId,
      source: "api",
      query,
      results: total,
    });

    const searchResults = results.map((result) => {
      const queryIsInTitle = !!result.document.title
        .toLowerCase()
        .match(escapeRegExp(query.toLowerCase()));
      return {
        title: result.document.title,
        url: result.document.url,
        context: queryIsInTitle ? undefined : result.context,
      };
    });

    ctx.body = presentSearchCard(searchResults, query, user.team.url);
  }
);

/**
 * Verify the HMAC-SHA256 signature from a Teams Outgoing Webhook request.
 * Teams computes HMAC-SHA256 over the JSON body using the webhook token as key (base64-decoded).
 *
 * @throws AuthenticationError if the secret is not configured or the signature is invalid.
 */
function verifyTeamsHmac(ctx: APIContext) {
  if (!env.MICROSOFT_BOT_SECRET) {
    throw AuthenticationError(
      "MICROSOFT_BOT_SECRET is not present in environment"
    );
  }

  const authHeader = ctx.request.headers.authorization ?? "";
  const match = authHeader.match(/^HMAC\s+(.+)$/i);

  if (!match) {
    throw AuthenticationError("Missing or invalid Authorization header");
  }

  const receivedSignature = match[1];

  let keyBuffer: Buffer;
  try {
    keyBuffer = Buffer.from(env.MICROSOFT_BOT_SECRET, "base64");
  } catch {
    throw AuthenticationError("MICROSOFT_BOT_SECRET is not valid base64");
  }

  const bodyStr = JSON.stringify(ctx.request.body);
  const expectedSignature = crypto
    .createHmac("sha256", keyBuffer)
    .update(Buffer.from(bodyStr, "utf8"))
    .digest("base64");

  if (!safeEqual(expectedSignature, receivedSignature)) {
    throw AuthenticationError("Invalid HMAC signature");
  }
}

/**
 * Find a matching Outline team for the given Teams tenant ID.
 *
 * @param serviceTeamId - the Teams tenant ID.
 * @returns the matching Outline Team, if found.
 */
async function findTeamForRequest(
  serviceTeamId: string
): Promise<Team | undefined> {
  const integration = await Integration.findOne({
    where: {
      service: IntegrationService.MicrosoftTeams,
      type: IntegrationType.LinkedAccount,
      settings: {
        teams: {
          serviceTeamId,
        },
      },
    },
    include: [
      {
        model: Team,
        as: "team",
        required: true,
      },
    ],
  });

  return integration?.team;
}

/**
 * Find a matching Outline user for the given Teams tenant and user IDs.
 *
 * @param serviceTeamId - the Teams tenant ID.
 * @param serviceUserId - the Teams user AAD object ID.
 * @returns the matching Outline User, if found.
 */
async function findUserForRequest(
  serviceTeamId: string,
  serviceUserId: string
): Promise<User | undefined> {
  const integration = await Integration.findOne({
    where: {
      service: IntegrationService.MicrosoftTeams,
      type: IntegrationType.LinkedAccount,
      settings: {
        teams: {
          serviceTeamId,
          serviceUserId,
        },
      },
    },
    include: [
      {
        model: User,
        as: "user",
        required: true,
      },
      {
        model: Team,
        as: "team",
        required: true,
      },
    ],
    order: [["createdAt", "DESC"]],
  });

  if (integration) {
    integration.user.team = integration.team;
    return integration.user;
  }

  return undefined;
}

export default router;

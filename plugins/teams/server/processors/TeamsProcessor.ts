import { differenceInMilliseconds } from "date-fns";
import { Op } from "sequelize";
import { IntegrationService, IntegrationType } from "@shared/types";
import { Minute } from "@shared/utils/time";
import { Document, Integration, Collection, Team } from "@server/models";
import BaseProcessor from "@server/queues/processors/BaseProcessor";
import type {
  DocumentEvent,
  IntegrationEvent,
  RevisionEvent,
  Event,
} from "@server/types";
import fetch from "@server/utils/fetch";
import { sleep } from "@shared/utils/timers";
import env from "../env";
import {
  presentMessageCard,
  presentWelcomeCard,
} from "../presenters/messageCard";

export default class TeamsProcessor extends BaseProcessor {
  static applicableEvents: Event["name"][] = [
    "documents.publish",
    "revisions.create",
    "integrations.create",
  ];

  async perform(event: Event) {
    switch (event.name) {
      case "documents.publish":
      case "revisions.create":
        // Wait a few seconds to give the document summary chance to be generated.
        await sleep(5000);
        return this.documentUpdated(event);

      case "integrations.create":
        return this.integrationCreated(event);

      default:
    }
  }

  async integrationCreated(event: IntegrationEvent) {
    const integration = (await Integration.findOne({
      where: {
        id: event.modelId,
        service: IntegrationService.MicrosoftTeams,
        type: IntegrationType.Post,
      },
      include: [
        {
          model: Collection,
          required: true,
          as: "collection",
        },
      ],
    })) as Integration<IntegrationType.Post>;

    if (!integration) {
      return;
    }

    const collection = integration.collection;
    if (!collection) {
      return;
    }

    const collectionUrl = `${env.URL}${collection.path}`;

    await fetch(integration.settings.url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(
        presentWelcomeCard(
          collection.name,
          collectionUrl,
          env.APP_NAME,
          collection.color
        )
      ),
    });
  }

  async documentUpdated(event: DocumentEvent | RevisionEvent) {
    // Never send notifications when batch importing documents.
    if (event.name === "documents.publish" && event.data?.source === "import") {
      return;
    }

    const [document, team] = await Promise.all([
      Document.scope("withCollection").findByPk(event.documentId),
      Team.findByPk(event.teamId),
    ]);

    if (!document || !team) {
      return;
    }

    // Never send notifications for draft documents.
    if (!document.publishedAt) {
      return;
    }

    // If the document was published less than a minute ago, don't send a
    // separate notification for the initial revision.
    if (
      event.name === "revisions.create" &&
      differenceInMilliseconds(document.updatedAt, document.publishedAt) <
        Minute.ms
    ) {
      return;
    }

    const integration = (await Integration.findOne({
      where: {
        teamId: document.teamId,
        collectionId: document.collectionId,
        service: IntegrationService.MicrosoftTeams,
        type: IntegrationType.Post,
        events: {
          [Op.contains]: [
            event.name === "revisions.create" ? "documents.update" : event.name,
          ],
        },
      },
    })) as Integration<IntegrationType.Post>;

    if (!integration) {
      return;
    }

    await fetch(integration.settings.url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(
        presentMessageCard(
          document,
          team,
          document.collection,
          event.name === "documents.publish"
        )
      ),
    });
  }
}

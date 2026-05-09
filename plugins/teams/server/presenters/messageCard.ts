import type { Collection, Document, Team } from "@server/models";

/**
 * Present a document as a Teams MessageCard for channel notifications.
 *
 * @param document - the document to present.
 * @param team - the team the document belongs to.
 * @param collection - the collection the document belongs to.
 * @param isNew - true if the document was just published, false if updated.
 * @returns a Teams MessageCard object.
 */
export function presentMessageCard(
  document: Document,
  team: Team,
  collection?: Collection | null,
  isNew?: boolean
) {
  const summary = document.getSummary();
  const collectionName = collection?.name;
  const authorName = isNew
    ? document.createdBy?.name
    : document.updatedBy?.name;

  const action = isNew
    ? `${authorName ? `${authorName} published` : "Published"}${collectionName ? ` to ${collectionName}` : ""}`
    : `${authorName ? `${authorName} updated` : "Updated"}${collectionName ? ` in ${collectionName}` : ""}`;

  // Purple for new documents, amber for updates
  const themeColor = isNew ? "6264A7" : "F0A30A";

  const facts = [
    { name: "Document", value: `[${document.title}](${team.url}${document.url})` },
  ];

  return {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    themeColor,
    summary: document.title,
    sections: [
      {
        activityTitle: action,
        facts,
        text: summary || undefined,
        markdown: true,
      },
    ],
    potentialAction: [
      {
        "@type": "OpenUri",
        name: "View Document",
        targets: [{ os: "default", uri: `${team.url}${document.url}` }],
      },
    ],
  };
}

/**
 * Present a welcome MessageCard sent when a Teams integration is first connected.
 *
 * @param collectionName - the name of the connected collection.
 * @param collectionUrl - the URL to the collection in Outline.
 * @param appName - the name of the Outline instance.
 * @param color - the collection color.
 * @returns a Teams MessageCard object.
 */
export function presentWelcomeCard(
  collectionName: string,
  collectionUrl: string,
  appName: string,
  color?: string | null
) {
  const themeColor = color?.replace("#", "") ?? "6264A7";

  return {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    themeColor,
    summary: `${appName} connected`,
    sections: [
      {
        activityTitle: `👋 Connected to ${appName}!`,
        activityText: `Documents published or updated in the **${collectionName}** collection will be posted to this channel.`,
        markdown: true,
      },
    ],
    potentialAction: [
      {
        "@type": "OpenUri",
        name: "View Collection",
        targets: [{ os: "default", uri: collectionUrl }],
      },
    ],
  };
}

/**
 * Present search results as a Teams bot message (Outgoing Webhook response format).
 *
 * @param results - array of result objects with title, url, and optional context.
 * @param query - the search query string.
 * @param teamUrl - the base URL of the Outline instance.
 * @returns a Teams bot message object.
 */
export function presentSearchCard(
  results: Array<{ title: string; url: string; context?: string }>,
  query: string,
  teamUrl: string
) {
  if (results.length === 0) {
    return {
      type: "message",
      text: `No documents found for **${query}**.`,
    };
  }

  const lines = results.map((result) => {
    const context = result.context
      ? ` — ${result.context.replace(/<\/?b>/g, "**").replace(/<[^>]+>/g, "").trim()}`
      : "";
    return `- [${result.title}](${teamUrl}${result.url})${context}`;
  });

  return {
    type: "message",
    text: `**Search results for "${query}":**\n\n${lines.join("\n")}`,
  };
}

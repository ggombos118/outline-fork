import { z } from "zod";

export const TeamsPostSchema = z.object({
  body: z.object({
    /** The collection to connect to Teams. */
    collectionId: z.uuid(),
    /** The Teams Incoming Webhook URL. */
    url: z.url(),
    /** The display name of the Teams channel. */
    channelName: z.string().min(1).max(255),
  }),
});

export type TeamsPostReq = z.infer<typeof TeamsPostSchema>;

export const TeamsHookSchema = z.object({
  body: z.object({
    /** The text of the message sent to the bot. */
    text: z.string().default(""),
    /** The ID of the Teams channel. */
    channelId: z.string().optional(),
    /** The display name of the Teams channel. */
    channelName: z.string().optional(),
    /** The ID of the Teams tenant. */
    tenantId: z.string().optional(),
    /** Information about the user who sent the message. */
    from: z
      .object({
        id: z.string(),
        name: z.string().nullish(),
        aadObjectId: z.string().nullish(),
      })
      .optional(),
    /** The Teams conversation reference. */
    conversation: z
      .object({
        id: z.string(),
        tenantId: z.string().nullish(),
        name: z.string().nullish(),
      })
      .optional(),
  }),
});

export type TeamsHookReq = z.infer<typeof TeamsHookSchema>;

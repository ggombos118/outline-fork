# Microsoft Teams Integration

This document describes the Microsoft Teams integration for Outline, covering all changes made to the codebase, files created, and setup instructions for self-hosted deployments.

---

## Changes to Existing Files

### `app/utils/PluginManager.ts`

Two fixes required to make the Teams plugin appear in the Integrations settings page:

1. **MobX static observable**: The `loaded` flag was declared as `@observable static loaded = false`, which the OXC/rolldown bundler compiles to a plain boolean (`sN.loaded = false`), bypassing MobX reactivity. Changed to `private static _loaded = observable.box(false)` with a getter, which survives bundling.

2. **MobX 4 observable arrays and `sortBy`**: MobX 4 observable arrays are not real JavaScript arrays (`Array.isArray()` returns `false`), so `sortBy` from `es-toolkit/compat` silently returned `[]` for them. Fixed by calling `.slice()` before `sortBy` to get a plain array:
   ```typescript
   return sortBy(arr ? arr.slice() : [], "priority") as Plugin<T>[];
   ```

### `app/hooks/useSettingsConfig.ts`

The settings navigation config is computed inside a `useComputed` hook, but it did not react to the plugin loading state. When the Teams plugin registered its settings hook asynchronously after the initial render, the hook list was already frozen. Fixed by adding a `pluginsLoaded` state variable driven by a MobX `reaction()` that fires when `PluginManager.isLoaded` becomes true, and including it in the `useComputed` dependency array so the config recomputes once plugins are ready.

### `shared/types.ts`

Three targeted changes:

1. **Added `MicrosoftTeams` to the `IntegrationService` enum:**
   ```typescript
   MicrosoftTeams = "microsoft-teams",
   ```

2. **Made `channelId` optional in `IntegrationSettings<IntegrationType.Post>`:**
   ```typescript
   // Before
   ? { url: string; channel: string; channelId: string }
   // After
   ? { url: string; channel: string; channelId?: string }
   ```
   The Slack post integration always provides `channelId` from its OAuth response; Teams incoming webhooks do not expose a channel ID.

3. **Added `teams` key to `IntegrationSettings<IntegrationType.LinkedAccount>`:**
   ```typescript
   teams?: { serviceTeamId: string; serviceUserId: string };
   ```
   Alongside the existing `slack` and `figma` keys, this stores the Azure AD tenant ID and user object ID after Microsoft account linking.

---

## New Files Created

### Plugin Metadata

| File | Description |
|---|---|
| `plugins/teams/plugin.json` | Plugin ID, name, description, and priority used by the plugin loader |

### Server — Configuration

| File | Description |
|---|---|
| `plugins/teams/server/env.ts` | Declares four optional environment variables: `MICROSOFT_TEAMS_CLIENT_ID` (public), `MICROSOFT_TEAMS_CLIENT_SECRET`, `MICROSOFT_TENANT_ID` (public), and `MICROSOFT_BOT_SECRET` |

### Server — Entry Point

| File | Description |
|---|---|
| `plugins/teams/server/index.ts` | Registers `Hook.API` and `Hook.Processor` unconditionally. Registers `Hook.AuthProvider` only when `MICROSOFT_TEAMS_CLIENT_ID` and `MICROSOFT_TEAMS_CLIENT_SECRET` are both set |

### Server — API Routes

| File | Description |
|---|---|
| `plugins/teams/server/api/hooks.ts` | Two routes: `POST teams.post` (create webhook integration) and `POST teams.hook` (receive Outgoing Webhook commands with HMAC verification). The HMAC is verified using `JSON.stringify(ctx.request.body)` — Teams sends standard JSON and the re-serialised body matches the signed bytes in practice. The `teams.hook` handler strips HTML tags and entities from the incoming message text before searching, and falls back to `conversation.tenantId` when the top-level `tenantId` field is absent or null (which Teams sometimes omits in channel contexts) |
| `plugins/teams/server/api/schema.ts` | Zod validation schemas for both routes (`TeamsPostSchema`, `TeamsHookSchema`). Fields nested inside `from` and `conversation` use `.nullish()` instead of `.optional()` because Teams sends explicit `null` values for absent fields |

### Server — Authentication

| File | Description |
|---|---|
| `plugins/teams/server/auth/teams.ts` | OAuth callback route (`GET teams.callback`) for Microsoft account linking. Exchanges the authorization code for an `id_token`, decodes the JWT to extract the AAD tenant ID (`tid`) and user object ID (`oid`), then creates a `LinkedAccount` integration |

### Server — Event Processor

| File | Description |
|---|---|
| `plugins/teams/server/processors/TeamsProcessor.ts` | Subscribes to `documents.publish`, `revisions.create`, and `integrations.create`. Posts Teams MessageCard payloads to the configured webhook URL when documents in connected collections are published or updated. Documents are fetched with the `withCollection` scope (not included in the default scope) so the collection name is available for the notification card. The event type is passed to `presentMessageCard` as `isNew` to distinguish publish from update styling |

### Server — Presenters

| File | Description |
|---|---|
| `plugins/teams/server/presenters/messageCard.ts` | Three presenter functions: `presentMessageCard` (document notification), `presentWelcomeCard` (sent on integration creation), `presentSearchCard` (Outgoing Webhook search results). Channel notifications use `MessageCard` format; search results use a plain `{ type: "message", text: "..." }` response because Teams Outgoing Webhooks do not render `MessageCard` payloads in bot replies. `presentMessageCard` distinguishes publish from update events: the `activityTitle` shows the action and author (e.g. "Gergely Gombos published to Engineering"), and the document title appears as a labelled `facts` entry with a clickable link. Theme colour is purple (`6264A7`) for new documents and amber (`F0A30A`) for updates |

### Shared Utilities

| File | Description |
|---|---|
| `plugins/teams/shared/TeamsUtils.ts` | OAuth state serialisation/deserialisation, URL builders (`callbackUrl`, `authUrl`, `errorUrl`), and the `TeamsOAuthNonceCookie` constant used for CSRF protection. The OAuth authorisation URL and token exchange endpoint both use the tenant-specific `/oauth2/v2.0/` path (`https://login.microsoftonline.com/<tenantId>/...`) when `MICROSOFT_TENANT_ID` is set, falling back to `/common` for multi-tenant apps |

### Client — Plugin Registration

| File | Description |
|---|---|
| `plugins/teams/client/index.tsx` | Registers `Hook.Settings` (the settings page) and `Hook.Icon` (the Teams icon) with the client-side plugin manager |

### Client — Settings Page

| File | Description |
|---|---|
| `plugins/teams/client/Settings.tsx` | Main settings UI. Three sections: (1) personal Microsoft account linking (shown only when OAuth is configured), (2) Outgoing Webhook setup information with the webhook endpoint URL, (3) collection list with inline webhook URL forms for unconnected collections and `TeamsListItem` for connected ones. Both `collections.fetchAll()` and `integrations.fetchAll()` are called on mount — the latter is required so that a freshly linked `LinkedAccount` integration is reflected in the UI without a page reload |

### Client — Components

| File | Description |
|---|---|
| `plugins/teams/client/Icon.tsx` | Microsoft Teams SVG icon with monochrome and colour variants |
| `plugins/teams/client/components/TeamsButton.tsx` | Connect button that generates a Microsoft OAuth authorisation URL with a CSRF nonce and redirects the browser |
| `plugins/teams/client/components/TeamsListItem.tsx` | Row for a connected collection showing the Teams channel name, publish/update event toggles, and a disconnect button |

---

## Architecture Overview

```
Outline server                     Microsoft Teams
──────────────────────────────     ────────────────────────────────
TeamsProcessor                ──▶  Incoming Webhook URL  (POST)
  (documents.publish)              (channel notifications)

POST /api/teams.hook          ◀──  Outgoing Webhook
  (HMAC-SHA256 verified)           (@Outline <query>)
  → SearchProviderManager
  → MessageCard response

GET /auth/teams.callback      ◀──  Microsoft OAuth redirect
  (code exchange)                  (account linking)
  → create LinkedAccount
```

**Key differences from the Slack integration:**

| | Slack | Teams |
|---|---|---|
| Webhook URL for notifications | Returned from OAuth response | Entered manually by user |
| Notification format | Slack Attachments | Teams MessageCard JSON |
| Bot/command trigger | Slash command (`/outline`) | Outgoing Webhook (`@Outline`) |
| Command authentication | Shared verification token | HMAC-SHA256 over raw request body |
| Search response format | Slack Attachments | Plain `type: "message"` text (MessageCard not supported in bot replies) |
| OAuth dependency | Required for Post type | Optional (only for account linking) |

---

## Setup Instructions

### Feature 1: Channel Notifications (no credentials required)

This feature works without any environment variables. Notifications are delivered via Teams Incoming Webhooks.

**In Microsoft Teams:**
1. Open the channel you want to receive notifications in.
2. Click **…** (More options) next to the channel name → **Connectors**.
3. Find **Incoming Webhook** and click **Configure**.
4. Give it a name (e.g. "Outline") and optionally upload an icon.
5. Click **Create** and copy the generated webhook URL.

**In Outline:**
1. Go to **Settings → Integrations → Microsoft Teams**.
2. In the **Collections** section, find the collection you want to connect.
3. Paste the webhook URL into the **Incoming Webhook URL** field.
4. Enter the channel name for display purposes.
5. Click **Connect**.

A welcome card will appear in the Teams channel confirming the connection. Documents published or updated in that collection will now post a card to the channel. You can toggle **Document published** and **Document updated** events independently.

---

### Feature 2: Search from Teams (Outgoing Webhook + account linking)

This lets team members type `@Outline <search term>` in any Teams channel and receive search results in a card.

#### Step A — Configure the Outgoing Webhook in Teams

1. Go to **Teams Admin Center** → **Teams apps** → **Manage apps** (or from within Teams: **Apps → More apps**).
2. Search for **Outgoing Webhooks** (available from the team settings menu: **… → Manage team → Apps → Create an outgoing webhook**).
3. Fill in:
   - **Name:** Outline
   - **Callback URL:** `https://<your-outline-url>/api/teams.hook`
   - **Description:** Search Outline from Teams
4. Click **Create**. Teams will display an **HMAC token** — copy it.

**In your Outline server environment:**
```bash
MICROSOFT_BOT_SECRET=<paste the HMAC token here>
```

Restart the Outline server after setting this variable.

#### Step B — Link personal Microsoft accounts

Users need to link their Outline account to their Microsoft identity so that search returns only documents they have access to.

1. Ensure `MICROSOFT_TEAMS_CLIENT_ID` and `MICROSOFT_TEAMS_CLIENT_SECRET` are configured (see Feature 3 below).
2. Each user goes to **Settings → Integrations → Microsoft Teams**.
3. Under **Personal account**, click **Connect Microsoft Account**.
4. Complete the Microsoft OAuth consent flow.

Without account linking, search commands from Teams will prompt the user to link their account first.

---

### Feature 3: Microsoft Account Linking via OAuth (optional)

Required only if you want the **Personal account** linking feature in the settings UI.

#### Step A — Register an Azure AD application

1. Sign in to the [Azure Portal](https://portal.azure.com).
2. Go to **Azure Active Directory → App registrations → New registration**.
3. Fill in:
   - **Name:** Outline
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts* (or restrict to your tenant as needed)
   - **Redirect URI:** `https://<your-outline-url>/auth/teams.callback`
4. Click **Register**.
5. Copy the **Application (client) ID** — this is your `MICROSOFT_TEAMS_CLIENT_ID`.
6. Go to **Certificates & secrets → New client secret**, create a secret, and copy the value — this is your `MICROSOFT_TEAMS_CLIENT_SECRET`.
7. Go to **API permissions → Add a permission → Microsoft Graph → Delegated permissions** and add:
   - `openid`
   - `profile`
   - `email`
8. Click **Grant admin consent** (or have a tenant admin do so).

#### Step B — Set environment variables

Add these to your Outline server configuration:

```bash
MICROSOFT_TEAMS_CLIENT_ID=<Application (client) ID>
MICROSOFT_TEAMS_CLIENT_SECRET=<Client secret value>
```

If your Azure app is registered as **single-tenant** (i.e. "Supported account types" is set to *Accounts in this organizational directory only*), also set:

```bash
MICROSOFT_TENANT_ID=<Directory (tenant) ID>
```

Without `MICROSOFT_TENANT_ID`, the OAuth flow uses the `/common` endpoint, which Microsoft rejects (`AADSTS50194`) for single-tenant applications. The tenant ID can be found on the app's **Overview** page in the Azure Portal.

Restart the Outline server. The **Personal account** section will now appear on the Teams integration settings page.

---

## Environment Variable Reference

| Variable | Required | Description |
|---|---|---|
| `MICROSOFT_TEAMS_CLIENT_ID` | No | Azure AD application client ID. Enables Microsoft account linking. Must be set together with `MICROSOFT_TEAMS_CLIENT_SECRET`. |
| `MICROSOFT_TEAMS_CLIENT_SECRET` | No | Azure AD application client secret. |
| `MICROSOFT_TENANT_ID` | No | Azure AD directory (tenant) ID. Required when the Azure app is registered as single-tenant. Without this, the OAuth flow uses the `/common` endpoint, which Microsoft rejects for single-tenant apps (`AADSTS50194`). |
| `MICROSOFT_BOT_SECRET` | No | HMAC token from the Teams Outgoing Webhook configuration. Required for `@Outline` search commands. |

All four variables are optional. The channel notification feature works with none of them set.

---

## Testing Checklist

- [x] Connect a collection webhook → welcome card appears in Teams channel
- [x] Publish a document → notification card appears in Teams channel
- [x] Update a document → notification card appears (when "Document updated" toggle is on)
- [x] Toggle "Document updated" off → no notification on next update
- [x] Disconnect integration → no further notifications
- [x] Type `@Outline <query>` in Teams → search results card returned (requires `MICROSOFT_BOT_SECRET`)
- [x] Link Microsoft account → visible in settings as connected (requires OAuth credentials)
- [x] Disconnect linked account → account removed from settings

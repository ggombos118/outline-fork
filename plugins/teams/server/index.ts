import { Hook, PluginManager } from "@server/utils/PluginManager";
import config from "../plugin.json";
import hooks from "./api/hooks";
import authRouter from "./auth/teams";
import env from "./env";
import TeamsProcessor from "./processors/TeamsProcessor";

// The Processor and API hooks are always registered: the Post integration
// (webhook URL entry) works without any environment variables.
PluginManager.add([
  {
    ...config,
    type: Hook.API,
    value: hooks,
  },
  {
    type: Hook.Processor,
    value: TeamsProcessor,
  },
]);

// The Auth provider (for LinkedAccount OAuth) is only registered when
// Microsoft OAuth credentials are configured.
if (env.MICROSOFT_TEAMS_CLIENT_ID && env.MICROSOFT_TEAMS_CLIENT_SECRET) {
  PluginManager.add({
    ...config,
    type: Hook.AuthProvider,
    value: { router: authRouter, id: config.id },
  });
}

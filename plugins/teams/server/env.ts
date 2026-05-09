import { IsOptional } from "class-validator";
import { Environment } from "@server/env";
import { Public } from "@server/utils/decorators/Public";
import environment from "@server/utils/environment";
import { CannotUseWithout } from "@server/utils/validators";

class TeamsPluginEnvironment extends Environment {
  /**
   * Microsoft OAuth2 client credentials. To enable account linking with Microsoft Teams.
   */
  @Public
  @IsOptional()
  @CannotUseWithout("MICROSOFT_TEAMS_CLIENT_SECRET")
  public MICROSOFT_TEAMS_CLIENT_ID = this.toOptionalString(
    environment.MICROSOFT_TEAMS_CLIENT_ID
  );

  @IsOptional()
  @CannotUseWithout("MICROSOFT_TEAMS_CLIENT_ID")
  public MICROSOFT_TEAMS_CLIENT_SECRET = this.toOptionalString(
    environment.MICROSOFT_TEAMS_CLIENT_SECRET
  );

  /**
   * Azure AD tenant ID. When set, the OAuth flow uses the tenant-specific endpoint
   * instead of /common. Required when the Azure app is single-tenant.
   */
  @Public
  @IsOptional()
  public MICROSOFT_TENANT_ID = this.toOptionalString(
    environment.MICROSOFT_TENANT_ID
  );

  /**
   * Secret token used to verify incoming requests from Teams Outgoing Webhooks.
   * Set this to the HMAC token provided when configuring the Outgoing Webhook in
   * Teams Admin Center.
   */
  @IsOptional()
  public MICROSOFT_BOT_SECRET = this.toOptionalString(
    environment.MICROSOFT_BOT_SECRET
  );
}

export default new TeamsPluginEnvironment();

import * as React from "react";
import { useTranslation } from "react-i18next";
import type { IntegrationType } from "@shared/types";
import Button from "~/components/Button";
import { generateOAuthStateNonce } from "~/utils/oauth";
import { redirectTo } from "~/utils/urls";
import { TeamsOAuthNonceCookie, TeamsUtils } from "../../shared/TeamsUtils";

type Props = {
  type: IntegrationType;
  state: { teamId: string };
  label?: string;
  icon?: React.ReactNode;
};

/**
 * A button that initiates the Microsoft OAuth flow for Teams account linking.
 */
function TeamsButton({ type, state: stateData, label, icon }: Props) {
  const { t } = useTranslation();

  const handleClick = () => {
    const nonce = generateOAuthStateNonce(TeamsOAuthNonceCookie);
    const { teamId, ...rest } = stateData;
    const state = TeamsUtils.createState(teamId, type, { nonce, ...rest });
    redirectTo(TeamsUtils.authUrl(state, TeamsUtils.callbackUrl()));
  };

  return (
    <Button onClick={handleClick} icon={icon} neutral>
      {label || t("Connect Microsoft Account")}
    </Button>
  );
}

export default TeamsButton;

import { observer } from "mobx-react";
import * as React from "react";
import { useTranslation, Trans } from "react-i18next";
import styled from "styled-components";
import { toast } from "sonner";
import { IntegrationService, IntegrationType } from "@shared/types";
import type Collection from "~/models/Collection";
import type Integration from "~/models/Integration";
import { ConnectedButton } from "~/scenes/Settings/components/ConnectedButton";
import { IntegrationScene } from "~/scenes/Settings/components/IntegrationScene";
import SettingRow from "~/scenes/Settings/components/SettingRow";
import Button from "~/components/Button";
import Flex from "~/components/Flex";
import Heading from "~/components/Heading";
import CollectionIcon from "~/components/Icons/CollectionIcon";
import Input from "~/components/Input";
import List from "~/components/List";
import ListItem from "~/components/List/Item";
import Notice from "~/components/Notice";
import Text from "~/components/Text";
import env from "~/env";
import useCurrentTeam from "~/hooks/useCurrentTeam";
import usePolicy from "~/hooks/usePolicy";
import useQuery from "~/hooks/useQuery";
import useStores from "~/hooks/useStores";
import { client } from "~/utils/ApiClient";
import TeamsIcon from "./Icon";
import TeamsButton from "./components/TeamsButton";
import TeamsListItem from "./components/TeamsListItem";

type WebhookFormState = {
  url: string;
  channelName: string;
  saving: boolean;
};

function Teams() {
  const team = useCurrentTeam();
  const { collections, integrations } = useStores();
  const { t } = useTranslation();
  const query = useQuery();
  const can = usePolicy(team);
  const error = query.get("error");

  const [webhookForms, setWebhookForms] = React.useState<
    Record<string, WebhookFormState>
  >({});

  React.useEffect(() => {
    void collections.fetchAll();
    void integrations.fetchAll();
  }, [collections, integrations]);

  const linkedAccountIntegration = integrations.find({
    type: IntegrationType.LinkedAccount,
    service: IntegrationService.MicrosoftTeams,
  });

  const groupedCollections = collections.orderedData
    .map<[Collection, Integration | undefined]>((collection) => {
      const integration = integrations.find({
        service: IntegrationService.MicrosoftTeams,
        collectionId: collection.id,
      });
      return [collection, integration];
    })
    .sort((a) => (a[1] ? -1 : 1));

  const appName = env.APP_NAME;
  const hasOAuth = !!env.MICROSOFT_TEAMS_CLIENT_ID;

  const getForm = (collectionId: string): WebhookFormState =>
    webhookForms[collectionId] ?? { url: "", channelName: "", saving: false };

  const updateForm = (
    collectionId: string,
    patch: Partial<WebhookFormState>
  ) => {
    setWebhookForms((prev) => ({
      ...prev,
      [collectionId]: { ...getForm(collectionId), ...patch },
    }));
  };

  const handleConnect =
    (collectionId: string) => async (ev: React.FormEvent) => {
      ev.preventDefault();
      const form = getForm(collectionId);

      if (!form.url) {
        return;
      }

      updateForm(collectionId, { saving: true });

      try {
        await client.post("/teams.post", {
          collectionId,
          url: form.url,
          channelName: form.channelName || t("Teams channel"),
        });
        await integrations.fetchAll();
        setWebhookForms((prev) => {
          const next = { ...prev };
          delete next[collectionId];
          return next;
        });
        toast.success(t("Teams channel connected"));
      } catch (err) {
        toast.error(err.message);
      } finally {
        updateForm(collectionId, { saving: false });
      }
    };

  return (
    <IntegrationScene title="Microsoft Teams" icon={<TeamsIcon />}>
      <Heading>Microsoft Teams</Heading>

      {error === "access_denied" && (
        <Notice>
          <Trans>
            Whoops, you need to accept the permissions in Microsoft to connect{" "}
            {{ appName }} to your account. Try again?
          </Trans>
        </Notice>
      )}
      {error === "unauthenticated" && (
        <Notice>
          <Trans>
            Something went wrong while authenticating your request. Please try
            logging in again.
          </Trans>
        </Notice>
      )}

      {hasOAuth && (
        <SettingRow
          name="link"
          label={t("Personal account")}
          description={
            <Trans>
              Link your {{ appName }} account to Microsoft Teams to enable
              searching for documents you have access to, directly within Teams.
            </Trans>
          }
        >
          <Flex align="flex-end" column>
            {linkedAccountIntegration ? (
              <ConnectedButton
                onClick={linkedAccountIntegration.delete}
                confirmationMessage={t(
                  "Disconnecting your personal account will prevent searching for documents from Teams. Are you sure?"
                )}
              />
            ) : (
              <TeamsButton
                type={IntegrationType.LinkedAccount}
                state={{ teamId: team.id }}
                label={t("Connect")}
              />
            )}
          </Flex>
        </SettingRow>
      )}

      {can.update && (
        <>
          <SettingRow
            name="command"
            border={false}
            label={t("Outgoing Webhook")}
            description={
              <Trans
                defaults="Search {{ appName }} from Teams using the <em>@Outline</em> mention. Configure an Outgoing Webhook in Teams Admin Center with your server's URL, then set <code>MICROSOFT_BOT_SECRET</code> to the provided HMAC token."
                values={{ appName }}
                components={{
                  em: <em />,
                  code: <Code />,
                }}
              />
            }
          >
            <Flex align="flex-end" column>
              <Text as="p" type="secondary">
                <Trans
                  defaults="Webhook URL: <code>{{ url }}/api/teams.hook</code>"
                  values={{ url: window.location.origin }}
                  components={{ code: <Code /> }}
                />
              </Text>
            </Flex>
          </SettingRow>

          <Heading as="h2">{t("Collections")}</Heading>
          <Text as="p" type="secondary">
            <Trans>
              Connect {{ appName }} collections to Teams channels. Messages will
              be automatically posted when documents are published or updated.
            </Trans>
          </Text>

          <List>
            {groupedCollections.map(([collection, integration]) => {
              if (integration) {
                return (
                  <TeamsListItem
                    key={integration.id}
                    collection={collection}
                    integration={
                      integration as Integration<IntegrationType.Post>
                    }
                  />
                );
              }

              const form = getForm(collection.id);

              return (
                <ListItem
                  key={collection.id}
                  title={
                    <Flex align="center" gap={6}>
                      <CollectionIcon collection={collection} />
                      {collection.name}
                    </Flex>
                  }
                  actions={
                    <WebhookForm onSubmit={handleConnect(collection.id)}>
                      <Input
                        type="text"
                        placeholder={t("Incoming Webhook URL")}
                        value={form.url}
                        onChange={(ev) =>
                          updateForm(collection.id, {
                            url: ev.currentTarget.value,
                          })
                        }
                        required
                        minLength={1}
                      />
                      <Input
                        type="text"
                        placeholder={t("Channel name")}
                        value={form.channelName}
                        onChange={(ev) =>
                          updateForm(collection.id, {
                            channelName: ev.currentTarget.value,
                          })
                        }
                      />
                      <Button
                        type="submit"
                        disabled={!form.url || form.saving}
                        neutral
                      >
                        {t("Connect")}
                      </Button>
                    </WebhookForm>
                  }
                />
              );
            })}
          </List>
        </>
      )}
    </IntegrationScene>
  );
}

const Code = styled.code`
  padding: 4px 6px;
  margin: 0 2px;
  background: ${(props) => props.theme.codeBackground};
  border-radius: 4px;
  font-size: 80%;
`;

const WebhookForm = styled.form`
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  justify-content: flex-end;
`;

export default observer(Teams);

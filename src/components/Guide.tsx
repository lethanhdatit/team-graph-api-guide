import { useState, useEffect } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import config from "../appSettings";
import { Client } from "@microsoft/microsoft-graph-client";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import "./Guide.css";

const LSR_SYSTEM_TEAM_ID_KEY = "SYSTEM_TEAM_ID";
const SSR_ACCESS_TOKEN_KEY = "AccessToken";
const FIXED_TEAM_NAME = "Fixed Team Name";

export function Guide() {
  const [logs, setLogs] = useState<{ n: number; t: string }[]>([]);
  const [graphClient, setGraphClient] = useState<Client | undefined>(undefined);
  const [accessToken, setAccessToken] = useState<string | undefined>(undefined);
  const [teamId, setTeamId] = useState<string | undefined>(
    localStorage.getItem(LSR_SYSTEM_TEAM_ID_KEY) ?? undefined
  );
  const [channelNameData, setChannelNameData] =
    useState<string>("Custom channel 01");
  const [memberEmails, setMemberEmails] = useState<string>("diem@modetour.com");
  const [channels, setChannels] = useState<object[]>([]);

  const log = (text: string) => {
    setLogs([
      ...logs,
      {
        n:
          logs && logs.length > 0
            ? logs.reduce((prev, current) =>
                prev.n > current.n ? prev : current
              ).n + 1
            : 0,
        t: text,
      },
    ]);
  };

  useEffect(() => {
    authorize().then();
  }, []);

  const authorize = async () => {
    log("Authorization is processing ...");
    setGraphClient(undefined);

    try {
      const authConfig = {
        clientId: config.clientId,
        initiateLoginEndpoint: config.initiateLoginEndpoint,
        cache: {
          cacheLocation: "localStorage",
        },
      };
      await new TeamsUserCredential(authConfig)!.login(config.apiScopes);
    } catch (error) {
      log("Authorization was failed. Please try again!");
    }

    setupGraphClient();
  };

  const setupGraphClient = () => {
    log("Graph client is initializing ...");
    const accessToken = sessionStorage.getItem(SSR_ACCESS_TOKEN_KEY);
    if (accessToken && accessToken.trim()) {
      try {
        const client = Client.init({
          authProvider: (done) => {
            done(null, accessToken);
          },
        });
        setGraphClient(client);
        setAccessToken(accessToken);
        log("Graph client ready for use!");
      } catch (error) {
        log(
          `Graph client was failed initialization: ex: ${JSON.stringify(error)}`
        );
      }
    } else
      log(
        "Graph client was failed initialization: 'AccessToken' is required, Let authorize first!"
      );
  };

  // GetTeamAsync in backend
  const getOrCreateTeam = async () => {
    if (!graphClient) {
      log("Graph client was not initialized");
      return;
    }

    const handle = async () => {
      if (!teamId) {
        const team = {
          displayName: FIXED_TEAM_NAME,
          description: FIXED_TEAM_NAME,
          visibility: "private",
          "template@odata.bind":
            "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        };

        let res = await graphClient!.api("/teams").post(team);
        if (!res?.id) {
          let joinedTeams = await graphClient!.api("/me/joinedTeams").get();
          res = joinedTeams?.value?.find(
            (f: any) => f.displayName === FIXED_TEAM_NAME
          );
        }
        log(`New team '${FIXED_TEAM_NAME}' was created successfully!`);
        return res;
      }

      return await graphClient!.api(`/teams/${teamId}`).get();
    };

    log("Getting or creating team is processing ...");
    const team = await handle();
    setTeamId(team.id);
    localStorage.setItem(LSR_SYSTEM_TEAM_ID_KEY, team.id);
    log(`Team ID: ${team.id}`);
  };

  // GetListAsync in backend
  const getListChannels = async (existedChannelName?: string) => {
    if (!graphClient) {
      log("Graph client was not initialized");
      return;
    }
    if (!teamId) {
      log("No any team existed!");
      return;
    }

    const handle = async () => {
      let res = await graphClient!.api(`/teams/${teamId}/channels`).get();
      let channels = res.value;
      if (channels && channels.length > 0) {
        channels = channels.filter(
          (f: any) =>
            f.displayName !== "General" &&
            (existedChannelName ? f.displayName === existedChannelName : true)
        );
      }
      return channels?.map((item: any) => ({
        displayName: item.displayName,
        id: item.id,
      }));
    };

    log("Getting list channel is processing ...");
    const chn = await handle();
    if (chn && chn.length > 0) setChannels(chn);
    log(`Getting list channel is success, (${chn.length}) count`);
    return chn;
  };

  // AddAsync in backend
  const addChannel = async () => {
    if (!graphClient) {
      log("Graph client was not initialized");
      return;
    }
    if (!teamId) {
      log("No any team existed!");
      return;
    }

    const handle = async (
      channelName: string | undefined,
      memberEmails: string[] | undefined
    ) => {
      if (!channelName || !channelName.trim()) {
        log("Channel name is required!");
        return;
      }

      const channels = await getListChannels(channelName);
      if (channels && channels.length > 0) {
        return channels[0];
      }

      log("Adding channel is continues to processing ...");

      const channel = {
        "@odata.type": "#Microsoft.Graph.channel",
        membershipType: "private",
        displayName: channelName,
      };

      const newChannel = await graphClient!
        .api(`/teams/${teamId}/channels`)
        .post(channel);

      if (newChannel && memberEmails && memberEmails.length > 0) {
        //todo add members
      }

      return newChannel;
    };

    log("Adding channel is processing ...");
    await handle(channelNameData, memberEmails?.split(","));
    await getListChannels();
    log(`'${channelNameData}' was added successfully`);
  };

  const render = () => (
    <div>
      <p>
        <b>Access token:</b>
      </p>
      <pre className="fixed">{accessToken}</pre>
      <Button appearance="primary" onClick={authorize}>
        Re-Authorize
      </Button>
      <br />
      <br />
      <p>
        <b>Team ID: </b>
        {teamId}
      </p>
      <p>
        <b>Team name: </b>
        {teamId && FIXED_TEAM_NAME}
      </p>
      <Button
        appearance="primary"
        disabled={!graphClient}
        onClick={getOrCreateTeam}
      >
        Get/create Team
      </Button>
      <br />
      <br />
      <br />
      <b>Channel list:</b>
      <ul>
        {channels.map((item: any, i) => (
          <li key={`${item.id}-${i}`}>
            {item.displayName} (ID: {item.id})
          </li>
        ))}
      </ul>
      <br />
      <Button
        appearance="primary"
        disabled={!graphClient}
        onClick={async () => await getListChannels()}
      >
        Get channels
      </Button>
      <br />
      <br />
      <br />
      <p>
        <b>New channel name: </b>{" "}
        <input
          onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
            setChannelNameData((e.target as HTMLInputElement).value)
          }
          type="email"
          style={{ width: "100%" }}
          defaultValue={channelNameData}
        />
      </p>
      <p>
        <b>Add members: </b>{" "}
        <input
          onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
            setMemberEmails((e.target as HTMLInputElement).value)
          }
          type="email"
          style={{ width: "100%" }}
          defaultValue={memberEmails}
        />
      </p>
      <br />
      <div className="control">
        <Button
          appearance="primary"
          disabled={!graphClient}
          onClick={async () => await addChannel()}
        >
          Add channel
        </Button>
      </div>
      <br />
      <div className="logs-pannel">
        <p>
          <b>Logs hictory:</b>
        </p>
        <div>
          <pre className="logs">
            {logs
              .sort((a, b) => b.n - a.n)
              .map((l, i) => {
                if (i === 0) {
                  return (
                    <p style={{fontSize: '16px'}} key={l.n}>
                      <b>{l.t}</b>
                    </p>
                  );
                }
                return <p key={l.n}>{l.t}</p>;
              })}
          </pre>
        </div>
      </div>
    </div>
  );

  return (
    <div className={"light"}>
      <div className="welcome page">
        <div className="narrow page-padding">
          <h1 className="center">Teams with Graph APIs Guidelines</h1>
          <div className="tabList">
            <div>{render()}</div>
          </div>
        </div>
      </div>
    </div>
  );
}

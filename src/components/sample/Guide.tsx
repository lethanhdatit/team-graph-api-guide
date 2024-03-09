import { useContext, useState, useEffect } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { TeamsFxContext } from "../Context";
import config from "./lib/config";
import { Client } from "@microsoft/microsoft-graph-client";

const LSR_SYSTEM_TEAM_ID_KEY = "SYSTEM_TEAM_ID";
const SSR_ACCESS_TOKEN_KEY = "AccessToken";
const FIXED_TEAM_NAME = "Fixed Team Name";

export function Guide() {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
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
    console.log(logs);
  }, [logs]);

  useEffect(() => {
    authorize().then();
  }, []);

  const authorize = async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    setGraphClient(undefined);
    log("Authorization is processing ...");
    await teamsUserCredential!.login(config.apiScopes);
    setupGraphClient();
  };

  const setupGraphClient = () => {
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
        return res;
      }

      return await graphClient!.api(`/teams/${teamId}`).get();
    };

    log("Getting or creating team is processing ...");
    const team = await handle();
    setTeamId(team.id);
    localStorage.setItem(LSR_SYSTEM_TEAM_ID_KEY, team.id);
    log(`Success! Team ID: ${team.id}`);
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
    log(`Adding channel is success`);
  };

  return (
    <div>
      <h2>Teams with Graph APIs Guidelines</h2>
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
        <b>Channel list:</b>
        <ul>
          {
            channels.map((item: any) => <li>{item.displayName} (ID: {item.id})</li>)
          }
        </ul>
      </p>
      <br />
      <p>
        <b>New channel name: </b>{" "}
        <input
          onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
            setChannelNameData((e.target as HTMLInputElement).value)
          }
          type="email"
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
          defaultValue={memberEmails}
        />
      </p>
      <br />
      <div className="control">
        <Button
          appearance="primary"
          disabled={!graphClient}
          onClick={getOrCreateTeam}
        >
          Get exist or create new Team
        </Button>
        <Button
          appearance="primary"
          disabled={!graphClient}
          onClick={async () => await getListChannels()}
        >
          Get list channel
        </Button>
        <Button
          appearance="primary"
          disabled={!graphClient}
          onClick={async () => await addChannel()}
        >
          Add channel
        </Button>
      </div>
      <br />
      <p>
        <b>Logs:</b>
      </p>
      <div>
        <pre className="logs">
          {logs
            .sort((a, b) => b.n - a.n)
            .map((l) => (
              <p key={l.n}>{l.t}</p>
            ))}
        </pre>
      </div>
    </div>
  );
}

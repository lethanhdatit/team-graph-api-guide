import { Client } from "@microsoft/microsoft-graph-client";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { IAppsettings } from "../appSettings";

const DEFAULT_CHANNEL_NAME = "General";

/**
 * Authorize the user in config's scopes that may required popup login or silent login.
 * After the authorization was success, we have the access token in session storage.
 * The final step will be performed in 'auth-end.html'.
 * @param config: {clientId, initiateLoginEndpoint, apiScopes}
 * @returns void
 */
const authorize = async (config: IAppsettings) => {
  if (!config) {
    throw new Error("Settings cannot be undefined");
  }

  const authConfig = {
    clientId: config.clientId,
    initiateLoginEndpoint: config.initiateLoginEndpoint,
    cache: {
      cacheLocation: "localStorage",
    },
  };
  await new TeamsUserCredential(authConfig).login(config.apiScopes);
};

/**
 * Initialize the graph client SDK that used to be call graph APIs
 * @param accessToken The token be provided after the successfully authorization flow
 * @returns 'Client' instance
 */
const initializeGraphClient = (accessToken: string): Client => {
  if (!accessToken || !accessToken.trim()) {
    throw new Error("Access token was invalid");
  }

  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
};

/**
 * Logic is the same with 'GetTeamAsync' in BE
 * @param graphClient The graph client instance
 * @param teamId use to fetch the existed team
 * @param teamName use to create the new team with this display name
 * @returns Object { id, displayName, ... }
 */
const getOrCreateNewTeam = async (
  graphClient: Client | undefined,
  teamId?: string,
  teamName?: string
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!teamName && !teamId) {
    throw new Error("Creating new team require the field 'teamName'");
  }

  // create new team
  if (!teamId) {
    const team = {
      displayName: teamName,
      description: teamName,
      visibility: "private",
      "template@odata.bind":
        "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
    };

    let res = await graphClient.api("/teams").post(team);
    if (!res?.id) {
      let joinedTeams = await graphClient.api("/me/joinedTeams").get();
      res = joinedTeams?.value?.find((f: any) => f.displayName === teamName);
    }
    return res;
  }

  // fetch the existed team
  return await graphClient.api(`/teams/${teamId}`).get();
};

/**
 * Logic is the same with 'GetListAsync' in BE
 * @param graphClient The graph client instance
 * @param teamId use to fetch the existed team
 * @param channelName use to find the existed team that match the display name
 * @returns Array { id, displayName }
 */
const getChannels = async (
  graphClient: Client | undefined,
  teamId: string | undefined,
  channelName?: string
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!teamId) {
    throw new Error("No any team be existed.");
  }

  let res = await graphClient!.api(`/teams/${teamId}/channels`).get();
  let channels = res.value;
  if (channels && channels.length > 0) {
    // skip the default channel and/or filter by display name
    channels = channels.filter(
      (f: any) =>
        f.displayName !== DEFAULT_CHANNEL_NAME &&
        (channelName ? f.displayName === channelName : true)
    );
  }

  return channels;
};

/**
 * Logic is the same with 'AddAsync' in BE
 * @param graphClient The graph client instance
 * @param teamId use to fetch the existed team
 * @param channelName use to define the display name of the new channel
 * @param memberEmails use to invite users to this channel following the emails
 * @returns Object { id, displayName, ... }
 */
const addChannel = async (
  graphClient: Client | undefined,
  teamId: string | undefined,
  channelName: string | undefined,
  memberEmails: string[] | undefined
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!teamId) {
    throw new Error("No any team be existed.");
  }
  if (!channelName || !channelName.trim()) {
    throw new Error("Adding new channel require the field 'channelName'");
  }

  // Existed channel case
  const channels = await getChannels(graphClient, teamId, channelName);
  if (channels && channels.length > 0) {
    const existedChannel = channels[0];
    // Add members to existed channel
    if (existedChannel && memberEmails && memberEmails.length > 0) {
      await addMembers(graphClient, teamId, existedChannel.id, memberEmails);
    }
    return existedChannel;
  }

  // New channel case
  const channel = {
    "@odata.type": "#Microsoft.Graph.channel",
    membershipType: "private",
    displayName: channelName,
  };

  const newChannel = await graphClient
    .api(`/teams/${teamId}/channels`)
    .post(channel);

  // Add members to new channel
  if (newChannel && memberEmails && memberEmails.length > 0) {
    await addMembers(graphClient, teamId, newChannel.id, memberEmails);
  }

  return newChannel;
};

/**
 * Logic is the same with 'AddMemberAsync' in BE
 * @param graphClient The graph client instance
 * @param teamId use to fetch the existed team
 * @param channelId use to define the channel for processing
 * @param memberEmails use to invite users to this channel following the emails
 * @returns void
 */
const addMembers = async (
  graphClient: Client | undefined,
  teamId: string | undefined,
  channelId: string | undefined,
  memberEmails: string[] | undefined
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!teamId) {
    throw new Error("No any team be existed.");
  }
  if (!channelId) {
    throw new Error("Adding members require the field 'channelId'");
  }
  if (!memberEmails || memberEmails.length === 0) {
    throw new Error("Email member not found.");
  }
};


/**
 * Fetch list Microsoft user by emails
 * @param graphClient The graph client instance
 * @param emails use to fetch MS user info
 * @returns Array {id, email...}
 */
const getUserByEmails = async (
  graphClient: Client | undefined,
  emails: string[] | undefined
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!emails || emails.length === 0) {
    throw new Error("Email list must not be empty.");
  }
};

export {
  authorize,
  initializeGraphClient,
  getOrCreateNewTeam,
  getChannels,
  addChannel,
  addMembers,
};

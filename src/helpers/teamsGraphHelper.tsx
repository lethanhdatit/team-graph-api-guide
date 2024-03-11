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
 * @param channelName use to find the existed channel that match the display name
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

  let query = `displayName ne '${DEFAULT_CHANNEL_NAME}' and membershipType eq 'private'`;
  if (channelName) {
    query += ` and displayName eq '${channelName}'`;
  }

  let res = await graphClient!
    .api(`/teams/${teamId}/channels`)
    .filter(query)
    .get();
  let channels = res.value;
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

  // Microsoft users search by emails
  const userIds = (await getUserByEmails(graphClient, memberEmails))?.map(
    (m: any) => m.id as string
  );

  // Members in team
  const res = await graphClient
    .api(`/teams/${teamId}/members`)
    .select(["microsoft.graph.aadUserConversationMember/userId"])
    .get();
  const memberIds = res?.value?.map((m: any) => m.userId) as string[];

  // Users need to add into the team
  const needAddUserIds = userIds.filter((f) => !memberIds.includes(f));

  // Perform add users into team
  if (needAddUserIds && needAddUserIds.length > 0) {
    const payload = needAddUserIds.map((id) => ({
      "@odata.type": "microsoft.graph.aadUserConversationMember",
      roles: [],
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${id}')`,
    }));

    await graphClient
      .api(`/teams/${teamId}/members/add`)
      .post({ values: payload });
  }

  // Perform add each user into channel
  for (let index = 0; index < userIds.length; index++) {
    const userId = userIds[index];
    const payload = {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      roles: [],
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`,
    };
    await graphClient
      .api(`/teams/${teamId}/channels/${channelId}/members`)
      .post(payload);
  }
};

/**
 * Logic is the same with 'AddMessageAsync' in BE
 * @param graphClient The graph client instance
 * @param teamId use to fetch the existed team
 * @param channelId use to define the channel for processing
 * @param content the content of the message, that can be html
 * @returns void
 */
const addMessage = async (
  graphClient: Client | undefined,
  teamId: string | undefined,
  channelId: string | undefined,
  content: string | undefined
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!teamId) {
    throw new Error("No any team be existed.");
  }
  if (!channelId) {
    throw new Error("Posting message require the field 'channelId'");
  }

  const payload = {
    body: {
      content: content,
      contentType: "html",
    },
  };

  await graphClient
    .api(`/teams/${teamId}/channels/${channelId}/messages`)
    .post(payload);
};

/**
 * Logic is the same with 'GetMessagesAsync' in BE
 * @param graphClient The graph client instance
 * @param teamId use to fetch the existed team
 * @param channelId use to define the channel for processing
 * @param currentPage the current index of the pagination
 * @param pageSize the item count for each page of the pagination
 * @returns void
 */
const getMessages = async (
  graphClient: Client | undefined,
  teamId: string | undefined,
  channelId: string | undefined,
  currentPage: number,
  pageSize: number
) => {
  if (!graphClient) {
    throw new Error("Graph client was not initialized");
  }
  if (!teamId) {
    throw new Error("No any team be existed.");
  }
  if (!channelId) {
    throw new Error("Posting message require the field 'channelId'");
  }

  // IMPORTANT: server side pagination, issue: messageType can not be filtered in backend so the result may not correct when pagination
  // const res = await graphClient
  //   .api(`/teams/${teamId}/channels/${channelId}/messages/delta`)
  //   .skip((currentPage - 1) * pageSize)
  //   .top(pageSize)
  //   .expand(["replies"])
  //   .get();

  // IMPORTANT:client side pagination, issue: not optimize performance
  const res = await graphClient
    .api(`/teams/${teamId}/channels/${channelId}/messages/delta`)
    .expand(["replies"])
    .get();

  let messages = res?.value as any[];
  if (messages && messages.length > 0) {
    const skip = (currentPage - 1) * pageSize;
    messages = messages
      .filter((f: any) => f.messageType === "message")
      .slice(skip, skip + pageSize);
  }

  console.log(messages);
  return messages as any[];
};

/**
 * Fetch list Microsoft user by emails
 * @param graphClient The graph client instance
 * @param emails use to fetch MS user info
 * @returns Array {id, email...}
 */
const getUserByEmails = async (graphClient: Client, emails: string[]) => {
  const maxCountEachTime = 2;
  let users = [] as object[];

  const processors = chunkArray(emails, maxCountEachTime);

  for (let index = 0; index < processors.length; index++) {
    // "mail in ('a1@gmail','a2@gmail','a3@gmail')"
    const query = `mail in (${processors[index]
      .map((m) => `'${m}'`)
      .join(",")})`;

    let res = await graphClient.api("/users").filter(query).get();
    if (res?.value) users = [...users, ...res.value];
  }

  return users;
};

const chunkArray = (array: any[], chunkSize: number): any[][] => {
  const chunks: string[][] = [];
  let index = 0;

  while (index < array.length) {
    chunks.push(array.slice(index, index + chunkSize));
    index += chunkSize;
  }

  return chunks;
};


export {
  authorize,
  initializeGraphClient,
  getOrCreateNewTeam,
  getChannels,
  addChannel,
  addMembers,
  addMessage,
  getMessages,
};

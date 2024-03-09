const config = {
  initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
  clientId: process.env.REACT_APP_CLIENT_ID,
  apiScopes: ["Channel.Create",
    "Channel.Delete.All",
    "Channel.ReadBasic.All",
    "ChannelMember.ReadWrite.All",
    "ChannelMessage.Read.All",
    "ChannelMessage.Send",
    "ChannelSettings.ReadWrite.All",
    "Chat.Create",
    "Chat.ReadBasic",
    "ChatMember.ReadWrite",
    "Files.ReadWrite",
    "Files.ReadWrite.All",
    "Team.Create",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All",
    "TeamsActivity.Send",
    "TeamSettings.ReadWrite.All",
    "User.Read",
    "User.ReadBasic.All"]
};

export default config;

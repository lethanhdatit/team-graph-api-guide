import { useState, useEffect } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { Client } from "@microsoft/microsoft-graph-client";
import config from "../appSettings";
import * as teamsGraphHelper from "../helpers/teamsGraphHelper";
import { ReactionType } from "../helpers/teamsGraphHelper";
import "./Guide.css";

const LSR_SYSTEM_TEAM_ID_KEY = "SYSTEM_TEAM_ID";
const SSR_ACCESS_TOKEN_KEY = "AccessToken";
const FIXED_TEAM_NAME = "Fixed Team Name 1";

export function Guide() {
  const [loading, setLoading] = useState<boolean>(false);
  const [logs, setLogs] = useState<{ n: number; t: string }[]>([]);
  const [graphClient, setGraphClient] = useState<Client | undefined>(undefined);
  const [accessToken, setAccessToken] = useState<string | undefined>(undefined);
  const [teamId, setTeamId] = useState<string | undefined>(
    localStorage.getItem(LSR_SYSTEM_TEAM_ID_KEY) ?? undefined
  );
  const [channelNameData, setChannelNameData] =
    useState<string>("Custom channel 01");
  const [channelIdData, setChannelIdData] = useState<string>();
  const [postData, setPostData] = useState<string>("test attachment");
  const [memberEmails, setMemberEmails] = useState<string>("diem@modetour.com,toandm@modetour.com,sonnv@modetour.com");
  const [channels, setChannels] = useState<object[]>([]);
  const [messages, setMessages] = useState<object[]>([]);
  const [messagePagination, setMessagePagination] = useState<{
    current: number;
    size: number;
  }>({
    current: 1,
    size: 10,
  });
  const [messageIdData, setMessageIdData] = useState<string>();
  const [messageAttachmentFiles, setMessageAttachmentFiles] =
    useState<File[] | null | undefined>();
  const [replyAttachmentFiles, setReplyAttachmentFiles] = useState<File[] | null | undefined>();
  const [replyContent, setReplyContent] = useState<string>();
  const [replyIdData, setReplyIdData] = useState<string | undefined>(undefined);
  const [reactionType, setReactionType] =
    useState<keyof typeof ReactionType>("HEART");

  useEffect(() => {
    const stretagy = async() => {
      await handleAuthorization()
    }
    stretagy()
  }, []);

  useEffect(() => {
    if(channels && channels.length > 0){
      setChannelIdData((channels[0] as any).id)
    }
  }, [channels]);

  const handleAuthorization = async () => {
    log("Authorization is processing ...");
    setLoading(true);
    setGraphClient(undefined);

    try {
      await teamsGraphHelper.authorize(config);
    } catch (error) {
      log("Authorization was failed. Please try again!");
      setLoading(false);
      return;
    }

    handleInitializationGraphClient();
    await handleFetchingTeam()
    await handleFetchingChannels()
    setLoading(false);
  };

  const handleInitializationGraphClient = () => {
    log("Graph client is initializing ...");
    setLoading(true);
    const accessToken = sessionStorage.getItem(SSR_ACCESS_TOKEN_KEY);
    if (accessToken && accessToken.trim()) {
      try {
        const client = teamsGraphHelper.initializeGraphClient(accessToken);

        setGraphClient(client);
        setAccessToken(accessToken);

        log("Graph client ready for use!");
      } catch (e: any) {
        log(`Graph client was failed initialization: ex: ${e.message}`);
      }
    } else
      log(
        "Graph client was failed initialization: 'AccessToken' is required, Let authorize first!"
      );
    setLoading(false);
  };

  const handleFetchingTeam = async () => {
    log("Fetching team is processing ...");
    setLoading(true);
    try {
      const team = await teamsGraphHelper.getOrCreateNewTeam(
        graphClient,
        teamId,
        FIXED_TEAM_NAME
      );
      setTeamId(team.id);
      localStorage.setItem(LSR_SYSTEM_TEAM_ID_KEY, team.id);
      log(`Team ID: ${team.id}`);
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleFetchingChannels = async () => {
    log("Getting list channel is processing ...");
    setLoading(true);
    try {
      const chn = await teamsGraphHelper.getChannels(graphClient, teamId);
      if (chn && chn.length > 0) setChannels(chn);
      log(`Getting list channel is success, (${chn.length}) count`);
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleAddingChannel = async () => {
    log("Adding channel is processing ...");
    setLoading(true);
    try {
      const members = memberEmails
        ?.split(",")
        .map((item) => item.trim())
        .filter((f) => f);
      await teamsGraphHelper.addChannel(
        graphClient,
        teamId,
        channelNameData,
        members
      );
      log(`'${channelNameData}' was added successfully`);
      await handleFetchingChannels();
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleAddingMembers = async () => {
    log("Adding members is processing ...");
    setLoading(true);
    try {
      const members = memberEmails
        ?.split(",")
        .map((item) => item.trim())
        .filter((f) => f);
      await teamsGraphHelper.addMembers(
        graphClient,
        teamId,
        channelIdData,
        members
      );
      log(
        `(${members.length}) members was added into channel '${channelIdData}'`
      );
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleAddingPost = async () => {
    log(`Posting message into channel '${channelIdData}' is processing ...`);
    setLoading(true);
    try {

      await teamsGraphHelper.addMessage(
        graphClient,
        teamId,
        channelIdData,
        postData,
        messageAttachmentFiles
      );
      log(`Posting new message into channel '${channelIdData}' successfully`);
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleFetchMessages = async () => {
    log(`Getting messages for channel '${channelIdData}' is processing ...`);
    setLoading(true);
    try {
      const mess = await teamsGraphHelper.getMessages(
        graphClient,
        teamId,
        channelIdData,
        messagePagination.current,
        messagePagination.size
      );
      setMessages(mess);
      log(
        `Getting messages for channel '${channelIdData}' current: ${messagePagination.current}, size: ${messagePagination.size}, count: ${mess.length}`
      );
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleReplyMessage = async () => {
    log(`Replying message into channel '${channelIdData}' is processing ...`);
    setLoading(true);
    try {
      await teamsGraphHelper.replyMessage(
        graphClient,
        teamId,
        channelIdData,
        messageIdData,
        replyContent,
        replyAttachmentFiles
      );
      log(`Replying new message into channel '${channelIdData}' successfully`);
      await handleFetchMessages();
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

  const handleReaction = async () => {
    log(`Reaction is processing ...`);
    setLoading(true);
    try {
      await teamsGraphHelper.setReaction(
        ReactionType[reactionType],
        graphClient,
        teamId,
        channelIdData,
        messageIdData,
        replyIdData
      );
      log(`Reaction successfully`);
      await handleFetchMessages();
    } catch (e: any) {
      log(e.message);
    }
    setLoading(false);
  };

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

  const summaryArrayByKey = (array: any[], key: string) => {
    return array.reduce((acc, item) => {
      acc[item[key]] = (acc[item[key]] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
  };

  const mapReactionType = (item: any) => {
    return Object.keys(item).reduce((acc, key) => {
      const newKey =
        ReactionType[key.toUpperCase() as keyof typeof ReactionType];
      if (newKey) {
        acc[newKey] = item[key];
      } else {
        acc[key] = item[key];
      }
      return acc;
    }, {} as Record<string, number>);
  };

  const readFileAsBase64 = (file: File, callBack: (base64: string | null) => void) => {
    const reader = new FileReader();

    reader.onload = () => {
        const base64String = reader.result?.toString().split(',')[1] || null;
        callBack(base64String);
    };

    reader.onerror = () => {
        console.error('Error occurred while reading file.');
    };

    reader.readAsDataURL(file);
};

  const render = () => (
    <div>
      <div className="authorize-pannel">
        <p>
          <b>Access token:</b>
        </p>
        <pre className="fixed">{accessToken}</pre>
        <Button
          disabled={loading}
          appearance="primary"
          onClick={handleAuthorization}
        >
          Re-Authorize
        </Button>
      </div>
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
        disabled={!graphClient || loading}
        onClick={handleFetchingTeam}
      >
        Get/Create Team
      </Button>
      <br />
      <br />
      <b>Channel list:</b>
      <ul>
        {channels.map((item: any, i) => (
          <li key={`${item.id}-${i}`}>
            <b>{item.id}</b>: {item.displayName}
          </li>
        ))}
      </ul>
      <Button
        appearance="primary"
        disabled={!graphClient || loading}
        onClick={async () => await handleFetchingChannels()}
      >
        Get channels
      </Button>
      <br />
      <br />
      <div>
        <b>Members: </b>(
        <i>
          separate by <b>","</b>
        </i>
        )
        <input
          onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
            setMemberEmails((e.target as HTMLInputElement).value)
          }
          disabled={loading}
          type="text"
          style={{ width: "100%" }}
          defaultValue={memberEmails}
        />
      </div>
      <br />
      <div className="channels">
        <div className="section">
          <div className="col-left">
            <b>New channel name: </b>
            <input
              onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
                setChannelNameData((e.target as HTMLInputElement).value)
              }
              disabled={loading}
              type="text"
              style={{ width: "100%" }}
              defaultValue={channelNameData}
            />
          </div>
          <div className="col-right">
            <Button
              appearance="primary"
              disabled={!graphClient || loading}
              onClick={async () => await handleAddingChannel()}
            >
              Add Channel
            </Button>
          </div>
        </div>
        <br />
        <div className="section">
          <div className="col-left">
            <b>Channel ID: </b>
            <input
              onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
                setChannelIdData((e.target as HTMLInputElement).value)
              }
              disabled={loading}
              type="text"
              style={{ width: "100%" }}
              defaultValue={channelIdData ?? (channels && channels.length > 0 ? (channels[0] as any).id : "")}
            />
          </div>
          <div className="col-right">
            <Button
              appearance="primary"
              disabled={!graphClient || loading}
              onClick={async () => await handleAddingMembers()}
            >
              Add Members
            </Button>
          </div>
        </div>
        <br />
        <div className="section column">
          <div className="col-left">
            <b>Message content: </b>
            <i>(text/html)</i>
            <textarea
              onKeyUp={(e: React.KeyboardEvent<HTMLTextAreaElement>) =>
                setPostData((e.target as HTMLTextAreaElement).value)
              }
              disabled={loading}
              rows={4}
              style={{ width: "100%" }}
              defaultValue={postData}
            />
          </div>
          <input
            disabled={loading}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
              const selectedFiles = e.target.files && Array.from(e.target.files);
              if (selectedFiles) {
                setMessageAttachmentFiles(selectedFiles)
              }
            }}
            multiple
            type="file"
            style={{ width: "100%" }}
          />
          <div className="col-right">
            <Button
              appearance="primary"
              disabled={!graphClient || loading}
              onClick={async () => await handleAddingPost()}
            >
              New Message
            </Button>
          </div>
        </div>
        <br />
        <div className="section column">
          <div className="col-left">
            <b>List messages: </b>
            <ul>
              {messages.map((m: any, i) => {
                return (
                  <li key={m.id}>
                    <b>{m.id}</b>-
                    {JSON.stringify(
                      mapReactionType(
                        summaryArrayByKey(m.reactions, "reactionType")
                      )
                    )}
                    {m.attachments.map((mm: any, i: any) => <span key={i}><a href={mm.contentUrl}>{mm.name}</a> | </span>)}
                    :{rawTextHtml(m.body.content)}
                    <ul>
                      {m.replies.map((r: any) => (
                        <li key={r.id}>
                          <b>{r.id}</b>-
                          {JSON.stringify(
                            mapReactionType(
                              summaryArrayByKey(r.reactions, "reactionType")
                            )
                          )}
                          {r.attachments.map((mm: any, i: any) => <span key={i}><a href={mm.contentUrl}>{mm.name}</a> | </span>)}
                          :{rawTextHtml(r.body.content)}
                        </li>
                      ))}
                    </ul>
                  </li>
                );
              })}
            </ul>
            <br />
            <b>Page .No: </b>
            <input
              onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
                setMessagePagination({
                  ...messagePagination,
                  ...{ current: Number((e.target as HTMLInputElement).value) },
                })
              }
              disabled={loading}
              type="number"
              style={{ width: "20%" }}
              defaultValue={messagePagination.current}
            />
            <span> </span>
            <b>Page size: </b>
            <input
              onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
                setMessagePagination({
                  ...messagePagination,
                  ...{ size: Number((e.target as HTMLInputElement).value) },
                })
              }
              disabled={loading}
              type="number"
              style={{ width: "20%" }}
              defaultValue={messagePagination.size}
            />
          </div>
          <div className="col-right">
            <Button
              appearance="primary"
              disabled={!graphClient || loading}
              onClick={async () => await handleFetchMessages()}
            >
              Get Messages
            </Button>
          </div>
        </div>
        <br />
        <div className="section column">
          <div className="col-left">
            <b>Message ID: </b>
            <input
              onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
                setMessageIdData((e.target as HTMLInputElement).value)
              }
              disabled={loading}
              type="text"
              style={{ width: "100%" }}
              defaultValue={messageIdData}
            />
            <b>Reply content: </b>
            <i>(text/html)</i>
            <textarea
              onKeyUp={(e: React.KeyboardEvent<HTMLTextAreaElement>) =>
                setReplyContent((e.target as HTMLTextAreaElement).value)
              }
              disabled={loading}
              rows={4}
              style={{ width: "100%" }}
              defaultValue={replyContent}
            />
            <input
              disabled={loading}
              onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                const selectedFiles = e.target.files && Array.from(e.target.files);
                if (selectedFiles) {
                  setReplyAttachmentFiles(selectedFiles)
                }
              }}
              multiple 
              type="file"
              style={{ width: "100%" }}
            />
          </div>
          <div className="col-right">
            <Button
              appearance="primary"
              disabled={!graphClient || loading}
              onClick={async () => await handleReplyMessage()}
            >
              Reply message
            </Button>
          </div>
        </div>
        <br />
        <div className="section column">
          <div className="col-left">
            <b>Reply ID: </b>
            <input
              onKeyUp={(e: React.KeyboardEvent<HTMLInputElement>) =>
                setReplyIdData((e.target as HTMLInputElement).value)
              }
              disabled={loading}
              type="text"
              style={{ width: "100%" }}
              defaultValue={replyIdData}
            />
            <b>Reaction type: </b>
            <select
              value={reactionType}
              onChange={(e) => {
                setReactionType(e.target.value as keyof typeof ReactionType);
              }}
            >
              {Object.keys(ReactionType).map((k: any) => {
                const type = ReactionType[k as keyof typeof ReactionType];
                return (
                  <option key={k} value={k}>
                    {type}
                  </option>
                );
              })}
            </select>
          </div>
          <div className="col-right">
            <Button
              appearance="primary"
              disabled={!graphClient || loading}
              onClick={async () => await handleReaction()}
            >
              Set Reaction
            </Button>
          </div>
        </div>
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
                    <p style={{ fontSize: "16px" }} key={l.n}>
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

  const rawTextHtml = (html: string) => (
    <span dangerouslySetInnerHTML={{ __html: html }} />
  );

  return (
    <div className={"light"}>
      <div className="welcome page">
        <div className="narrow page-padding">
          <h1 className="center">Teams with Graph APIs Guidelines</h1>
          <div className="tabList">
            {loading && (
              <div className="loading">
                <div className="area">
                  <Spinner />
                </div>
              </div>
            )}
            {render()}
          </div>
        </div>
      </div>
    </div>
  );
}

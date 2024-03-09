import { useContext, useState } from "react";
import {
  Image
} from "@fluentui/react-components";
import "./Welcome.css";
import { Guide } from "./Guide";
import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";


export function Welcome() {
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? "" : data!.displayName;

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <h1 className="center">
          Congratulations{userName ? ", " + userName : ""}!
        </h1>
        <div className="tabList">
          <div>
            <Guide />
          </div>
        </div>
      </div>
    </div>
  );
}

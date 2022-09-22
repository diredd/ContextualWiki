import {Image} from "@fluentui/react-northstar";
import {useData} from "@microsoft/teamsfx-react";
import {useContext} from "react";
import {TeamsFxContext} from "../Context";
import {Graph} from "./Graph";
import "./Welcome.css";

export function Welcome(props: {environment?: string}) {
  const {environment} = {
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment"
    }[environment] || "local environment";

  const {teamsfx} = useContext(TeamsFxContext);
  const {loading, data, error} = useData(async () => {
    if (teamsfx) {
      const userInfo = await teamsfx.getUserInfo();
      return userInfo;
    }
  });
  const userName = loading || error ? "" : data!.displayName;
  return (
    <div className='welcome page'>
      <div className='narrow page-padding'>
        <Image src='hello.png' />
        <h1 className='center'>
          Congratulations{userName ? ", " + userName : ""}!
        </h1>
        <p className='center'>
          Your app is running in your {friendlyEnvironmentName}
        </p>
        <div className='sections'>
          <Graph />
        </div>
      </div>
    </div>
  );
}

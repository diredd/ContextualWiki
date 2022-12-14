import {
  Avatar,
  Button,
  Chat,
  ChatMessage,
  PersonIcon
} from "@fluentui/react-northstar";
import {Providers, ProviderState} from "@microsoft/mgt-element";
import {Get, ResponseType} from "@microsoft/mgt-react";
import {TeamsFxProvider} from "@microsoft/mgt-teamsfx-provider";
import {app} from "@microsoft/teams-js";
import {useGraph} from "@microsoft/teamsfx-react";
import {useContext, useEffect, useState} from "react";
import {TeamsFxContext} from "../Context";
import "./Graph.css";
import {PersonCardGraphToolkit} from "./PersonCardGraphToolkit";

import {MgtTemplateProps} from "@microsoft/mgt-react";

const MyEvent = (props: MgtTemplateProps) => {
  console.log("props", props);
  let messages = props.dataContext.value.map((x: any) => ({
    gutter: <Avatar icon={<PersonIcon />} />,
    key: x.id,
    message: (
      <ChatMessage
        author={x.from.user.displayName}
        content={x.body.content}
        timestamp={x.createdTime}
      />
    )
  }));

  return <Chat items={messages} />;
};

export function Graph() {
  const {teamsfx} = useContext(TeamsFxContext);
  const [appContext, setappContext] = useState<app.Context | undefined>();
  const {loading, error, data, reload} = useGraph(
    async (graph, teamsfx, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();

      // Initialize Graph Toolkit TeamsFx provider
      const provider = new TeamsFxProvider(teamsfx, scope);
      Providers.globalProvider = provider;
      Providers.globalProvider.setState(ProviderState.SignedIn);
      let photoUrl = "";
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return {profile, photoUrl};
    },
    {scope: ["User.Read"], teamsfx: teamsfx}
  );

  useEffect(() => {
    app.getContext().then((context) => setappContext(context));
  }, []);

  return (
    <div>
      <div className='section-margin'>
        <Button primary content='Authorize' onClick={reload} />
        <PersonCardGraphToolkit loading={loading} data={data} error={error} />
        {appContext?.chat?.id && (
          <Get
            type={ResponseType.json}
            resource={`/chats/${appContext.chat?.id ?? ""}/messages`}
            version='beta'
            scopes={["Chat.Read"]}
            max-pages='4'>
            <MyEvent template='default'></MyEvent>
          </Get>
        )}
      </div>
    </div>
  );
}

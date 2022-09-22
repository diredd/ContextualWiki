import {PersonCard} from "@microsoft/mgt-react";
import {useContext} from "react";
import {TeamsFxContext} from "../Context";

export function PersonCardGraphToolkit(props: {
  loading?: boolean;
  error?: any;
  data?:
    | {
        profile: any;
        photoUrl: string;
      }
    | undefined;
}) {
  const {themeString} = useContext(TeamsFxContext);

  return (
    <div className='section-margin'>
      {!props.loading && props.error && (
        <div className='error'>
          Failed to read your profile. Please try again later. <br /> Details:{" "}
          {props.error.toString()}
        </div>
      )}
      {!props.loading && !props.error && props.data && (
        <div className={themeString === "default" ? "mgt-light" : "mgt-dark"}>
          <PersonCard personQuery='me' isExpanded={false}></PersonCard>
        </div>
      )}
    </div>
  );
}

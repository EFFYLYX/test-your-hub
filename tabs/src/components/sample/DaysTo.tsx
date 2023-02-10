import { useContext, useState } from "react";
import { Image, Menu } from "@fluentui/react-northstar";
import "./DaysTo.css";
import { EditCode } from "./EditCode";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";

export function DaysTo(props: { environment?: string }) {
  const { environment } = {
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const steps = ["local", "azure", "publish"];
  const friendlyStepsName: { [key: string]: string } = {
    local: "1. Build your app locally",
    azure: "2. Provision and Deploy to the Cloud",
    publish: "3. Publish to Teams",
  };
  const [selectedMenuItem, setSelectedMenuItem] = useState("local");
  const items = steps.map((step) => {
    return {
      key: step,
      content: friendlyStepsName[step] || "",
      onClick: () => setSelectedMenuItem(step),
    };
  });

  const { teamsfx } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsfx) {
      const userInfo = await teamsfx.getUserInfo();
      return userInfo;
    }
  });
  const userName = (loading || error) ? "": data!.displayName;


//   today = datetime.date.today()
//   friday = today + datetime.timedelta( (4-today.weekday()) % 7 )

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        {/* <Image src="hello.png" /> */}
        <h1 className="center">
          Hi{userName ? ", " + userName : ""}!
        </h1>
        {/* <p className="center">
          This app is running in your {friendlyEnvironmentName}
        </p>
        <p className="center">
          Let's test how your hub interacts with this MetaOS app!
        </p>
        <Menu defaultActiveIndex={0} items={items} underlined secondary />
        <div className="sections">
          {selectedMenuItem === "local" && (
            <div>
              <EditCode />
              <CurrentUser userName={userName} />
              <Graph />
            </div>
          )}
          {selectedMenuItem === "azure" && (
            <div>
              <Deploy />
            </div>
          )}
          {selectedMenuItem === "publish" && (
            <div>
              <Publish />
            </div>
          )}
        </div> */}
      </div>
    </div>
  );
}

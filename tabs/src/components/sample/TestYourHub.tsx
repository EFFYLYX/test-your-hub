import { useContext, useState } from "react";
import { Button, Image, Menu } from "@fluentui/react-northstar";
import "./DaysTo.css";
import { EditCode } from "./EditCode";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";
import { Authentication } from "./Authentication";
import { Deeplinks } from "./Deeplinks";
import { app, appInitialization, calendar } from "@microsoft/teams-js";


export function TestYourHub(props: { environment?: string }) {
  const { environment } = {
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const steps = ["authentication", "deeplinks", "calendar"];
  const friendlyStepsName: { [key: string]: string } = {
    // local: "1. Build your app locally",
    // azure: "2. Provision and Deploy to the Cloud",
    // publish: "3. Publish to Teams",
    authentication: "Authentication",
    deeplinks: "Deep links",
    calendar: "Calendar",
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
  app.notifySuccess();
  const userName = (loading || error) ? "" : data!.displayName;
  var actionResult = ""
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        {/* <Image src="hello.png" /> */}
        <h1 className="center">
          Hi{userName ? ", " + userName : ""}!
        </h1>
        <p className="center">
          This app is running in your {friendlyEnvironmentName}
        </p>
        <p className="center">
          Let's test how your hub interacts with this MetaOS app!
        </p>
      <Menu defaultActiveIndex={0} items={items} underlined secondary />
        <div className="sections">
          {selectedMenuItem === "authentication" && (
            <div>
              <h2>Authentication</h2>
              <Authentication />
            </div>
          )}
          {selectedMenuItem === "deeplinks" && (
            <div>
              <h2>Deep links</h2>
              <Deeplinks />
            </div>
          )}
          {selectedMenuItem === "calendar" && (
            <div>

            <Button primary content="composeMeeting" onClick={() => {
              if (calendar.isSupported()) {
                  const parms = {
                    attendees: ["joe@contoso.com", "bob@contoso.com"],
                    content: "test content",
                    startTime: new Date().toISOString(),
                    subject: "test subject",
                    endTime: "2023-10-24T10:30:00-07:00",
                  }
                  const promise = calendar.composeMeeting(parms);
                  promise.then((result) => {
                    window.alert("success"  );
                    console.log(result);  // "success"  
                  }, (error) => {
                    window.alert("error"  );
                    console.log(error);  // "error" 
                  });
              }
            }} />
            <p></p>

            <Button primary content="openCalendarItem" onClick={() => {
              if (calendar.isSupported()) {
                  const parms = {
                    itemId: "1234567"
                  }
                  const promise = calendar.openCalendarItem(parms);
                  promise.then((result) => {
                    // window.alert("success");
                    actionResult = "success"
                    console.log(result);  // "success"  
                  }, (error) => {
                    actionResult = "error"
                    console.log(error);  // "error" 
                  });
              }
            }} />
            <p>Action Result: {actionResult}</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

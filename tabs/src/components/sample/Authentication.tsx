import "./Graph.css";
import { useGraph } from "@microsoft/teamsfx-react";
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { Button } from "@fluentui/react-northstar";
import { Design } from './Design';
import { PersonCardFluentUI } from './PersonCardFluentUI';
import { PersonCardGraphToolkit } from './PersonCardGraphToolkit';
import { useContext } from "react";
import { TeamsFxContext } from "../Context";
import { app, authentication } from "@microsoft/teams-js";
import { Auth } from "./Auth";

export function Authentication() {
  const { teamsfx } = useContext(TeamsFxContext);
  const { loading, error, data, reload } = useGraph(
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
      return { profile, photoUrl };
    },
    { scope: ["User.Read"], teamsfx: teamsfx }
  );

  return (
    <div>
      {/* <Design /> */}
    <h3>External OAuth</h3>
    <Button primary content="External OAuth" onClick={() => { 
        // const params = {
        //   resouces: ["https://graph.microsoft.com/.default"],
        //   silent: true
        // };
        const mockOAuth = false;
        const hubAuthCallbackUrl = "ms-outlook://mos";
        authentication.authenticate({
          url: `auth_start.html?oauthRedirectMethod={oauthRedirectMethod}&authId=${mockOAuth ? "1" : "{authId}"}&mockOAuth=${mockOAuth}`,
          // url: `auth_start.html?oauthRedirectMethod={oauthRedirectMethod}&authId=${mockOAuth ? "1" : "{authId}"}&mockOAuth=${mockOAuth}&hubAuthCallbackUrl=${encodeURIComponent(hubAuthCallbackUrl)}`,
          isExternal: true,
          successCallback: function (result: any) {
            // output("Success:" + result);
            console.log("success");
          },
          failureCallback: function (reason: any) {
            // output("Failure:" + reason);
            console.log("failure");
          }
        });
        }} />

      <h3>Get the user's profile</h3>
      <div className="section-margin">
        <p>Click below to authorize button to grant permission to using Microsoft Graph. It will get user profile information</p>
        <pre>{`const teamsfx = new TeamsFx(); \nawait teamsfx.login(scope);`}</pre>
        <Button primary content="Authorize" disabled={loading} onClick={reload} />
        <p>Below are two different implementations of retrieving profile photo for currently signed-in user using Fluent UI component and Graph Toolkit respectively.</p>
        <h4>1. Display user profile using Fluent UI Component</h4>
        <PersonCardFluentUI loading={loading} data={data} error={error} />
        <h4>2. Display user profile using Graph Toolkit</h4>
        <PersonCardGraphToolkit loading={loading} data={data} error={error} />
      </div>
      <h3>Get the access token</h3>
      {/* <Button primary content="Access Token" onClick={() => { 
        // window.open('https://www.figma.com/community/file/916836509871353159', '_blank', 'noreferrer')
        // app.getContext().then((context) => {
        //   // setState({
        //   //   teamsContext: context
        //   // });
        // });
        // const params = {
        //   resouces: ["https://graph.microsoft.com/.default"],
        //   silent: true
        // };
        // authentication.getAuthToken(params).then((token) => {

        // });
         }} /> */}

         <Auth />

    </div>
  );
}

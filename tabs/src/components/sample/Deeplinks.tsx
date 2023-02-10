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
import { getDeepLinkTabStatic } from "./DeepLinkTabHelper";
import { pages } from "@microsoft/teams-js";


export function Deeplinks() {
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

  // if (pages.isSupported()) {
  //   const navPromise = pages.navigateToApp({ appId: , pageId: <pageId>, webUrl: <webUrl>, subPageId: <subPageId>, channelId:<channelId>});
  //   navPromise.
  //      then((result) => {/*Successful navigation*/}).
  //      catch((error) => {/*Failed navigation*/});
  // }
  // else { /* handle case where capability isn't supported */ }



  var botsDeepLink = getDeepLinkTabStatic("topic1", "1", "This is description", process.env.MicrosoftAppId)

  // const { } = useCalendar(
  //       // Open a scheduling dialog from your tab
  //   if(calendar.isSupported()) {
  //     const calendarPromise = calendar.composeMeeting({
  //       attendees: ["joe@contoso.com", "bob@contoso.com"],
  //       content: "test content",
  //       endTime: "2018-10-24T10:30:00-07:00",
  //       startTime: "2018-10-24T10:00:00-07:00",
  //       subject: "test subject"});
  //     calendarPromise.
  //       then((result) => {/*Successful operation*/}).
  //       catch((error) => {/*Unsuccessful operation*/});
  //   }
  //   else { /* handle case where capability isn't supported */ }
  // );

  return (
    <div>
      {/* <Design /> */}
      {/* <h3>Get the user's profile</h3>
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
      <h3>Get the access token</h3> */}

      <h2>navigateToApp</h2>
      <h3>Deep link to external link</h3>
      {/* <p>{botsDeepLink.linkUrl}</p> */}

      <Button primary content="external link" onClick={() => {
        // window.open(botsDeepLink.linkUrl, "_blank");
        // window.open("www.bing.com", "_blank");

        window.open('https://www.figma.com/community/file/916836509871353159', '_blank', 'noreferrer')


        // if (pages.isSupported()) {

        //   const navPromise = pages.navigateCrossDomain("www.bing.com");
        //   navPromise.then((result) => {
        //     console.log(result);
        //   }).catch((error) => {
        //     console.log(error);
        //   });
        // }
      }} />

      <h3>Deep link to a static tab</h3>
      <pre>{`
        /**
         * Parameters for the NavigateToApp API
         */
     interface NavigateToAppParams {
        /**
           * ID of the application to navigate to
           */
        appId: string;
        /**
           * Developer-defined ID of the Page to navigate to within the application (Formerly EntityID)
           */
        pageId: string;
        /**
           * Optional URL to open if the navigation cannot be completed within the host
           */
        webUrl?: string;
        /**
           * Optional developer-defined ID describing the content to navigate to within the Page. This will be passed
           * back to the application via the Context object.
           */
        subPageId?: string;
        /**
           * Optional ID of the Teams Channel where the application should be opened
           */
        channelId?: string;
     }
      `}</pre>

      <Button primary content="Navigate to tab DaysTo (Debug)" onClick={() => {
        if (pages.isSupported()) {
          const params = {
            appId: "103392a1-340c-4e5b-8449-40e1ae8f3a3a",
            pageId: "daysTo"
          }
          const navPromise = pages.navigateToApp(params);
          navPromise.then((result) => {
            console.log(result);
          }).catch((error) => {
            console.log(error);
          });
        }
      }} />
      <p></p>
      <Button primary content="Navigate to tab DaysTo (Dev)" onClick={() => {
        if (pages.isSupported()) {
          const params = {
            appId: "0b5dd87a-bdb9-4ee2-81d8-4dc365407fa2",
            pageId: "daysTo"
          }
          const navPromise = pages.navigateToApp(params);
          navPromise.then((result) => {
            console.log(result);
          }).catch((error) => {
            console.log(error);
          });
        }
      }} />
      <p></p>
      {/* <Button primary content="Navigate to Yammer" onClick={() => {
        if (pages.isSupported()) {
          const params = { p
            appId: "955070e9-99a6-4319-b8df-32adf59949aa",
            pageId: "1"
          }
          const navPromise = pages.navigateToApp(params);
          navPromise.then((result) => {
            alert(JSON.stringify({ result}));

            console.log(result);
          }).catch((error) => {
            alert(JSON.stringify({ error}));

            console.log(error);
          });
        }
      }} /> */}
            <Button primary content="Navigate to Top Stories" onClick={() => {
        if (pages.isSupported()) {
          const params = {
            appId: "1377aafa-2795-4e5e-8fa8-170b64e8a3e7",
            pageId: "index"
          }
          const navPromise = pages.navigateToApp(params);
          navPromise.then((result) => {
            alert(JSON.stringify({ result}));

            console.log(result);
          }).catch((error) => {
            alert(JSON.stringify({ error}));

            console.log(error);
          });
        }
      }} />


      <h2>shareDeepLink</h2>
      <pre>{`
interface ShareDeepLinkParameters {
  /**
      * The developer-defined unique ID for the sub-page to which this deep link points in the current page.
      * This field should be used to restore to a specific state within a page, such as scrolling to or activating a specific piece of content.
      */
  subPageId: string;
  /**
      * The label for the sub-page that should be displayed when the deep link is rendered in a client.
      */
  subPageLabel: string;
  /**
      * The fallback URL to which to navigate the user if the client cannot render the page.
      * This URL should lead directly to the sub-entity.
      */
  subPageWebUrl?: string;
}
      `}</pre>
      {/* <Button primary content="ShareDeepLink" onClick={() => {
        if (pages.isSupported()) {
          const params = {
            subPageId: "955070e9-99a6-4319-b8df-32adf59949aa",
            subPageLabel: "subPageLabel"
          }
          const navPromise = pages.shareDeepLink(params);
          navPromise.then((result) => {
            console.log(result);
          }).catch((error) => {
            console.log(error);
          });
        }
      }} /> */}

    </div>
  );
}


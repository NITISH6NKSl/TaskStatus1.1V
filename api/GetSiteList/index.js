/* This code sample provides a starter kit to implement server side logic for your Teams App in JavaScript,
 * refer to https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference for complete Azure Functions
 * developer guide.
 */

// Import polyfills for fetch required by msgraph-sdk-javascript.
require("isomorphic-fetch");
const teamsfxSdk = require("@microsoft/teamsfx");
const { Client } = require("@microsoft/microsoft-graph-client");
const {
  TokenCredentialAuthenticationProvider,
} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const config = require("../config");
const NewList = require("./NewList.json");

/**
 * This function handles requests from teamsfx client.
 * The HTTP request should contain an SSO token queried from Teams in the header.
 * Before trigger this function, teamsfx binding would process the SSO token and generate teamsfx configuration.
 *
 * This function initializes the teamsfx SDK with the configuration and calls these APIs:
 * - new OnBehalfOfUserCredential(ssoToken, authConfig)  - Construct OnBehalfOfUserCredential instance with the received SSO token and initialized configuration.
 * - getUserInfo() - Get the user's information from the received SSO token.
 *
 * The response contains multiple message blocks constructed into a JSON object, including:
 * - An echo of the request body.
 * - The display name encoded in the SSO token.
 * - Current user's Microsoft 365 profile if the user has consented.
 *
 * @param {Context} context - The Azure Functions context object.
 * @param {HttpRequest} req - The HTTP request.
 * @param {teamsfxContext} { [key: string]: any; } - The context generated by teamsfx binding.
 */
module.exports = async function (context, req, teamsfxContext) {
  context.log("HTTP trigger function processed a request.");

  // Initialize response.
  const res = {
    status: 200,
    body: {},
  };

  // Put an echo into response body.
  res.body.receivedHTTPRequestBody = req.body || "";

  // Prepare access token.
  const ssoToken = teamsfxContext["AccessToken"];
  if (!ssoToken) {
    return {
      status: 400,
      body: {
        error: "No access token was found in request header.",
      },
    };
  }

  // Construct TeamsFx using user identity.
  let credential;
  try {
    const authConfig = {
      authorityHost: config.authorityHost,
      tenantId: config.tenantId,
      clientId: config.clientId,
      clientSecret: config.clientSecret,
    };
    credential = new teamsfxSdk.OnBehalfOfUserCredential(ssoToken, authConfig);
  } catch (e) {
    context.log.error(e);
    return {
      status: 500,
      body: {
        error:
          "Failed to construct OnBehalfOfUserCredential using your ssoToken. " +
          "Ensure your function app is configured with the right Azure AD App registration.",
      },
    };
  }

  // Query user's information from the access token.
  try {
    const currentUser = await credential.getUserInfo();
    if (currentUser && currentUser.displayName) {
      res.body.userInfoMessage = `User display name is ${currentUser.displayName}.`;
    } else {
      res.body.userInfoMessage =
        "No user information was found in access token.";
    }
  } catch (e) {
    context.log.error(e);
    return {
      status: 400,
      body: {
        error: "Access token is invalid.",
      },
    };
  }

  // Create a graph client to access user's Microsoft 365 data after user has consented.
  try {
    // Create an instance of the TokenCredentialAuthenticationProvider by passing the tokenCredential instance and options to the constructor
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ["https://graph.microsoft.com/.default"],
    });

    // Initialize Graph client instance with authProvider
    const graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });
    console.log(
      "This is a context in backend of ????????????????",
      `/sites/${context.req.body.tanentUrl}:?$search="Teams_Site"&$select=sharepointIds`
    );

    // console.log("Login the body for list id", context.bindings.req.body);
    // const listId1 = context.req.body.listid1;
    // const listId2 = context.req.body.listid2;
    try{
    // console.log("This is to check json is comming???????", NewList[0].listSub);
    // console.log("This is site name the body", context.req.body.siteName);
    const profile = await graphClient
      .api(`/sites/${context.req.body.tanentUrl}:/sites/Teams_Site`)
      .get();
    // const siteId =
    //   res.body.data.graphClientMessage.value[0]?.sharepointIds.siteId;
    console.log("This is a api respone----->????>>???>?", profile);
    const siteId = profile.id.split(",");
    res.body.graphClientMessage =siteId[1];
    console.log("This is data of site id?????????", siteId[1]);
    if (profile.id) {
      const listIdMain = await graphClient
        .api(`/sites/${siteId[1]}/lists/${context.req.body?.listTodo}`)
        .get();
      res.body.listIdToDo = listIdMain.id;
      console.log("This is a data of list check_______", listIdMain);
      const listIdEntry = await graphClient
        .api(`/sites/${siteId[1]}/lists/${context.req.body?.listTaskEntry}`)
        .get();
      console.log("This is a sub list data id >>>>>>>>>>", listIdEntry);
      res.body.listIdToDoEntry = listIdEntry.id;
    }
    } catch(e) {
      console.log("This is for console e.message//////////////////////////////////////////////////////",e.message)
      if(e.message.includes("Requested site could not be found")){
        const teamsObj = {
          "template@odata.bind":
            "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
          displayName: "Teams_SiteFinalAPP",
          description: "My Sample Team’s Description",
        };
        console.log("This is else suite ----->>>>>>????");
        const createSite = await graphClient.api(`/teams`).post(teamsObj);
  
        console.log("This is a response in createSite?????????/", createSite);
        const siteCheckAgain = await graphClient
          .api(`/sites/${context.req.body.tanentUrl}:/sites/Teams_Site`)
          .get();
        const findSiteId = siteCheckAgain.id.split(",");
        console.log(
          "This is a site at atime of creating new site>>>>>>>>>><>>><>><<>",
          findSiteId
        );
        res.body.graphClientMessage=findSiteId[1]
  
        console.log(
          "This is a find site after creating a sit--------e",
          siteCheckAgain
        );
        const list1 = await graphClient
          .api(`/sites/${findSiteId[1]}/lists`)
          .post(NewList[0].listMain);
          res.body.listIdToDo=list1.id
        console.log("This is main list created sucsee??????", list1);
        const list2 = await graphClient
          .api(`/sites/${findSiteId[1]}/lists`)
          .post(NewList[0].listSub);
          res.body.listIdToDoEntry=list2.id
        console.log("This is second list created sucess>>>>>>>>>>>>>>>>>", list2);
      

      }
    }
     
  } catch (e) {
    console.log("can we get the error value------", e);
    context.log.error("This is the context catch error well------->>>>>", e);
    return {
      status: 500,
      body: {
        error:
          "Failed to retrieve user profile from Microsoft Graph. The application may not be authorized.",
      },
    };
  }

  return res;
};

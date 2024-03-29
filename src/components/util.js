import config from "./sample/lib/config";
import { createApiClient } from "@microsoft/teamsfx";
import { BearerTokenAuthProvider } from "@microsoft/teamsfx";

const callBackendFun = async (teamsUserCredential, obj,func) => {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post(func, obj);
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
};

const CallApiItems = async (teamsUserCredential, obj) => {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post("GetItemsCall", obj);

    return response.data.graphClientMessage;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
};
const callProfileApi = async (teamsUserCredential) => {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.get("userProfile");
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
};
const Notifiy = async (teamsUserCredential, sendActivity) => {
  try {
    const functionRes = await callBackendFun(teamsUserCredential, sendActivity,"GetNotifictaion");
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const GetItems = async (teamsUserCredential, Obj) => {
  try {
    const functionRes = await CallApiItems(teamsUserCredential, Obj);
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const Update = async (teamsUserCredential, obj) => {
  try {
    const functionRes = await callBackendFun(teamsUserCredential, obj,"fieldSet");

    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};

const addTasklist = async (teamsUserCredential, obj) => {
  try {
    const functionRes = await callBackendFun(teamsUserCredential, obj,"addTask");

    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const RemoveTask = async (teamsUserCredential, obj) => {
  try {
    const functionRes = await callBackendFun(teamsUserCredential, obj,"DeleteTask");
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const playPause = async (teamsUserCredential, obj) => {
  try {
    const functionRes = await callBackendFun(teamsUserCredential, obj,"playPause");

    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const GetSite = async (teamsUserCredential, obj) => {
  try {
    const functionRes = await callBackendFun(teamsUserCredential, obj,"GetSiteList");
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const getprofile = async (teamsUserCredential) => {
  try {
    const functionRes = await callProfileApi(teamsUserCredential);
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};
const getLogData=async (teamsUserCredential,obj)=>{
  try {
    const functionRes = await callBackendFun(teamsUserCredential,obj,"logTimes");
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }

}
///////Try to  Get Data from call Ongoing tab/////


const callGetDataApi = async (teamsUserCredential,obj) => {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post("getData",obj);
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
};
const getListDataCall = async (teamsUserCredential,obj) => {
  try {
    const functionRes = await callGetDataApi(teamsUserCredential,obj);
    return functionRes;
  } catch (error) {
    if (error.message.includes("The application may not be authorized.")) {
    }
  }
};


export { Update, addTasklist, playPause, Notifiy, GetSite, GetItems,getprofile ,RemoveTask,getListDataCall,getLogData};

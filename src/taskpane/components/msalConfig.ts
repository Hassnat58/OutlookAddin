import { Configuration } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: "edc31410-d406-476a-8ac1-b062e1df4a77", // Replace with your App Registration client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://5jd7y6.sharepoint.com/sites/HScommunication/Outlook",
  },
  cache: {
    cacheLocation: "localStorage",
  },
};

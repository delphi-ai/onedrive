const msalParams = {
  auth: {
    authority:
      "https://login.microsoftonline.com/452dc845-891d-4630-a9b3-a3146f06f38e",
    clientId: "1a541581-bada-4818-8149-b3ab0b544ed3",
    redirectUri: "https://site-production-4e08.up.railway.app",
  },
};

const app = new msal.PublicClientApplication(msalParams);

async function getToken() {
  let accessToken = "";

  const authParams = { scopes: ["Files.ReadWrite.All", "User.Read"] };

  try {
    const accounts = app.getAllAccounts();
    console.log(accounts);
    if (accounts.length === 0) {
      // No accounts detected, user must login
      try {
        const loginResponse = await app.loginPopup(authParams);
        app.setActiveAccount(loginResponse.account);
      } catch (err) {
        console.error(err); // Handle or log errors from loginPopup
        return null; // or handle this appropriately
      }
    }
    // see if we have already the idtoken saved
    const resp = await app.acquireTokenSilent(authParams);
    accessToken = resp.accessToken;
  } catch (e) {
    console.error(e); // Log the error for debugging purposes
    if (e instanceof msal.InteractionRequiredAuthError) {
      // If interaction is required, user must authenticate with a popup or redirect
      try {
        const loginResponse = await app.loginPopup(authParams);
        app.setActiveAccount(loginResponse.account);
        const resp = await app.acquireTokenSilent(authParams);
        accessToken = resp.accessToken;
      } catch (err) {
        console.error(err); // Handle or log any errors from the popup method
      }
    }
  }

  return accessToken;
}

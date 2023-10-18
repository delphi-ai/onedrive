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

  authParams = { scopes: ["Files.ReadWrite.All"] };

  try {
    // see if we have already the idtoken saved
    const resp = await app.acquireTokenSilent(authParams);
    accessToken = resp.accessToken;
  } catch (e) {
    console.log(e);
    // per examples we fall back to popup
    const resp = await app.loginPopup(authParams);
    app.setActiveAccount(resp.account);

    if (resp.idToken) {
      const resp2 = await app.acquireTokenSilent(authParams);
      accessToken = resp2.accessToken;
    }
  }

  return accessToken;
}

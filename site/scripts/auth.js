const msalParams = {
  auth: {
    authority: "https://login.microsoftonline.com/common",
    clientId: "7d4ac4d9-1861-4102-9323-982fc0815db5",
    redirectUri: "https://onedrive-production.up.railway.app",
  },
};

const app = new msal.PublicClientApplication(msalParams);

async function getToken() {
  let accessToken = "";

  const authParams = { scopes: ["Files.ReadWrite", "Files.ReadWrite.All"] };

  try {
    // see if we have already the idtoken saved
    const resp = await app.acquireTokenSilent(authParams);
    accessToken = resp.accessToken;
  } catch (e) {
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

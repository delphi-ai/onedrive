const msalParams = {
  auth: {
    authority: "https://login.microsoftonline.com/consumers",
    clientId: "f3328d55-d849-448e-8b62-beb9857f5c1e",
    redirectUri: "https://onedrive-production.up.railway.app",
  },
};

const app = new msal.PublicClientApplication(msalParams);

async function getToken() {
  let accessToken = "";

  const authParams = { scopes: ["OneDrive.ReadWrite"] };

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

const msalParams = {
  auth: {
    authority: "https://login.microsoftonline.com/consumers",
    clientId: "439d927e-1e3d-418d-912b-da1364f5a46c",
    redirectUri: "https://site-production-4e08.up.railway.app",
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

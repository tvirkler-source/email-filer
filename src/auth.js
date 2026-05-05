const Auth = (() => {
  const BACKEND_URL = 'https://emailfiler.tvirkler.workers.dev/';
  const CLIENT_ID   = '3eb85312-1c24-4c02-aca9-14cd706709c2';
  const TENANT_ID   = '32655b6b-2430-4872-bb66-28dbd84ee0c4';

  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      redirectUri: 'https://email-filer.pages.dev/redirect.html',
    },
    cache: { cacheLocation: 'sessionStorage' }
  };

  const scopes = [
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Files.ReadWrite.All',
    'openid', 'profile'
  ];

  let _msalInstance = null;
  let _account = null;

  function getInstance() {
    if (!_msalInstance) {
      _msalInstance = new msal.PublicClientApplication(msalConfig);
    }
    return _msalInstance;
  }

  async function getToken() {
    const app = getInstance();
    await app.initialize();

    // Handle redirect response first
    await app.handleRedirectPromise();

    const accounts = app.getAllAccounts();
    if (accounts.length > 0) {
      _account = accounts[0];
    }

    // Try silent first
    if (_account) {
      try {
        const result = await app.acquireTokenSilent({
          scopes, account: _account
        });
        return result.accessToken;
      } catch (e) {
        // Silent failed, fall through to popup
      }
    }

    // Popup
    const result = await app.acquireTokenPopup({ scopes });
    _account = result.account;
    return result.accessToken;
  }

  function clearToken() {
    _account = null;
    if (_msalInstance) {
      const accounts = _msalInstance.getAllAccounts();
      accounts.forEach(a => _msalInstance.removeAccount(a));
    }
  }

  return { getToken, clearToken, BACKEND_URL };
})();

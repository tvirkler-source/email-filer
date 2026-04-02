// auth.js — MSAL + Office SSO with On-Behalf-Of handoff to backend
// The add-in never calls Graph directly; it passes the Office identity token to the backend.

const Auth = (() => {
  const BACKEND_URL = 'https://email-filer.pages.dev/'; // set at deploy time

  let _accessToken = null;
  let _tokenExpiry = 0;

  /**
   * Get a valid backend access token.
   * Uses Office SSO (getAccessToken) which returns a token scoped to the add-in.
   * The backend validates and exchanges this via OBO for a Graph token.
   */
  async function getToken() {
    const now = Date.now();

    // Return cached token if still valid (with 60s buffer)
    if (_accessToken && _tokenExpiry - 60_000 > now) {
      return _accessToken;
    }

    try {
      // Office SSO — this is a bootstrap token, NOT a Graph token
      const bootstrapToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });

      // Exchange with backend — backend does the OBO flow
      const resp = await fetch(`${BACKEND_URL}/api/auth/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ bootstrapToken }),
      });

      if (!resp.ok) throw new Error(`Token exchange failed: ${resp.status}`);

      const data = await resp.json();
      _accessToken = data.accessToken;
      _tokenExpiry = now + data.expiresInMs;

      return _accessToken;
    } catch (err) {
      console.error('Auth error:', err);
      _accessToken = null;
      throw err;
    }
  }

  function clearToken() {
    _accessToken = null;
    _tokenExpiry = 0;
  }

  return { getToken, clearToken, BACKEND_URL };
})();

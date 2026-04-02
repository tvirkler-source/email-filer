// api.js — all calls to the EmailFiler backend

const API = (() => {

  async function _request(path, options = {}) {
    const token = await Auth.getToken();
    const url = `${Auth.BACKEND_URL}${path}`;

    const resp = await fetch(url, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`,
        ...(options.headers || {}),
      },
    });

    if (resp.status === 401) {
      Auth.clearToken();
      throw new Error('Unauthorized — please refresh');
    }

    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`API error ${resp.status}: ${text}`);
    }

    return resp.json();
  }

  /** Get folder suggestions for the current email */
  async function getSuggestions(subject, bodySnippet, senderEmail) {
    return _request('/api/folders/suggest', {
      method: 'POST',
      body: JSON.stringify({ subject, bodySnippet, senderEmail }),
    });
  }

  /** Typeahead search */
  async function searchFolders(query) {
    return _request(`/api/folders/search?q=${encodeURIComponent(query)}`);
  }

  /** Save email to a folder */
  async function saveEmail(folderDriveId, folderItemId, emailData) {
    return _request('/api/email/save', {
      method: 'POST',
      body: JSON.stringify({ folderDriveId, folderItemId, ...emailData }),
    });
  }

  /** Load user preferences (pinned/recent) */
  async function getPreferences() {
    return _request('/api/user/preferences');
  }

  /** Save user preferences */
  async function savePreferences(prefs) {
    return _request('/api/user/preferences', {
      method: 'PUT',
      body: JSON.stringify(prefs),
    });
  }

  return { getSuggestions, searchFolders, saveEmail, getPreferences, savePreferences };
})();

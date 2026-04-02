// folderCache.js
// Caches folder search results in sessionStorage to avoid hammering the API on each keystroke.

const FolderCache = (() => {
  const CACHE_KEY = 'ef_folder_cache';
  const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

  function _load() {
    try {
      const raw = sessionStorage.getItem(CACHE_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (Date.now() > parsed.expiry) { sessionStorage.removeItem(CACHE_KEY); return null; }
      return parsed.folders;
    } catch { return null; }
  }

  function _save(folders) {
    try {
      sessionStorage.setItem(CACHE_KEY, JSON.stringify({
        folders,
        expiry: Date.now() + CACHE_TTL_MS,
      }));
    } catch { /* storage full — ignore */ }
  }

  /**
   * Filter cached folders by query string.
   * Returns null if cache is empty (caller should fall back to API).
   */
  function filterCached(query) {
    const folders = _load();
    if (!folders) return null;
    const q = query.toLowerCase();
    return folders.filter(f =>
      f.name.toLowerCase().includes(q) ||
      f.path.toLowerCase().includes(q)
    );
  }

  function populate(folders) {
    _save(folders);
  }

  function invalidate() {
    sessionStorage.removeItem(CACHE_KEY);
  }

  return { filterCached, populate, invalidate };
})();

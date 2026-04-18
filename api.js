// Backend API: talks to Apps Script if configured, else falls back to localStorage.
const API = (() => {
  const LS_URL_KEY = 'docxviewer.apiUrl';
  const LS_DATA_KEY = 'docxviewer.localData';

  const getUrl = () => localStorage.getItem(LS_URL_KEY) || '';
  const setUrl = (u) => localStorage.setItem(LS_URL_KEY, u || '');

  const emptyStore = () => ({
    documents: [], paragraphs: [], comments: [], tags: [], comment_tags: []
  });

  const loadLocal = () => {
    try { return JSON.parse(localStorage.getItem(LS_DATA_KEY)) || emptyStore(); }
    catch { return emptyStore(); }
  };
  const saveLocal = (d) => localStorage.setItem(LS_DATA_KEY, JSON.stringify(d));

  async function post(action, payload) {
    const url = getUrl();
    if (!url) return localHandler(action, payload);
    const res = await fetch(url, {
      method: 'POST',
      // text/plain avoids CORS preflight with Apps Script
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, payload })
    });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'Error en la API');
    return json.data;
  }

  async function loadAll() {
    const url = getUrl();
    if (!url) return loadLocal();
    const res = await fetch(url, { method: 'GET' });
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'Error cargando datos');
    return json.data;
  }

  // --- Local fallback (same API as remote) ---
  function localHandler(action, payload) {
    const d = loadLocal();
    const find  = (arr, key, val) => arr.find(r => String(r[key]) === String(val));
    const without = (arr, pred) => arr.filter(r => !pred(r));

    switch (action) {
      case 'saveDocument': {
        const { document: doc, paragraphs, comments } = payload;
        d.documents = without(d.documents, r => r.doc_id === doc.doc_id);
        d.paragraphs = without(d.paragraphs, r => r.doc_id === doc.doc_id);
        d.comments   = without(d.comments,   r => r.doc_id === doc.doc_id);
        d.documents.push(doc);
        d.paragraphs.push(...paragraphs);
        d.comments.push(...comments);
        saveLocal(d); return { saved: true };
      }
      case 'deleteDocument': {
        const { doc_id } = payload;
        const cids = d.comments.filter(c => c.doc_id === doc_id).map(c => String(c.comment_id));
        d.comment_tags = without(d.comment_tags, r => cids.includes(String(r.comment_id)));
        d.comments     = without(d.comments,     r => r.doc_id === doc_id);
        d.paragraphs   = without(d.paragraphs,   r => r.doc_id === doc_id);
        d.documents    = without(d.documents,    r => r.doc_id === doc_id);
        saveLocal(d); return { deleted: true };
      }
      case 'updateDocument': {
        const { doc_id, patch } = payload;
        const doc = find(d.documents, 'doc_id', doc_id);
        if (doc) Object.assign(doc, patch);
        saveLocal(d); return { updated: true };
      }
      case 'saveTag':   { d.tags.push(payload); saveLocal(d); return { saved: true }; }
      case 'updateTag': {
        const { tag_id, patch } = payload;
        const t = find(d.tags, 'tag_id', tag_id);
        if (t) Object.assign(t, patch);
        saveLocal(d); return { updated: true };
      }
      case 'deleteTag': {
        const { tag_id } = payload;
        d.comment_tags = without(d.comment_tags, r => String(r.tag_id) === String(tag_id));
        d.tags         = without(d.tags,         r => String(r.tag_id) === String(tag_id));
        saveLocal(d); return { deleted: true };
      }
      case 'setCommentTags': {
        const { comment_id, tag_ids } = payload;
        d.comment_tags = without(d.comment_tags, r => String(r.comment_id) === String(comment_id));
        d.comment_tags.push(...tag_ids.map(tid => ({ comment_id, tag_id: tid })));
        saveLocal(d); return { updated: true };
      }
    }
    return null;
  }

  return { getUrl, setUrl, loadAll, post };
})();

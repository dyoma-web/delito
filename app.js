// Main application logic.
const DEFAULT_COLORS = ['#FFE4B5', '#B5E4FF', '#D4FFB5', '#FFB5D4', '#E4B5FF', '#FFFFB5', '#B5FFD4'];
let nextDefaultColor = 0;

// ---------- URL params: view mode for counterparts ----------
const _params = new URLSearchParams(location.search);
const VIEW_MODE = _params.get('view') === '1';
const VIEW_HIDE = new Set((_params.get('hide') || '').split(',').filter(Boolean));
const VIEW_HIDE_DOCS = new Set((_params.get('hidedocs') || '').split(',').filter(Boolean));

if (VIEW_MODE) {
  document.body.classList.add('view-mode');
  for (const h of VIEW_HIDE) document.body.classList.add('hide-' + h);
}

const state = {
  documents: [], paragraphs: [], comments: [], tags: [], comment_tags: [],
  filters: { text: '', author: '', tagId: '', showResolved: false },
  contextExpansion: {},
  pendingCtagCommentId: null
};

// ---------- Bootstrap ----------
document.addEventListener('DOMContentLoaded', async () => {
  wireEvents();
  await reloadAll();
});

async function reloadAll() {
  showLoading('Cargando datos…');
  try {
    const data = await API.loadAll();
    Object.assign(state, data);
    render();
  } catch (e) {
    toast('Error cargando datos: ' + e.message, true);
  } finally {
    hideLoading();
  }
}

// ---------- Events ----------
function wireEvents() {
  document.getElementById('file-input').addEventListener('change', onFileUpload);
  document.getElementById('btn-tags').addEventListener('click', openTagModal);
  document.getElementById('btn-config').addEventListener('click', openConfigModal);
  document.getElementById('btn-export').addEventListener('click', exportCSV);
  document.getElementById('btn-backup').addEventListener('click', exportBackup);
  document.getElementById('btn-share').addEventListener('click', openShareModal);
  document.getElementById('btn-pdf-view').addEventListener('click', () => window.print());
  document.getElementById('btn-share-copy').addEventListener('click', copyShareURL);
  document.getElementById('btn-share-open').addEventListener('click', openShareURL);
  document.getElementById('btn-share-pdf').addEventListener('click', generateSharePDF);
  document.getElementById('btn-reload').addEventListener('click', reloadAll);

  document.getElementById('filter-text').addEventListener('input', e => { state.filters.text = e.target.value.toLowerCase(); renderDocs(); });
  document.getElementById('filter-author').addEventListener('change', e => { state.filters.author = e.target.value; renderDocs(); });
  document.getElementById('filter-tag').addEventListener('change', e => { state.filters.tagId = e.target.value; renderDocs(); });
  document.getElementById('filter-resolved').addEventListener('change', e => { state.filters.showResolved = e.target.checked; renderDocs(); });

  document.getElementById('btn-tag-add').addEventListener('click', onTagAdd);
  document.getElementById('btn-cfg-save').addEventListener('click', onConfigSave);
  document.getElementById('btn-ctag-save').addEventListener('click', onCtagSave);

  document.querySelectorAll('[data-close]').forEach(btn => btn.addEventListener('click', e => {
    e.target.closest('.modal').hidden = true;
  }));
  document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', e => {
    if (e.target === m) m.hidden = true;
  }));
}

// ---------- Loader ----------
function showLoading(msg) {
  const el = document.getElementById('loader-overlay');
  if (!el) return;
  el.querySelector('.loader-text').textContent = msg || 'Cargando…';
  el.hidden = false;
}
function hideLoading() {
  const el = document.getElementById('loader-overlay');
  if (el) el.hidden = true;
}
function setLoaderText(msg) {
  const el = document.getElementById('loader-overlay');
  if (el && !el.hidden) el.querySelector('.loader-text').textContent = msg;
}

// ---------- File upload ----------
async function onFileUpload(e) {
  const files = Array.from(e.target.files);
  e.target.value = '';
  showLoading('Procesando archivo…');
  try {
    for (const f of files) {
      try { await handleFile(f); }
      catch (err) { toast(`Error con ${f.name}: ${err.message}`, true); }
    }
    render();
  } finally {
    hideLoading();
  }
}

async function handleFile(file) {
  setLoaderText(`Parseando ${file.name}…`);
  const parsed = await parseDocx(file);
  const existing = state.documents.find(d => d.doc_id === parsed.docId);

  let docId = parsed.docId;
  if (existing) {
    const choice = confirm(
      `"${file.name}" ya está cargado como "${existing.filename}".\n\n` +
      `¿Reemplazar con esta versión? (Cancelar = duplicar como copia nueva)`
    );
    if (choice) {
      // keep same doc_id, overwrite
    } else {
      docId = parsed.docId + '_' + Date.now();
    }
  }

  const color = existing && existing.doc_id === docId
    ? existing.color
    : DEFAULT_COLORS[nextDefaultColor++ % DEFAULT_COLORS.length];

  const doc = {
    doc_id: docId,
    filename: file.name,
    color,
    uploaded_at: new Date().toISOString(),
    visible: true
  };
  const paragraphs = parsed.paragraphs.map(p => ({
    doc_id: docId,
    paragraph_index: p.paragraph_index,
    text: p.text,
    page_number: p.page_number,
    page_approx: p.page_approx
  }));
  const comments = parsed.comments.map(c => ({
    comment_id: `${docId}_${c.id}`,
    doc_id: docId,
    paragraph_index: c.paragraph_index,
    comment_text: c.text,
    author: c.author,
    created_at: c.date,
    parent_comment_id: c.parentId ? `${docId}_${c.parentId}` : '',
    resolved: c.resolved ? 1 : 0,
    observation: ''
  }));

  setLoaderText(`Guardando ${comments.length} comentarios…`);
  await API.post('saveDocument', { document: doc, paragraphs, comments });

  state.documents   = state.documents.filter(d => d.doc_id !== docId).concat([doc]);
  state.paragraphs  = state.paragraphs.filter(p => p.doc_id !== docId).concat(paragraphs);
  state.comments    = state.comments.filter(c => c.doc_id !== docId).concat(comments);

  toast(`${file.name}: ${comments.length} comentarios importados.`);
}

// ---------- Render ----------
function render() {
  renderFilters();
  renderDocToggles();
  renderDocs();
}

function renderFilters() {
  const authors = [...new Set(state.comments.map(c => c.author).filter(Boolean))].sort();
  const authorSel = document.getElementById('filter-author');
  const currentA = authorSel.value;
  authorSel.innerHTML = '<option value="">Todos los autores</option>' +
    authors.map(a => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join('');
  authorSel.value = currentA;

  const tagSel = document.getElementById('filter-tag');
  const currentT = tagSel.value;
  tagSel.innerHTML = '<option value="">Todas las etiquetas</option>' +
    state.tags.map(t => `<option value="${t.tag_id}">${escapeHtml(t.name)}</option>`).join('');
  tagSel.value = currentT;
}

function renderDocToggles() {
  const c = document.getElementById('doc-toggles');
  c.innerHTML = '';
  for (const d of state.documents) {
    const el = document.createElement('span');
    el.className = 'doc-toggle' + (isVisible(d) ? '' : ' hidden-doc');
    el.style.background = d.color;
    el.style.borderColor = d.color;
    el.textContent = shortName(d.filename);
    el.title = 'Clic para mostrar/ocultar: ' + d.filename;
    el.addEventListener('click', () => toggleDocVisibility(d));
    c.appendChild(el);
  }
}

function isVisible(d) {
  return d.visible === true || d.visible === 'TRUE' || d.visible === 1 || d.visible === '1' || d.visible === undefined;
}

async function toggleDocVisibility(d) {
  const newVal = !isVisible(d);
  d.visible = newVal;
  try {
    await API.post('updateDocument', { doc_id: d.doc_id, patch: { visible: newVal } });
  } catch (e) { toast('Error: ' + e.message, true); }
  render();
}

function renderDocs() {
  const container = document.getElementById('docs-container');
  container.innerHTML = '';
  let visibleDocs = state.documents.filter(isVisible);
  if (VIEW_MODE && VIEW_HIDE_DOCS.size) {
    visibleDocs = visibleDocs.filter(d => !VIEW_HIDE_DOCS.has(d.doc_id));
  }
  if (!visibleDocs.length) {
    container.innerHTML = `<div class="empty-state" id="empty-state">
      <p>No hay documentos cargados o todos están ocultos.</p>
      <p>Sube un <code>.docx</code> con comentarios para empezar.</p>
    </div>`;
    return;
  }
  for (const d of visibleDocs) container.appendChild(renderDoc(d));
}

function renderDoc(d) {
  const card = document.createElement('div');
  card.className = 'doc-card';
  card.dataset.docId = d.doc_id;
  card.style.background = d.color;

  const header = document.createElement('div');
  header.className = 'doc-header';
  header.innerHTML = `
    <span class="title" title="${escapeAttr(d.filename)}">${escapeHtml(d.filename)}</span>
    <input type="color" value="${d.color}" title="Color de fondo">
    <button class="btn small" data-act="hide">Ocultar</button>
    <button class="btn small danger" data-act="delete">Eliminar</button>
  `;
  header.querySelector('input[type=color]').addEventListener('input', async e => {
    d.color = e.target.value;
    card.style.background = d.color;
    try { await API.post('updateDocument', { doc_id: d.doc_id, patch: { color: d.color } }); }
    catch (er) { toast('Error: ' + er.message, true); }
    renderDocToggles();
  });
  header.querySelector('[data-act=hide]').addEventListener('click', () => toggleDocVisibility(d));
  header.querySelector('[data-act=delete]').addEventListener('click', () => deleteDoc(d));
  card.appendChild(header);

  const body = document.createElement('div');
  body.className = 'doc-body';
  body.appendChild(renderDocTable(d));
  card.appendChild(body);
  return card;
}

function renderDocTable(d) {
  const paragraphs = state.paragraphs
    .filter(p => p.doc_id === d.doc_id)
    .sort((a, b) => Number(a.paragraph_index) - Number(b.paragraph_index));

  const paraMap = {};
  paragraphs.forEach(p => { paraMap[Number(p.paragraph_index)] = p; });

  const docComments = state.comments.filter(c => c.doc_id === d.doc_id);
  const topLevel    = docComments.filter(c => !c.parent_comment_id);
  const repliesByParent = {};
  for (const c of docComments) {
    if (c.parent_comment_id) {
      (repliesByParent[c.parent_comment_id] ||= []).push(c);
    }
  }

  const filtered = topLevel.filter(c => matchesFilters(c, repliesByParent[c.comment_id] || []));
  const byPara = {};
  for (const c of filtered) {
    const idx = Number(c.paragraph_index);
    (byPara[idx] ||= []).push(c);
  }
  const paraIndexes = Object.keys(byPara).map(Number).sort((a, b) => a - b);

  const table = document.createElement('table');
  table.className = 'doc-table';
  table.innerHTML = `<thead><tr>
    <th class="col-page">Pg</th>
    <th class="col-para">Párrafo</th>
    <th class="col-comment">Comentario</th>
    <th class="col-obs">Observación</th>
  </tr></thead>`;
  const tbody = document.createElement('tbody');

  if (paraIndexes.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:#64748b;padding:24px">
      ${docComments.length ? 'Sin resultados con los filtros actuales.' : 'Este documento no tiene comentarios.'}
    </td></tr>`;
  }

  for (const idx of paraIndexes) {
    const comments = byPara[idx];
    const p = paraMap[idx];
    const pageNum = p ? p.page_number : '';
    const approx  = p && (Number(p.page_approx) === 1 || p.page_approx === 'TRUE' || p.page_approx === true);
    const rowSpan = comments.length;

    comments.forEach((c, ci) => {
      const tr = document.createElement('tr');

      if (ci === 0) {
        const tdPage = document.createElement('td');
        tdPage.className = 'col-page cell-page';
        tdPage.rowSpan = rowSpan;
        tdPage.appendChild(renderPageBadge(pageNum, approx));
        tr.appendChild(tdPage);

        const tdP = document.createElement('td');
        tdP.className = 'col-para cell-paragraph';
        tdP.rowSpan = rowSpan;
        tdP.appendChild(renderParagraphCell(d, idx, paraMap));
        tr.appendChild(tdP);
      }

      const tdC = document.createElement('td');
      tdC.className = 'col-comment cell-comment';
      tdC.appendChild(renderCommentNode(c, repliesByParent[c.comment_id] || [], comments.length > 1, ci + 1, comments.length));
      tr.appendChild(tdC);

      const tdO = document.createElement('td');
      tdO.className = 'col-obs cell-observation';
      tdO.appendChild(renderObservation(c));
      tr.appendChild(tdO);

      tbody.appendChild(tr);
    });
  }

  table.appendChild(tbody);
  return table;
}

function renderPageBadge(pg, approx) {
  const span = document.createElement('span');
  span.className = 'page-badge' + (approx ? ' page-approx' : '');
  span.textContent = pg === '' || pg === undefined || pg === null
    ? '—'
    : (approx ? `~${pg}` : String(pg));
  span.title = approx
    ? 'Página estimada (el documento no trae marcas de paginado de Word)'
    : 'Página';
  return span;
}

function renderParagraphCell(d, idx, paraMap) {
  const wrap = document.createElement('div');
  const expKey = `${d.doc_id}:${idx}`;
  const exp = state.contextExpansion[expKey] || { before: 0, after: 0 };
  const mainObj = paraMap[idx];
  const mainText = mainObj ? mainObj.text : '(párrafo no encontrado)';

  for (let i = exp.before; i > 0; i--) {
    const t = paraMap[idx - i];
    if (t) {
      const div = document.createElement('div');
      div.className = 'context-paragraph';
      div.textContent = t.text;
      wrap.appendChild(div);
    }
  }
  const main = document.createElement('div');
  main.textContent = mainText;
  wrap.appendChild(main);
  for (let i = 1; i <= exp.after; i++) {
    const t = paraMap[idx + i];
    if (t) {
      const div = document.createElement('div');
      div.className = 'context-paragraph';
      div.textContent = t.text;
      wrap.appendChild(div);
    }
  }

  const btns = document.createElement('div');
  btns.className = 'expand-btns';
  btns.innerHTML = `
    <button data-act="exp-bef--">− antes</button>
    <button data-act="exp-bef++">+ antes</button>
    <button data-act="exp-aft--">− después</button>
    <button data-act="exp-aft++">+ después</button>
  `;
  btns.addEventListener('click', (e) => {
    const act = e.target.dataset.act;
    if (!act) return;
    const cur = { ...(state.contextExpansion[expKey] || { before: 0, after: 0 }) };
    if (act === 'exp-bef++') cur.before = Math.min(cur.before + 1, idx);
    if (act === 'exp-bef--') cur.before = Math.max(0, cur.before - 1);
    if (act === 'exp-aft++') cur.after  = cur.after + 1;
    if (act === 'exp-aft--') cur.after  = Math.max(0, cur.after - 1);
    state.contextExpansion[expKey] = cur;
    renderDocs();
  });
  wrap.appendChild(btns);
  return wrap;
}

function renderCommentNode(c, replies, isMulti, n, total) {
  const node = document.createElement('div');
  node.className = 'comment-root';

  const meta = document.createElement('div');
  meta.className = 'comment-meta';
  const left = document.createElement('span');
  left.innerHTML = `<span class="comment-author">${escapeHtml(c.author || '')}</span>` +
                   (c.created_at ? `<span class="comment-date"> · ${formatDate(c.created_at)}</span>` : '') +
                   (isMulti ? `<span class="multi-comment-note">Comentario ${n}/${total} del mismo párrafo</span>` : '');
  const right = document.createElement('span');
  if (Number(c.resolved) === 1) right.innerHTML = `<span class="resolved-badge">Resuelto</span>`;
  meta.appendChild(left); meta.appendChild(right);
  node.appendChild(meta);

  const text = document.createElement('div');
  text.className = 'comment-text';
  text.textContent = c.comment_text || '';
  node.appendChild(text);

  for (const r of replies.sort((a, b) => String(a.created_at).localeCompare(String(b.created_at)))) {
    const rep = document.createElement('div');
    rep.className = 'comment-reply';
    const rmeta = document.createElement('div');
    rmeta.className = 'comment-meta';
    rmeta.innerHTML = `<span class="comment-author">↳ ${escapeHtml(r.author || '')}</span>` +
                      (r.created_at ? `<span class="comment-date"> · ${formatDate(r.created_at)}</span>` : '');
    const rtxt = document.createElement('div');
    rtxt.className = 'comment-text';
    rtxt.textContent = r.comment_text || '';
    rep.appendChild(rmeta); rep.appendChild(rtxt);
    node.appendChild(rep);
  }

  const badges = document.createElement('div');
  badges.className = 'comment-badges';
  const tagsOfC = tagsOfComment(c.comment_id);
  for (const t of tagsOfC) {
    const chip = document.createElement('span');
    chip.className = 'tag-chip';
    chip.style.background = t.color;
    chip.textContent = t.name;
    badges.appendChild(chip);
  }
  const btn = document.createElement('button');
  btn.className = 'btn small';
  btn.textContent = tagsOfC.length ? 'Editar etiquetas' : '+ Etiqueta';
  btn.addEventListener('click', () => openCtagModal(c.comment_id));
  badges.appendChild(btn);
  node.appendChild(badges);

  return node;
}

function renderObservation(c) {
  const wrap = document.createElement('div');
  wrap.className = 'obs-wrap';

  const ta = document.createElement('textarea');
  ta.className = 'obs-input';
  ta.placeholder = 'Añadir observación…';
  ta.maxLength = 2000;
  ta.value = c.observation || '';
  ta.rows = 3;

  const del = document.createElement('button');
  del.className = 'btn small danger obs-del';
  del.textContent = 'Eliminar';
  del.style.display = (c.observation && String(c.observation).length > 0) ? 'inline-block' : 'none';

  let savedValue = ta.value;
  const save = async (newVal) => {
    const v = newVal !== undefined ? newVal : ta.value;
    if (v === savedValue) return;
    savedValue = v;
    c.observation = v;
    try {
      await API.post('updateCommentObservation', { comment_id: c.comment_id, observation: v });
    } catch (e) {
      toast('Error guardando observación: ' + e.message, true);
    }
  };
  ta.addEventListener('blur', () => save());
  ta.addEventListener('input', () => {
    del.style.display = ta.value ? 'inline-block' : 'none';
  });
  wrap.appendChild(ta);

  del.addEventListener('click', async () => {
    if (!confirm('¿Eliminar la observación? Esta acción no se puede deshacer.')) return;
    ta.value = '';
    del.style.display = 'none';
    await save('');
  });
  wrap.appendChild(del);

  return wrap;
}

// ---------- Filters ----------
function matchesFilters(c, replies) {
  const f = state.filters;
  if (Number(c.resolved) === 1 && !f.showResolved) return false;

  if (f.author && c.author !== f.author && !replies.some(r => r.author === f.author)) return false;

  if (f.tagId) {
    const ids = tagIdsOfComment(c.comment_id);
    if (!ids.includes(f.tagId) && !replies.some(r => tagIdsOfComment(r.comment_id).includes(f.tagId))) return false;
  }

  if (f.text) {
    const hay = [
      c.comment_text, c.author, c.observation,
      ...replies.map(r => r.comment_text), ...replies.map(r => r.author),
      getParagraphText(c.doc_id, c.paragraph_index)
    ].join(' ').toLowerCase();
    if (!hay.includes(f.text)) return false;
  }
  return true;
}

function getParagraphText(docId, idx) {
  const p = state.paragraphs.find(p => p.doc_id === docId && Number(p.paragraph_index) === Number(idx));
  return p ? p.text : '';
}

function tagIdsOfComment(cid) {
  return state.comment_tags
    .filter(r => String(r.comment_id) === String(cid))
    .map(r => String(r.tag_id));
}
function tagsOfComment(cid) {
  const ids = tagIdsOfComment(cid);
  return state.tags.filter(t => ids.includes(String(t.tag_id)));
}

// ---------- Delete doc ----------
async function deleteDoc(d) {
  if (!confirm(`¿Eliminar "${d.filename}" y todos sus comentarios?`)) return;
  showLoading('Eliminando documento…');
  try {
    await API.post('deleteDocument', { doc_id: d.doc_id });
    const cids = state.comments.filter(c => c.doc_id === d.doc_id).map(c => String(c.comment_id));
    state.comment_tags = state.comment_tags.filter(r => !cids.includes(String(r.comment_id)));
    state.comments     = state.comments.filter(c => c.doc_id !== d.doc_id);
    state.paragraphs   = state.paragraphs.filter(p => p.doc_id !== d.doc_id);
    state.documents    = state.documents.filter(x => x.doc_id !== d.doc_id);
    render();
    toast('Documento eliminado.');
  } catch (e) {
    toast('Error eliminando: ' + e.message, true);
  } finally {
    hideLoading();
  }
}

// ---------- Tag manager ----------
function openTagModal() {
  renderTagList();
  document.getElementById('tag-modal').hidden = false;
}
function renderTagList() {
  const ul = document.getElementById('tag-list');
  ul.innerHTML = '';
  for (const t of state.tags) {
    const li = document.createElement('li');
    li.innerHTML = `
      <span class="tag-swatch" style="background:${t.color}"></span>
      <input class="tag-name" type="text" value="${escapeAttr(t.name)}">
      <input type="color" value="${t.color}">
      <button class="btn small" data-act="save">Guardar</button>
      <button class="btn small danger" data-act="del">Eliminar</button>
    `;
    li.querySelector('[data-act=save]').addEventListener('click', async () => {
      const name  = li.querySelector('.tag-name').value.trim();
      const color = li.querySelector('input[type=color]').value;
      if (!name) return;
      showLoading('Guardando etiqueta…');
      try {
        await API.post('updateTag', { tag_id: t.tag_id, patch: { name, color } });
        t.name = name; t.color = color;
        renderTagList(); render();
        toast('Etiqueta actualizada.');
      } catch (e) { toast('Error: ' + e.message, true); }
      finally { hideLoading(); }
    });
    li.querySelector('[data-act=del]').addEventListener('click', async () => {
      if (!confirm(`¿Eliminar la etiqueta "${t.name}"?`)) return;
      showLoading('Eliminando etiqueta…');
      try {
        await API.post('deleteTag', { tag_id: t.tag_id });
        state.tags = state.tags.filter(x => x.tag_id !== t.tag_id);
        state.comment_tags = state.comment_tags.filter(r => String(r.tag_id) !== String(t.tag_id));
        renderTagList(); render();
      } catch (e) { toast('Error: ' + e.message, true); }
      finally { hideLoading(); }
    });
    ul.appendChild(li);
  }
}
async function onTagAdd() {
  const name  = document.getElementById('tag-name').value.trim();
  const color = document.getElementById('tag-color').value;
  if (!name) { toast('Pon un nombre a la etiqueta.', true); return; }
  const tag_id = 'tag_' + Date.now();
  const tag = { tag_id, name, color };
  showLoading('Guardando etiqueta…');
  try {
    await API.post('saveTag', tag);
    state.tags.push(tag);
    document.getElementById('tag-name').value = '';
    renderTagList(); render();
  } catch (e) { toast('Error: ' + e.message, true); }
  finally { hideLoading(); }
}

// ---------- Comment tag picker ----------
function openCtagModal(commentId) {
  state.pendingCtagCommentId = commentId;
  const ul = document.getElementById('ctag-list');
  ul.className = 'tag-list picker';
  ul.innerHTML = '';
  if (!state.tags.length) {
    ul.innerHTML = `<li style="color:#64748b">No hay etiquetas aún. Crea una desde "Etiquetas".</li>`;
  }
  const currentIds = tagIdsOfComment(commentId);
  for (const t of state.tags) {
    const li = document.createElement('li');
    li.innerHTML = `
      <label style="display:flex;align-items:center;gap:8px;flex:1">
        <input type="checkbox" value="${t.tag_id}" ${currentIds.includes(String(t.tag_id)) ? 'checked' : ''}>
        <span class="tag-swatch" style="background:${t.color}"></span>
        <span class="tag-name">${escapeHtml(t.name)}</span>
      </label>
    `;
    ul.appendChild(li);
  }
  document.getElementById('ctag-modal').hidden = false;
}
async function onCtagSave() {
  const cid = state.pendingCtagCommentId;
  const ids = Array.from(document.querySelectorAll('#ctag-list input[type=checkbox]:checked'))
    .map(el => el.value);
  showLoading('Guardando etiquetas…');
  try {
    await API.post('setCommentTags', { comment_id: cid, tag_ids: ids });
    state.comment_tags = state.comment_tags.filter(r => String(r.comment_id) !== String(cid));
    state.comment_tags.push(...ids.map(tid => ({ comment_id: cid, tag_id: tid })));
    document.getElementById('ctag-modal').hidden = true;
    render();
  } catch (e) { toast('Error: ' + e.message, true); }
  finally { hideLoading(); }
}

// ---------- Config ----------
function openConfigModal() {
  document.getElementById('cfg-api-url').value = API.getUrl();
  document.getElementById('config-modal').hidden = false;
}
async function onConfigSave() {
  const url = document.getElementById('cfg-api-url').value.trim();
  API.setUrl(url);
  document.getElementById('config-modal').hidden = true;
  toast('Config guardada. Recargando…');
  await reloadAll();
}

// ---------- Share / PDF ----------
function openShareModal() {
  const docsUl = document.getElementById('share-docs');
  docsUl.innerHTML = '';
  for (const d of state.documents) {
    const li = document.createElement('li');
    const safeId = escapeAttr(d.doc_id);
    li.innerHTML = `<label><input type="checkbox" data-doc-id="${safeId}" checked> ${escapeHtml(d.filename)}</label>`;
    docsUl.appendChild(li);
  }
  document.querySelectorAll('#share-modal [data-hide]').forEach(cb => cb.checked = false);
  document.querySelectorAll('#share-modal input[type=checkbox]').forEach(cb => {
    cb.removeEventListener('change', updateShareURL);
    cb.addEventListener('change', updateShareURL);
  });
  updateShareURL();
  document.getElementById('share-modal').hidden = false;
}

function readShareSettings() {
  const hides = [...document.querySelectorAll('#share-modal [data-hide]:checked')].map(cb => cb.dataset.hide);
  const included = [...document.querySelectorAll('#share-docs input[type=checkbox]:checked')].map(cb => cb.dataset.docId);
  const excluded = state.documents.filter(d => !included.includes(d.doc_id)).map(d => d.doc_id);
  return { hides, excluded };
}

function updateShareURL() {
  const { hides, excluded } = readShareSettings();
  const params = new URLSearchParams();
  params.set('view', '1');
  if (hides.length) params.set('hide', hides.join(','));
  if (excluded.length) params.set('hidedocs', excluded.join(','));
  document.getElementById('share-url').value =
    `${location.origin}${location.pathname}?${params.toString()}`;
}

async function copyShareURL() {
  const url = document.getElementById('share-url').value;
  try {
    await navigator.clipboard.writeText(url);
    toast('URL copiada al portapapeles.');
  } catch {
    document.getElementById('share-url').select();
    try { document.execCommand('copy'); toast('URL copiada.'); }
    catch { toast('No se pudo copiar. Selecciona manualmente.', true); }
  }
}

function openShareURL() {
  window.open(document.getElementById('share-url').value, '_blank');
}

function generateSharePDF() {
  const { hides, excluded } = readShareSettings();

  const oldClasses = [...document.body.classList];
  document.body.classList.add('view-mode');
  for (const h of hides) document.body.classList.add('hide-' + h);

  const hiddenCards = [];
  document.querySelectorAll('.doc-card').forEach(card => {
    if (excluded.includes(card.dataset.docId)) {
      hiddenCards.push([card, card.style.display]);
      card.style.display = 'none';
    }
  });

  document.getElementById('share-modal').hidden = true;

  const restore = () => {
    document.body.className = oldClasses.join(' ');
    hiddenCards.forEach(([c, prev]) => { c.style.display = prev; });
    window.removeEventListener('afterprint', restore);
  };
  window.addEventListener('afterprint', restore);

  setTimeout(() => window.print(), 100);
}

// ---------- Export CSV ----------
function exportCSV() {
  const rows = [[
    'documento', 'pagina', 'pagina_aproximada', 'parrafo_idx', 'parrafo_texto',
    'comentario', 'autor', 'fecha', 'resuelto',
    'etiquetas', 'es_respuesta_a', 'observacion'
  ]];
  for (const c of state.comments) {
    const doc = state.documents.find(d => d.doc_id === c.doc_id);
    const tags = tagsOfComment(c.comment_id).map(t => t.name).join('; ');
    const p = state.paragraphs.find(pp =>
      pp.doc_id === c.doc_id && Number(pp.paragraph_index) === Number(c.paragraph_index));
    const approx = p && (Number(p.page_approx) === 1 || p.page_approx === 'TRUE' || p.page_approx === true);
    rows.push([
      doc ? doc.filename : c.doc_id,
      p ? p.page_number : '',
      approx ? 'sí' : 'no',
      c.paragraph_index,
      p ? p.text : '',
      c.comment_text,
      c.author,
      c.created_at,
      Number(c.resolved) === 1 ? 'sí' : 'no',
      tags,
      c.parent_comment_id || '',
      c.observation || ''
    ]);
  }
  const csv = rows.map(r => r.map(csvCell).join(',')).join('\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `comentarios_${Date.now()}.csv`; a.click();
  URL.revokeObjectURL(url);
}
function csvCell(v) {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// ---------- Backup JSON ----------
async function exportBackup() {
  showLoading('Descargando backup…');
  try {
    const data = await API.loadAll();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `backup_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast('Backup descargado.');
  } catch (e) {
    toast('Error descargando backup: ' + e.message, true);
  } finally {
    hideLoading();
  }
}

// ---------- Utilities ----------
function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}
function escapeAttr(s) { return escapeHtml(s); }
function shortName(s) { return s.length > 28 ? s.slice(0, 26) + '…' : s; }
function formatDate(s) {
  if (!s) return '';
  const d = new Date(s);
  if (isNaN(d)) return s;
  return d.toLocaleDateString('es', { year: 'numeric', month: 'short', day: 'numeric' });
}
function toast(msg, err=false) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.style.background = err ? '#b91c1c' : '#1e293b';
  t.hidden = false;
  clearTimeout(toast._h);
  toast._h = setTimeout(() => t.hidden = true, 2500);
}

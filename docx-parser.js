// Parses a .docx file and returns { docId, paragraphs, comments }.
// docId = SHA-256 hex of the file bytes (to detect duplicates).
const W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

async function sha256Hex(buf) {
  const hash = await crypto.subtle.digest('SHA-256', buf);
  return [...new Uint8Array(hash)].map(b => b.toString(16).padStart(2, '0')).join('');
}

function textOfCommentEl(el) {
  // Preserve paragraph breaks in the comment body.
  const paragraphs = Array.from(el.getElementsByTagNameNS(W, 'p'));
  if (paragraphs.length === 0) {
    return Array.from(el.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('');
  }
  return paragraphs.map(p =>
    Array.from(p.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('')
  ).join('\n');
}

function paragraphTextWithBreaks(p) {
  // Walk runs in order, concatenating text and line breaks.
  let text = '';
  const walk = (node) => {
    if (node.nodeType !== 1) return;
    const ln = node.localName;
    if (ln === 't') { text += node.textContent; return; }
    if (ln === 'tab') { text += '\t'; return; }
    if (ln === 'br') { text += '\n'; return; }
    for (const child of node.childNodes) walk(child);
  };
  walk(p);
  return text.trim();
}

async function parseDocx(file) {
  const buf = await file.arrayBuffer();
  const docId = await sha256Hex(buf);

  const zip = await JSZip.loadAsync(buf);
  const docXmlFile      = zip.file('word/document.xml');
  const commentsFile    = zip.file('word/comments.xml');
  const commentsExtFile = zip.file('word/commentsExtended.xml');

  if (!docXmlFile) throw new Error('document.xml no encontrado');
  const docXml = await docXmlFile.async('string');
  const parser = new DOMParser();
  const docDoc = parser.parseFromString(docXml, 'application/xml');

  // --- Parse all comments (text, author, date, paraId) ---
  const commentsById = {};
  const commentByParaId = {};
  if (commentsFile) {
    const commentsXml = await commentsFile.async('string');
    const cDoc = parser.parseFromString(commentsXml, 'application/xml');
    const cEls = cDoc.getElementsByTagNameNS(W, 'comment');
    for (const el of cEls) {
      const id = el.getAttributeNS(W, 'id');
      const author = el.getAttributeNS(W, 'author') || 'Desconocido';
      const date = el.getAttributeNS(W, 'date') || '';
      const text = textOfCommentEl(el);
      // paraId is on the first w:p of the comment body (in w14 namespace).
      const firstP = el.getElementsByTagNameNS(W, 'p')[0];
      const paraId = firstP ? firstP.getAttributeNS(W14, 'paraId') : null;
      const c = { id, author, date, text, paraId, parentId: null, resolved: false };
      commentsById[id] = c;
      if (paraId) commentByParaId[paraId] = c;
    }
  }

  // --- Extended: parent (replies) + resolved ---
  if (commentsExtFile) {
    const extXml = await commentsExtFile.async('string');
    const eDoc = parser.parseFromString(extXml, 'application/xml');
    const eEls = eDoc.getElementsByTagNameNS(W15, 'commentEx');
    for (const el of eEls) {
      const paraId       = el.getAttributeNS(W15, 'paraId');
      const paraIdParent = el.getAttributeNS(W15, 'paraIdParent');
      const done         = el.getAttributeNS(W15, 'done') === '1';
      const c = commentByParaId[paraId];
      if (!c) continue;
      c.resolved = done;
      if (paraIdParent && commentByParaId[paraIdParent]) {
        c.parentId = commentByParaId[paraIdParent].id;
      }
    }
  }

  // --- Walk document body paragraphs ---
  const body = docDoc.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('body no encontrado');

  const paragraphs = [];
  const commentToParaIndex = {}; // commentId -> first paragraph index where it starts

  const openComments = new Set(); // comments whose Range started but hasn't ended yet (spanning multiple paragraphs)
  let paraIdx = 0;

  const processParagraph = (p) => {
    const idx = paraIdx++;
    const text = paragraphTextWithBreaks(p);
    paragraphs.push({ paragraph_index: idx, text });

    // Find commentRangeStart elements inside this paragraph (in order).
    const starts = p.getElementsByTagNameNS(W, 'commentRangeStart');
    for (const s of starts) {
      const cid = s.getAttributeNS(W, 'id');
      if (commentToParaIndex[cid] === undefined) commentToParaIndex[cid] = idx;
      openComments.add(cid);
    }
    const ends = p.getElementsByTagNameNS(W, 'commentRangeEnd');
    for (const e of ends) openComments.delete(e.getAttributeNS(W, 'id'));

    // If no commentRangeStart in this paragraph but a comment is "open" from earlier,
    // the paragraph is still part of that comment's span — but we don't re-anchor.
  };

  // Iterate only direct children of body that are paragraphs or tables (skip sectPr).
  for (const child of body.childNodes) {
    if (child.nodeType !== 1) continue;
    const ln = child.localName;
    if (ln === 'p') {
      processParagraph(child);
    } else if (ln === 'tbl') {
      // Include paragraphs inside table cells so we don't lose commented text in tables.
      const innerPs = child.getElementsByTagNameNS(W, 'p');
      for (const p of innerPs) processParagraph(p);
    }
  }

  // Build output comments with paragraph_index.
  const comments = Object.values(commentsById)
    .filter(c => commentToParaIndex[c.id] !== undefined)
    .map(c => ({
      id: c.id,
      paragraph_index: commentToParaIndex[c.id],
      text: c.text,
      author: c.author,
      date: c.date,
      parentId: c.parentId,
      resolved: c.resolved
    }));

  return { docId, paragraphs, comments };
}

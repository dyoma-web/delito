// Parses a .docx file and returns { docId, paragraphs, comments }.
// docId = SHA-256 hex of the file bytes (to detect duplicates).
const W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

const CHARS_PER_PAGE_FALLBACK = 3500;

async function sha256Hex(buf) {
  const hash = await crypto.subtle.digest('SHA-256', buf);
  return [...new Uint8Array(hash)].map(b => b.toString(16).padStart(2, '0')).join('');
}

function textOfCommentEl(el) {
  const paragraphs = Array.from(el.getElementsByTagNameNS(W, 'p'));
  if (paragraphs.length === 0) {
    return Array.from(el.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('');
  }
  return paragraphs.map(p =>
    Array.from(p.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('')
  ).join('\n');
}

function paragraphTextWithBreaks(p) {
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

// Returns starting page number from the last sectPr (1 if absent).
function getStartPage(docDoc) {
  const sectPrs = docDoc.getElementsByTagNameNS(W, 'sectPr');
  for (const sp of sectPrs) {
    const pgNumType = sp.getElementsByTagNameNS(W, 'pgNumType')[0];
    if (pgNumType) {
      const start = pgNumType.getAttributeNS(W, 'start');
      if (start) return parseInt(start, 10) || 1;
    }
  }
  return 1;
}

// True if the doc contains Word-generated page-break hints.
function docHasPageHints(docDoc) {
  return docDoc.getElementsByTagNameNS(W, 'lastRenderedPageBreak').length > 0;
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

  // ----- Parse comments text/author/date -----
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
      const firstP = el.getElementsByTagNameNS(W, 'p')[0];
      const paraId = firstP ? firstP.getAttributeNS(W14, 'paraId') : null;
      const c = { id, author, date, text, paraId, parentId: null, resolved: false };
      commentsById[id] = c;
      if (paraId) commentByParaId[paraId] = c;
    }
  }

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

  // ----- Walk body paragraphs, compute page numbers and anchor comments -----
  const body = docDoc.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('body no encontrado');

  const startPage = getStartPage(docDoc);
  const hasHints = docHasPageHints(docDoc);

  const paragraphs = [];
  const commentToParaIndex = {};

  let paraIdx = 0;
  let currentPage = startPage;   // page where the NEXT paragraph starts
  let charCount = 0;             // for fallback estimation

  // processParagraph anchors the paragraph to its starting page and walks in
  // document order so that any lastRenderedPageBreak encountered MID-paragraph
  // increments currentPage for subsequent paragraphs (and for later comments
  // in this paragraph).
  const processParagraph = (p) => {
    const idx = paraIdx++;
    const text = paragraphTextWithBreaks(p);

    // pageBreakBefore: this paragraph starts on a fresh page.
    const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
    if (pPr && pPr.getElementsByTagNameNS(W, 'pageBreakBefore').length > 0) {
      currentPage++;
    }

    const paragraphPage = currentPage;
    const approx = !hasHints;

    // Walk paragraph descendants in order; react to page breaks and
    // commentRangeStart to anchor comment IDs.
    const walk = (node) => {
      for (const child of node.childNodes) {
        if (child.nodeType !== 1) continue;
        const ln = child.localName;
        const ns = child.namespaceURI;
        if (ns === W) {
          if (ln === 'lastRenderedPageBreak') { currentPage++; continue; }
          if (ln === 'br' && child.getAttributeNS(W, 'type') === 'page') { currentPage++; continue; }
          if (ln === 'commentRangeStart') {
            const cid = child.getAttributeNS(W, 'id');
            if (commentToParaIndex[cid] === undefined) commentToParaIndex[cid] = idx;
            continue;
          }
        }
        walk(child);
      }
    };
    walk(p);

    // Fallback page estimation if no hints anywhere in the doc.
    let pageNumber = paragraphPage;
    if (!hasHints) {
      pageNumber = startPage + Math.floor(charCount / CHARS_PER_PAGE_FALLBACK);
      charCount += text.length + 1;
    }

    paragraphs.push({
      paragraph_index: idx,
      text,
      page_number: pageNumber,
      page_approx: approx ? 1 : 0
    });
  };

  for (const child of body.childNodes) {
    if (child.nodeType !== 1) continue;
    const ln = child.localName;
    if (ln === 'p') {
      processParagraph(child);
    } else if (ln === 'tbl') {
      const innerPs = child.getElementsByTagNameNS(W, 'p');
      for (const p of innerPs) processParagraph(p);
    }
  }

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

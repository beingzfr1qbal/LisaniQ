/* Unani Terminology Converter - Office.js task pane logic
   Matching is a deterministic dictionary lookup against WHO IUMT 2022 / CCRUM 2012 -
   nothing is generated or guessed. If a phrase isn't in the database, it is left unchanged. */

const MAX_NGRAM = 10;
let TERMDB = null;      // { ambiguous: [...], terms: { "english phrase": [ {unani,source,termId,def}, ... ] } }
let AMBIGUOUS = new Set();
let conversionLog = [];

Office.onReady(() => {
  fetch('../assets/terminology.json')
    .then(r => r.json())
    .then(data => {
      TERMDB = data.terms;
      AMBIGUOUS = new Set(data.ambiguous);
      document.getElementById('status').textContent =
        `Loaded ${Object.keys(TERMDB).length.toLocaleString()} terms (WHO IUMT 2022 + CCRUM 2012).`;
      document.getElementById('btnConvertSelection').disabled = false;
      document.getElementById('btnConvertDocument').disabled = false;
      document.getElementById('searchBox').disabled = false;
    })
    .catch(err => {
      document.getElementById('status').textContent =
        'Failed to load terminology database: ' + err.message;
    });

  document.getElementById('btnConvertSelection').onclick = () => runConversion('selection');
  document.getElementById('btnConvertDocument').onclick = () => runConversion('document');
  document.getElementById('searchBox').addEventListener('input', onSearchInput);
});

// ---------- Matching ----------

function tokenize(text) {
  // keep positions so we can map back to ranges if needed later
  const re = /[A-Za-z][A-Za-z'-]*|[^A-Za-z\s]|\s+/g;
  const tokens = [];
  let m;
  while ((m = re.exec(text)) !== null) tokens.push(m[0]);
  return tokens;
}

function isWordToken(t) {
  return /^[A-Za-z]/.test(t);
}

/**
 * Find matches in text. Returns array of {original, unani, source, termId, confidence}
 * in order of first occurrence, de-duplicated by exact phrase text.
 */
function findMatches(text, includeReview) {
  const tokens = tokenize(text);
  const wordIdx = [];
  tokens.forEach((t, i) => { if (isWordToken(t)) wordIdx.push(i); });

  const found = [];
  const seen = new Set();
  let skipUntil = -1;

  for (let wi = 0; wi < wordIdx.length; wi++) {
    if (wordIdx[wi] <= skipUntil) continue;
    let matchedLen = 0;
    let matchedKey = null;

    for (let n = Math.min(MAX_NGRAM, wordIdx.length - wi); n >= 1; n--) {
      const startTok = wordIdx[wi];
      const endTok = wordIdx[wi + n - 1];
      const phrase = tokens.slice(startTok, endTok + 1).join('').toLowerCase();
      // normalize internal whitespace to single spaces for lookup
      const key = phrase.replace(/\s+/g, ' ').trim();
      if (TERMDB.hasOwnProperty(key)) {
        matchedLen = n;
        matchedKey = key;
        break; // longest match wins
      }
    }

    if (matchedKey) {
      const isAmbig = AMBIGUOUS.has(matchedKey);
      if (isAmbig && !includeReview) {
        // still report it in the log as REVIEW, but don't apply
        const candidates = TERMDB[matchedKey];
        const original = tokens.slice(wordIdx[wi], wordIdx[wi + matchedLen - 1] + 1).join('');
        if (!seen.has(matchedKey)) {
          seen.add(matchedKey);
          found.push({ original, unani: candidates[0].unani, source: candidates[0].source,
            termId: candidates[0].termId, confidence: 'review', apply: false });
        }
      } else {
        const candidates = TERMDB[matchedKey];
        const original = tokens.slice(wordIdx[wi], wordIdx[wi + matchedLen - 1] + 1).join('');
        found.push({ original, unani: candidates[0].unani, source: candidates[0].source,
          termId: candidates[0].termId, confidence: isAmbig ? 'review' : 'high', apply: true,
          startTok: wordIdx[wi], endTok: wordIdx[wi + matchedLen - 1] });
      }
      skipUntil = wordIdx[wi + matchedLen - 1];
    }
  }
  return found;
}

// ---------- Word document actions ----------

async function runConversion(scope) {
  const includeReview = document.getElementById('optIncludeReview').checked;
  const highlight = document.getElementById('optHighlight').checked;
  conversionLog = [];

  try {
    await Word.run(async (context) => {
      const range = scope === 'selection'
        ? context.document.getSelection()
        : context.document.body;
      range.load('text');
      await context.sync();

      const text = range.text;
      if (!text || !text.trim()) {
        appendLog('Nothing to convert (empty selection/document).');
        return;
      }

      const matches = findMatches(text, includeReview);
      const toApply = matches.filter(m => m.apply);
      const reviewOnly = matches.filter(m => !m.apply);

      // Apply each distinct matched phrase via Word's native search-and-replace,
      // which preserves surrounding formatting since only the found range is touched.
      const uniqueApplied = new Map();
      for (const m of toApply) {
        if (uniqueApplied.has(m.original.toLowerCase())) continue;
        uniqueApplied.set(m.original.toLowerCase(), m);
      }

      for (const m of uniqueApplied.values()) {
        const searchResults = range.search(m.original, { matchCase: false, matchWholeWord: true });
        searchResults.load('text');
        await context.sync();

        searchResults.items.forEach(found => {
          const replacementText = `${m.unani} [${m.original}]`;
          found.insertText(replacementText, Word.InsertLocation.replace);
          if (highlight) {
            found.font.highlightColor = m.confidence === 'high' ? '#C6EFCE' : '#FFEB9C';
          }
        });
        await context.sync();
      }

      conversionLog = matches;
      renderLog(matches, includeReview);
    });
  } catch (err) {
    appendLog('Error: ' + (err && err.message ? err.message : String(err)));
    console.error(err);
  }
}

function renderLog(matches, includeReview) {
  const logArea = document.getElementById('logArea');
  if (!matches.length) {
    logArea.innerHTML = 'No matching terminology found in the selected text.';
    return;
  }
  logArea.innerHTML = matches.map(m => {
    const cls = m.confidence === 'high' ? 'conf-high' : 'conf-review';
    const status = m.apply ? 'applied' : (includeReview ? 'applied' : 'flagged - not applied');
    return `<div class="match-row">
      <span class="${cls}">[${m.confidence.toUpperCase()}]</span>
      "${escapeHtml(m.original)}" &rarr; <b>${escapeHtml(m.unani)}</b>
      <br/><small>${m.source} · ${m.termId} · ${status}</small>
    </div>`;
  }).join('');
}

function appendLog(msg) {
  const logArea = document.getElementById('logArea');
  logArea.innerHTML = `<div class="match-row">${escapeHtml(msg)}</div>` + logArea.innerHTML;
}

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ---------- Search box ----------

function onSearchInput(e) {
  const q = e.target.value.trim().toLowerCase();
  const resultsDiv = document.getElementById('searchResults');
  if (!q || !TERMDB) { resultsDiv.innerHTML = ''; return; }

  const hits = Object.keys(TERMDB)
    .filter(k => k.includes(q))
    .slice(0, 25);

  if (!hits.length) {
    resultsDiv.innerHTML = '<div class="match-row">No matches.</div>';
    return;
  }

  resultsDiv.innerHTML = hits.map(k => {
    const c = TERMDB[k][0];
    const cls = AMBIGUOUS.has(k) ? 'conf-review' : 'conf-high';
    return `<div class="match-row">
      <b>${escapeHtml(k)}</b> &rarr; <span class="${cls}">${escapeHtml(c.unani)}</span>
      <br/><small>${c.source} · ${c.termId}</small>
      <br/><small>${escapeHtml(c.def)}</small>
    </div>`;
  }).join('');
}

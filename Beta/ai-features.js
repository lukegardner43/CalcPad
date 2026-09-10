/* ══════════════════════════════════════════════════════════════════════
   CalcPad Beta — AI features
   ══════════════════════════════════════════════════════════════════════

   Loaded by index.html only in beta mode (see the beta loader at the foot
   of the main script). This file is additive: it reads the document
   through the same helpers the Word export uses, and writes back through
   the same paste pipeline typing does. It never redefines anything the
   core app owns, so the production app is byte-for-byte unaffected when
   this file is not loaded.

   Two features:
     1. AI QA      — reviews document flow, spelling and grammar.
     2. Assistant  — a chat box that drafts CalcPad lines from a prompt.

   The API key is the tester's own, held in sessionStorage for the tab
   only. It is never written into an autosave, a .perega file, or a URL.
   ══════════════════════════════════════════════════════════════════════ */

(function () {
'use strict';

/* ── Configuration ─────────────────────────────────────────────────── */

const ENDPOINT = 'https://api.anthropic.com/v1/messages';
const API_VER  = '2023-06-01';
const KEY_SLOT = 'calcpad_beta_anthropic_key';

// Per-model request capabilities. Sending `thinking` or `effort` to a model
// that does not take them is a 400, so each is gated rather than assumed.
const MODELS = {
  'claude-opus-5':    { label: 'Opus 5 — most capable', thinking: true,  effort: true  },
  'claude-sonnet-5':  { label: 'Sonnet 5 — balanced',   thinking: true,  effort: true  },
  'claude-haiku-4-5': { label: 'Haiku 4.5 — cheapest',  thinking: false, effort: false },
};

const state = {
  model:   'claude-opus-5',
  effort:  'high',
  running: null,     // AbortController while a request is in flight
  findings: [],
  filter:  'all',
  chat:    [],       // [{role:'user'|'assistant', content:'…'}]
};

/* ── Small helpers ─────────────────────────────────────────────────── */

const $ = id => document.getElementById(id);

const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g,
  c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function el(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text != null) n.textContent = text;
  return n;
}

/* sessionStorage throws outright in some privacy modes, so every touch is
   guarded and a failure degrades to "no key stored" rather than a broken panel. */
function getKey() {
  try { return sessionStorage.getItem(KEY_SLOT) || ''; } catch (e) { return ''; }
}
function setKey(k) {
  try { k ? sessionStorage.setItem(KEY_SLOT, k) : sessionStorage.removeItem(KEY_SLOT); }
  catch (e) { /* nothing to do — the modal reopens next run */ }
  syncKeyState();
}

/* ── Reading the document ──────────────────────────────────────────── */

// chunkPlainText() is the core app's own serialiser — the one behind copy and
// Word export — so what the model reviews is exactly what a reader sees:
// chips come through as "M_Ed = w*L^2/8 = 45 kN·m", not as markup.
function pageTexts() {
  return getAllEditors().map(ed => chunkPlainText(ed));
}

function documentIsEmpty() {
  return pageTexts().every(t => !t.trim());
}

// One flat line numbering across every page, so a finding can cite "line 24"
// and the reader can find it. Page breaks are marked but not numbered.
function numberedDocument() {
  const out = [];
  let n = 0;
  pageTexts().forEach((txt, p) => {
    out.push('[Page ' + (p + 1) + ']');
    txt.split('\n').forEach(line => { out.push(String(++n).padStart(3, ' ') + ' | ' + line); });
  });
  return out.join('\n');
}

function plainDocument() {
  return pageTexts()
    .map((t, i) => '--- Page ' + (i + 1) + ' ---\n' + (t.trim() || '(empty)'))
    .join('\n\n');
}

const META_LABELS = {
  'm-project': 'Project', 'm-projno': 'Project no', 'm-date': 'Date',
  'm-desc': 'Description', 'm-sheet': 'Sheet', 'm-orig': 'Originator',
  'm-deliv': 'Deliverable no', 'm-checked': 'Checked', 'm-suit': 'Suitability',
  'm-rev': 'Revision',
};

function metaText() {
  const m = collectMeta();
  const rows = Object.keys(META_LABELS)
    .filter(id => (m[id] || '').trim())
    .map(id => META_LABELS[id] + ': ' + m[id]);
  return rows.length ? rows.join('\n') : '(title block not filled in)';
}

// Every live definition, in document order, with its computed value in the
// units actually shown on screen.
function variableTable() {
  const chips = Array.from(document.querySelectorAll('.var-def-chip'));
  if (!chips.length) return '(no variables defined)';
  return chips.map(c => '  ' + varChipExportText(c)).join('\n');
}

function preferredUnitsText() {
  try {
    return Object.keys(preferredUnits).map(k => k + ' → ' + preferredUnits[k]).join(', ');
  } catch (e) { return '(default)'; }
}

function documentContext(numbered) {
  return [
    'TITLE BLOCK',
    metaText(),
    '',
    'DEFAULT OUTPUT UNITS',
    preferredUnitsText(),
    '',
    'VARIABLES CURRENTLY DEFINED (name = expression = computed value)',
    variableTable(),
    '',
    'DOCUMENT',
    numbered ? numberedDocument() : plainDocument(),
  ].join('\n');
}

/* ── Writing back into the document ────────────────────────────────── */

// Hands the lines to the app's own paste path, which builds real chips,
// evaluates them, and asks about duplicate names — identical to typing them.
function insertCalcLines(text) {
  const eds = getAllEditors();
  if (!eds.length) return false;

  let range = null;
  // Prefer wherever the caret last sat in a calc area; the panel stealing
  // focus is why the core app tracks lastEditRange at all.
  if (typeof lastEditRange !== 'undefined' && lastEditRange) {
    for (const ed of eds) {
      if (ed.contains(lastEditRange.startContainer)) { range = lastEditRange.cloneRange(); break; }
    }
  }
  if (!range) {
    // Never used the editor this session — append to the end of the last page,
    // breaking the line first so the block does not run into existing prose.
    const ed = eds[eds.length - 1];
    range = document.createRange();
    range.selectNodeContents(ed);
    range.collapse(false);
    if (ed.textContent.trim()) text = '\n' + text;
  }

  const host = range.startContainer.nodeType === Node.ELEMENT_NODE
    ? range.startContainer.closest('.calc-area')
    : (range.startContainer.parentElement && range.startContainer.parentElement.closest('.calc-area'));
  if (host) { setActiveEditor(host); host.focus(); }

  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  insertPastedText(text, range);
  return true;
}

/* ── Locating a quoted finding on the page ─────────────────────────── */

// Draws a temporary overlay over the matching text instead of wrapping it in
// markup: the editor DOM is left untouched, so nothing is marked dirty and no
// chip is disturbed.
function clearHighlights() {
  document.querySelectorAll('.ai-hl').forEach(n => n.remove());
}

function flashRange(range) {
  clearHighlights();
  const rects = Array.from(range.getClientRects());
  if (!rects.length) return false;
  const sx = window.scrollX, sy = window.scrollY;
  rects.forEach(r => {
    const box = el('div', 'ai-hl');
    box.style.left   = (r.left + sx) + 'px';
    box.style.top    = (r.top  + sy) + 'px';
    box.style.width  = r.width  + 'px';
    box.style.height = r.height + 'px';
    document.body.appendChild(box);
  });
  window.scrollTo({ top: rects[0].top + sy - window.innerHeight / 3, behavior: 'smooth' });
  setTimeout(clearHighlights, 2600);
  return true;
}

// Build a flat map of an editor's text so a plain-string search can be turned
// back into a Range across whatever nodes it happens to straddle.
function textIndex(root) {
  const nodes = [], starts = [];
  let flat = '';
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = walker.nextNode())) {
    nodes.push(n); starts.push(flat.length); flat += n.nodeValue;
  }
  return { nodes, starts, flat };
}

function rangeAt(idx, from, len) {
  const end = from + len;
  let sNode = null, sOff = 0, eNode = null, eOff = 0;
  for (let i = 0; i < idx.nodes.length; i++) {
    const a = idx.starts[i], b = a + idx.nodes[i].nodeValue.length;
    if (sNode === null && from >= a && from < b) { sNode = idx.nodes[i]; sOff = from - a; }
    if (end > a && end <= b) { eNode = idx.nodes[i]; eOff = end - a; break; }
  }
  if (!sNode) return null;
  if (!eNode) { eNode = idx.nodes[idx.nodes.length - 1]; eOff = eNode.nodeValue.length; }
  const r = document.createRange();
  r.setStart(sNode, sOff); r.setEnd(eNode, eOff);
  return r;
}

// A quote comes back from the model as it appeared in the serialised text,
// which may differ from the on-screen run (chips render their value, prose
// may wrap). Try the whole quote, then shorter leading slices, before
// admitting defeat.
function locateQuote(quote) {
  const q = String(quote || '').replace(/\s+/g, ' ').trim();
  if (q.length < 3) return false;
  const tries = [q, q.slice(0, 60), q.slice(0, 30), q.slice(0, 16)]
    .filter((s, i, a) => s.length >= 4 && a.indexOf(s) === i);

  for (const ed of getAllEditors()) {
    const idx = textIndex(ed);
    const hay = idx.flat.replace(/\s+/g, ' ').toLowerCase();
    // Whitespace collapsing shifts offsets, so search the raw text too and
    // prefer whichever hits — the raw hit maps straight back to nodes.
    const raw = idx.flat.toLowerCase();
    for (const t of tries) {
      const needle = t.toLowerCase();
      let at = raw.indexOf(needle);
      if (at === -1 && hay.indexOf(needle) !== -1) {
        // Collapsed-only match: fall back to the first word, which survives
        // collapsing intact and is enough to scroll the reader to the spot.
        const w = needle.split(' ')[0];
        at = w.length >= 4 ? raw.indexOf(w) : -1;
        if (at !== -1) { const r = rangeAt(idx, at, w.length); if (r && flashRange(r)) return true; }
        continue;
      }
      if (at !== -1) { const r = rangeAt(idx, at, t.length); if (r && flashRange(r)) return true; }
    }
  }
  return false;
}

/* ── Talking to the API ────────────────────────────────────────────── */

function buildRequest(o) {
  const caps = MODELS[state.model] || MODELS['claude-opus-5'];
  const req = {
    model: state.model,
    max_tokens: o.maxTokens || 32000,
    stream: true,
    system: o.system,
    messages: o.messages,
  };
  // Adaptive thinking is what makes the live progress readout possible;
  // Haiku 4.5 does not take it, and rejects `effort` outright.
  if (caps.thinking) req.thinking = { type: 'adaptive', display: 'summarized' };
  const oc = {};
  if (caps.effort) oc.effort = state.effort;
  if (o.schema) oc.format = { type: 'json_schema', schema: o.schema };
  if (Object.keys(oc).length) req.output_config = oc;
  return req;
}

function friendlyError(status, body) {
  const msg = (body && body.error && body.error.message) || body || ('HTTP ' + status);
  if (status === 401) return 'Your API key was rejected (401). Check it and enter it again.';
  if (status === 403) return 'That key is not permitted to call this model (403). ' + msg;
  if (status === 429) return 'Rate limited (429) — wait a moment and try again.';
  if (status === 400 && /credit balance|billing/i.test(msg))
    return 'Your Anthropic account has no credit available. ' + msg;
  if (status >= 500) return 'Anthropic returned a server error (' + status + '). Try again shortly.';
  return msg;
}

// Streams the response over SSE. Streaming rather than a single JSON reply
// keeps a long run from tripping an idle-connection timeout, and lets the
// panel show progress instead of hanging silently for a minute.
async function streamMessage(req, on) {
  const key = getKey();
  if (!key) throw new Error('No API key set for this session.');

  const ctl = new AbortController();
  state.running = ctl;

  let res;
  try {
    res = await fetch(ENDPOINT, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': key,
        'anthropic-version': API_VER,
        // Required for a browser to call the API directly; without it the
        // request never leaves the page.
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify(req),
      signal: ctl.signal,
    });
  } catch (e) {
    state.running = null;
    if (e.name === 'AbortError') throw e;
    throw new Error('Could not reach api.anthropic.com — check your connection. (' + e.message + ')');
  }

  if (!res.ok) {
    state.running = null;
    let body = null;
    try { body = await res.json(); } catch (e) { try { body = await res.text(); } catch (e2) {} }
    throw new Error(friendlyError(res.status, body));
  }

  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let buf = '', stopReason = null, usage = { input: 0, output: 0 };

  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });

      let cut;
      while ((cut = buf.indexOf('\n\n')) !== -1) {
        const block = buf.slice(0, cut);
        buf = buf.slice(cut + 2);
        for (const line of block.split('\n')) {
          if (line.slice(0, 5) !== 'data:') continue;
          const payload = line.slice(5).trim();
          if (!payload) continue;
          let ev;
          try { ev = JSON.parse(payload); } catch (e) { continue; }

          if (ev.type === 'error') throw new Error((ev.error && ev.error.message) || 'Stream error');
          if (ev.type === 'message_start' && ev.message && ev.message.usage)
            usage.input = ev.message.usage.input_tokens || 0;
          if (ev.type === 'content_block_delta' && ev.delta) {
            if (ev.delta.type === 'text_delta' && on.text) on.text(ev.delta.text);
            else if (ev.delta.type === 'thinking_delta' && on.thinking) on.thinking(ev.delta.thinking);
          }
          if (ev.type === 'message_delta') {
            if (ev.delta && ev.delta.stop_reason) stopReason = ev.delta.stop_reason;
            if (ev.usage && ev.usage.output_tokens) usage.output = ev.usage.output_tokens;
          }
        }
      }
    }
  } finally {
    state.running = null;
  }

  // A safety decline arrives as a normal 200 with this stop reason, so it has
  // to be checked rather than inferred from empty content.
  if (stopReason === 'refusal')
    throw new Error('The model declined to answer this request.');
  if (stopReason === 'max_tokens')
    throw new Error('The reply hit the length limit before finishing. Try a shorter document or a narrower question.');

  return { stopReason, usage };
}

function showUsage(u) {
  const n = $('aiUsage');
  if (n && u) n.textContent = u.input.toLocaleString() + ' in / ' + u.output.toLocaleString() + ' out tokens';
}

/* ── Feature 1: AI QA ──────────────────────────────────────────────── */

const QA_CATEGORIES = ['flow', 'spelling', 'grammar', 'clarity', 'units', 'consistency', 'completeness'];

const QA_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['summary', 'findings'],
  properties: {
    summary: {
      type: 'string',
      description: 'Two or three sentences on how the calculation reads overall.',
    },
    findings: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['category', 'severity', 'line', 'quote', 'issue', 'suggestion'],
        properties: {
          category: { type: 'string', enum: QA_CATEGORIES },
          severity: { type: 'string', enum: ['high', 'medium', 'low'] },
          line: { type: 'integer', description: 'Line number from the numbered document, or 0 if it applies throughout.' },
          quote: { type: 'string', description: 'The exact text as it appears in the document, so the reader can find it. Empty if the finding is about the document as a whole.' },
          issue: { type: 'string', description: 'What is wrong, in one or two sentences.' },
          suggestion: { type: 'string', description: 'The specific correction or rewording to apply.' },
        },
      },
    },
  },
};

const QA_SYSTEM = [
  'You are checking an engineering calculation sheet written in CalcPad, a structural',
  'calculation tool. You are acting as the checker a chartered engineer would hand the',
  'sheet to before it is issued.',
  '',
  'Report on three things above all:',
  '',
  '1. FLOW — does the document read as a coherent argument? Is there a stated purpose,',
  '   are assumptions and references given before they are relied on, is each quantity',
  '   introduced before it is used, do the steps follow in a logical order, and does it',
  '   reach a stated conclusion? Call out sections that jump, repeat, or trail off, and',
  '   results that appear without saying what they are for or whether they pass.',
  '2. SPELLING — misspelt words in the prose, in headings, and in the title block.',
  '   British English is the house style: "analyse", "centre", "utilisation", "metre".',
  '3. GRAMMAR — sentence structure, agreement, tense, punctuation, and phrasing that a',
  '   reader would stumble over. Prefer plain, direct engineering prose.',
  '',
  'Also flag, where you see it: unclear or unlabelled quantities; units that look wrong',
  'or inconsistent with the quantity being described; a defined variable that is never',
  'used, or a symbol used in prose that is never defined; and obvious gaps such as a',
  'load case, a check, or a code reference that the document implies but does not carry.',
  '',
  'CalcPad notation, so you do not mistake it for an error:',
  '  "L = 6 m"                 a variable definition with its unit',
  '  "M = w*L^2/8 = 45 kN·m"   a definition, its expression, and the computed value',
  '  "=w*L^2/8 = 45 kN·m"      a standalone formula and its result',
  '  "[kNm]" or "[#MPa]"       an output unit; the # form deliberately forces a unit',
  'Values shown after "=" are computed by the tool, not typed, so do not report them as',
  'typing errors. Do not re-derive arithmetic — you cannot see the full precision.',
  '',
  'Rules for your findings:',
  '- Quote the text exactly as it appears in the document, so the reader can locate it.',
  '- Give the line number from the numbered document. Use 0 only when the finding is',
  '  about the document as a whole.',
  '- Every finding must carry a specific correction, not just an observation. For',
  '  spelling and grammar give the corrected wording verbatim.',
  '- severity: "high" where a reader would be misled or the check is unsound,',
  '  "medium" where the meaning is unclear or the reading is disrupted,',
  '  "low" for polish.',
  '- Report real problems only. An empty findings list is the right answer for a clean',
  '  document — do not pad it. Never report the same problem twice.',
  '- Order findings by severity, highest first.',
].join('\n');

async function runQA() {
  if (documentIsEmpty()) {
    qaMessage('There is nothing to review yet — write some of the calculation first.');
    return;
  }
  if (!ensureKey(runQA)) return;

  const btn = $('aiQARun');
  btn.disabled = true;
  btn.textContent = 'Reviewing…';
  $('aiQAStop').style.display = '';
  setStatus('aiQAStatus', 'Reading the document…', true);
  $('aiQAResults').innerHTML = '';
  state.findings = [];
  state.filter = 'all';

  const think = $('aiQAThink');
  think.className = 'ai-think';
  think.innerHTML = '';

  let json = '';
  let thinking = '';

  try {
    const { usage } = await streamMessage(buildRequest({
      system: [{ type: 'text', text: QA_SYSTEM }],
      messages: [{ role: 'user', content:
        'Review this calculation sheet.\n\n' + documentContext(true) }],
      schema: QA_SCHEMA,
      maxTokens: 32000,
    }), {
      thinking: t => {
        thinking += t;
        think.className = 'ai-think on';
        think.innerHTML = '<b>Reviewing</b>' + esc(thinking.slice(-1200));
        think.scrollTop = think.scrollHeight;
      },
      text: t => {
        json += t;
        setStatus('aiQAStatus', 'Writing up findings…', true);
      },
    });
    showUsage(usage);

    let data;
    try { data = JSON.parse(json); }
    catch (e) { throw new Error('The reply was not valid JSON — try running the review again.'); }
    state.findings = Array.isArray(data.findings) ? data.findings : [];
    renderQA(data.summary || '');
    setStatus('aiQAStatus', state.findings.length
      ? state.findings.length + ' finding' + (state.findings.length === 1 ? '' : 's') + '.'
      : 'No issues found.', false);
    think.className = 'ai-think';
  } catch (e) {
    think.className = 'ai-think';
    if (e.name === 'AbortError') setStatus('aiQAStatus', 'Review stopped.', false);
    else { setStatus('aiQAStatus', '', false); qaError(e.message); }
  } finally {
    btn.disabled = false;
    btn.textContent = 'Run QA review';
    $('aiQAStop').style.display = 'none';
  }
}

function qaMessage(msg) {
  $('aiQAResults').innerHTML = '';
  const box = el('div', 'ai-empty');
  box.innerHTML = '<svg viewBox="0 0 24 24"><path d="M12 2l2.4 6.4L21 10l-5 4.2L17.6 21 12 17.3 6.4 21 8 14.2 3 10l6.6-1.6z"/></svg>';
  box.appendChild(el('div', null, msg));
  $('aiQAResults').appendChild(box);
}

function qaError(msg) {
  $('aiQAResults').innerHTML = '';
  $('aiQAResults').appendChild(el('div', 'ai-err', msg));
}

function renderQA(summary) {
  const host = $('aiQAResults');
  host.innerHTML = '';

  if (summary) host.appendChild(el('div', 'ai-summary', summary));

  if (!state.findings.length) {
    const ok = el('div', 'ai-empty');
    ok.innerHTML = '<svg viewBox="0 0 24 24"><path d="M9 16.2l-3.5-3.5L4 14.2 9 19.2 20 8.2l-1.5-1.5z"/></svg>';
    ok.appendChild(el('div', null, 'Nothing flagged. The flow, spelling and grammar read clean.'));
    host.appendChild(ok);
    return;
  }

  // Category filter chips, ordered by how many findings each carries.
  const counts = {};
  state.findings.forEach(f => { counts[f.category] = (counts[f.category] || 0) + 1; });
  const row = el('div', 'ai-count-row');
  const mk = (key, label) => {
    const c = el('button', 'ai-chip' + (state.filter === key ? ' on' : ''), label);
    c.onclick = () => { state.filter = key; renderQA(summary); };
    return c;
  };
  row.appendChild(mk('all', 'All ' + state.findings.length));
  Object.keys(counts).sort((a, b) => counts[b] - counts[a])
    .forEach(k => row.appendChild(mk(k, k + ' ' + counts[k])));
  host.appendChild(row);

  const shown = state.filter === 'all'
    ? state.findings
    : state.findings.filter(f => f.category === state.filter);

  shown.forEach(f => host.appendChild(findingCard(f)));
}

function findingCard(f) {
  const sev = ['high', 'medium', 'low'].indexOf(f.severity) === -1 ? 'low' : f.severity;
  const card = el('div', 'ai-find sev-' + sev);

  const top = el('div', 'ai-find-top');
  top.appendChild(el('span', 'ai-cat', f.category || 'note'));
  top.appendChild(el('span', 'ai-sev sev-' + sev, sev));
  if (f.line) top.appendChild(el('span', 'ai-line', 'line ' + f.line));
  card.appendChild(top);

  card.appendChild(el('div', 'ai-issue', f.issue || ''));
  if (f.quote) card.appendChild(el('div', 'ai-quote', '“' + f.quote + '”'));

  if (f.suggestion) {
    const fix = el('div', 'ai-fix');
    fix.appendChild(el('b', null, 'Suggested: '));
    fix.appendChild(document.createTextNode(f.suggestion));
    card.appendChild(fix);
  }

  card.title = 'Click to find this in the document';
  card.onclick = () => {
    if (!f.quote || !locateQuote(f.quote)) {
      const was = card.title;
      card.title = 'Could not locate this text on the page';
      setTimeout(() => { card.title = was; }, 1500);
    }
  };
  return card;
}

/* ── Feature 2: the Assistant (vibe prompting) ─────────────────────── */

const CALCPAD_GUIDE = [
  'CalcPad is a live calculation sheet. Each line of the document is either prose or',
  'one of these:',
  '',
  '  name = value unit        a definition, e.g.  L = 6 m     f_ck = 30 MPa',
  '  name = expression        computed, e.g.      A = b*h',
  '  = expression             a standalone formula, e.g.  =w*L^2/8',
  '  Span L = 6 m             text before the name is kept as a label',
  '  b = 200 mm; h = 400 mm   several definitions on one line, split by ";"',
  '',
  'Output units go in square brackets at the end of the line:',
  '  M_Ed = w*L^2/8 [kN·m]    show the answer in kN·m; the dimension must agree',
  '  A_s = ... [mm2/m]        kept as written rather than reduced',
  '  k = 1+sqrt(f_ck/200) [#MPa]   "#" forces the unit with no conversion and no',
  '                                dimension check — for empirical code formulas',
  '  x = expr [#]             strip the unit, leave a plain number',
  '  f_ck[#MPa]               inside a formula: that term as a plain number in MPa.',
  '                           Strip every term in a formula or none of them.',
  '',
  'Names: letters, digits and underscores, starting with a letter. The underscore is a',
  'subscript, so M_Ed renders as M with subscript Ed. A name that is a Greek letter',
  'renders as one: gamma_c → γ_c, and Sigma → Σ. Use the conventional Eurocode symbols.',
  '',
  'Units understood: mm cm m km · N kN MN · g kg t · Pa kPa MPa GPa · s ms min hr ·',
  'deg rad. Compound units are written with * and / and ^, e.g. kN/m^2, kN·m, mm^4.',
  '',
  'Functions: SUM AVG MIN MAX COUNT SQRT ABS ROUND(x,d) POW IF(cond,t,f) LOG(x[,base])',
  'EXP PERCENT PI() E() FLOOR CEIL MOD SIN COS TAN ASIN ACOS ATAN ATAN2 SINH COSH TANH',
  'DEG RAD. Angles are radians — write 30deg for degrees. Comparison operators work',
  'inside IF. x = solver(a == b) solves for x, and must be the whole definition.',
  '',
  'Rules that matter:',
  '- A name may be defined once only. Never redefine a name that already exists in the',
  '  document; pick a distinct one (M_Ed_1, M_Ed_2) or reuse the existing value.',
  '- Define a quantity before the line that uses it.',
  '- Do not write the answer yourself: give the expression and let CalcPad evaluate it.',
  '  Write "M_Ed = w*L^2/8 [kN·m]", never "M_Ed = 45 kN·m" for something derived.',
  '- Plain sentences between the definitions are encouraged — they are what makes the',
  '  sheet readable. Reference the code clause you are working to.',
].join('\n');

const CHAT_SYSTEM = [
  'You are the CalcPad assistant, helping a structural engineer draft and extend an',
  'engineering calculation sheet. You write in British English, in the plain, direct',
  'prose an issued calculation uses.',
  '',
  CALCPAD_GUIDE,
  '',
  'HOW TO ANSWER',
  '',
  'When the engineer asks for calculation content, put the lines to be inserted in a',
  'fenced block tagged calcpad:',
  '',
  '```calcpad',
  'Design of a simply supported beam to EN 1992-1-1.',
  'Span L = 6 m',
  'Uniformly distributed load w = 12 kN/m',
  'Design moment M_Ed = w*L^2/8 [kN·m]',
  '```',
  '',
  'Everything inside that block is inserted into the sheet verbatim, one document line',
  'per line, so it must be valid CalcPad and nothing else — no markdown, no bullets, no',
  'commentary, no ``` inside it. Put your explanation outside the block.',
  '',
  'Use several blocks when the work falls into distinct sections, so the engineer can',
  'insert them one at a time. Keep each block to what was asked for.',
  '',
  'Before the block, say in a sentence or two what you are calculating and on what',
  'basis. After it, note any assumption you had to make and anything the engineer must',
  'check or confirm. Where a design code governs, name the clause.',
  '',
  'You can see the current state of the sheet, including every variable already',
  'defined. Build on it: reuse those variables rather than redefining them, follow the',
  'symbols and units already in use, and continue from where the document has got to.',
  '',
  'If the request is ambiguous in a way that changes the answer — the code, the exposure',
  'class, the support conditions — state the assumption you are making and carry on.',
  'Ask a question only when no reasonable assumption is available. Never present a',
  'result as checked or compliant; you are drafting, the engineer is responsible.',
].join('\n');

function chatContextBlock() {
  return 'CURRENT STATE OF THE SHEET\n\n' + documentContext(false);
}

async function sendChat() {
  const box = $('aiChatInput');
  const prompt = box.value.trim();
  if (!prompt) return;
  if (!ensureKey(sendChat)) return;

  box.value = '';
  state.chat.push({ role: 'user', content: prompt });
  addChatBubble('user', prompt);

  const btn = $('aiChatSend');
  btn.disabled = true;
  $('aiChatStop').style.display = '';
  setStatus('aiChatStatus', 'Thinking…', true);

  const bubble = addChatBubble('assistant', '');
  const body = bubble.querySelector('.ai-msg-body');
  body.className = 'ai-msg-body ai-stream';

  let answer = '', thinking = '';

  // The sheet moves between turns, so the live state rides along with the
  // newest question rather than being frozen into the system prompt.
  const withContext = state.chat.map((m, i) => {
    if (m.role === 'user' && i === state.chat.length - 1 && $('aiChatCtx').checked)
      return { role: 'user', content: chatContextBlock() + '\n\n---\n\n' + m.content };
    return m;
  });

  try {
    const { usage } = await streamMessage(buildRequest({
      // A cache breakpoint on the guide: it is identical on every turn, so
      // repeat questions in a session re-read it instead of re-sending it.
      system: [{ type: 'text', text: CHAT_SYSTEM, cache_control: { type: 'ephemeral' } }],
      messages: withContext,
      maxTokens: 32000,
    }), {
      thinking: t => {
        thinking += t;
        setStatus('aiChatStatus', thinking.slice(-140).replace(/\s+/g, ' '), true);
      },
      text: t => {
        answer += t;
        body.textContent = answer;
        scrollChat();
      },
    });
    showUsage(usage);

    state.chat.push({ role: 'assistant', content: answer });
    body.className = 'ai-msg-body';
    renderRich(body, answer);
    setStatus('aiChatStatus', '', false);
    scrollChat();
  } catch (e) {
    if (e.name === 'AbortError') {
      // Keep the partial answer out of the history — a truncated turn would
      // mislead the next one.
      state.chat.pop();
      bubble.remove();
      setStatus('aiChatStatus', 'Stopped.', false);
    } else {
      state.chat.pop();
      bubble.remove();
      $('aiChatLog').appendChild(el('div', 'ai-err', e.message));
      setStatus('aiChatStatus', '', false);
      scrollChat();
    }
  } finally {
    btn.disabled = false;
    $('aiChatStop').style.display = 'none';
  }
}

function addChatBubble(who, text) {
  const wrap = el('div', 'ai-msg ' + (who === 'user' ? 'ai-msg-user' : 'ai-msg-ai'));
  wrap.appendChild(el('div', 'ai-msg-who', who === 'user' ? 'You' : 'CalcPad AI'));
  const body = el('div', 'ai-msg-body');
  if (who === 'user') body.textContent = text; else body.className = 'ai-msg-body ai-stream';
  wrap.appendChild(body);
  $('aiChatLog').appendChild(wrap);
  scrollChat();
  return wrap;
}

function scrollChat() {
  const v = $('aiChatView');
  v.scrollTop = v.scrollHeight;
}

/* ── Rendering a reply safely ──────────────────────────────────────── */

// Model output is untrusted text. Everything is escaped first and only then
// given the handful of tags below, so nothing in a reply can execute.
function inlineMd(s) {
  return esc(s)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>');
}

function renderRich(host, src) {
  host.innerHTML = '';
  // Splitting on a capturing group interleaves the captures with the text:
  // [text, tag, text, tag, text, …]. Even indices are the text between fences
  // and alternate prose, code, prose, …; odd indices are the language tag that
  // opened the block that follows.
  const parts = String(src).split(/```([a-zA-Z]*)\n?/);
  for (let i = 0; i < parts.length; i += 2) {
    if ((i / 2) % 2 === 0) renderProse(host, parts[i]);
    else renderCode(host, parts[i], (parts[i - 1] || '').toLowerCase());
  }
}

function renderProse(host, text) {
  const lines = String(text).split('\n');
  let para = [], list = null;

  const flushPara = () => {
    if (!para.length) return;
    const p = el('p');
    p.innerHTML = inlineMd(para.join(' '));
    host.appendChild(p);
    para = [];
  };
  const flushList = () => { list = null; };

  lines.forEach(raw => {
    const line = raw.trim();
    if (!line) { flushPara(); flushList(); return; }

    const h = line.match(/^(#{1,6})\s+(.*)$/);
    if (h) {
      flushPara(); flushList();
      const tag = el(h[1].length <= 3 ? 'h3' : 'h4');
      tag.innerHTML = inlineMd(h[2]);
      host.appendChild(tag);
      return;
    }

    const b = line.match(/^[-*•]\s+(.*)$/);
    const o = line.match(/^\d+[.)]\s+(.*)$/);
    if (b || o) {
      flushPara();
      const want = b ? 'UL' : 'OL';
      if (!list || list.tagName !== want) { list = el(b ? 'ul' : 'ol'); host.appendChild(list); }
      const li = el('li');
      li.innerHTML = inlineMd((b || o)[1]);
      list.appendChild(li);
      return;
    }

    flushList();
    para.push(line);
  });
  flushPara();
}

function renderCode(host, code, lang) {
  const text = String(code).replace(/\n+$/, '');
  if (!text.trim()) return;
  const isCalc = lang === 'calcpad' || lang === 'calc';

  const box = el('div', 'ai-code');
  const head = el('div', 'ai-code-head');
  head.appendChild(el('span', 'ai-code-tag', isCalc ? 'CalcPad lines' : (lang || 'code')));

  const btns = el('div', 'ai-code-btns');
  if (isCalc) {
    const ins = el('button', 'ai-code-btn', 'Insert at cursor');
    ins.onclick = () => {
      if (insertCalcLines(text)) {
        ins.textContent = 'Inserted';
        setTimeout(() => { ins.textContent = 'Insert at cursor'; }, 1600);
      }
    };
    btns.appendChild(ins);
  }
  const cp = el('button', 'ai-code-btn', 'Copy');
  cp.onclick = () => {
    navigator.clipboard.writeText(text).then(
      () => { cp.textContent = 'Copied'; setTimeout(() => { cp.textContent = 'Copy'; }, 1400); },
      () => { cp.textContent = 'Copy failed'; }
    );
  };
  btns.appendChild(cp);
  head.appendChild(btns);

  const pre = el('pre', null, text);
  box.appendChild(head);
  box.appendChild(pre);
  host.appendChild(box);
}

/* ── API key handling ──────────────────────────────────────────────── */

let _afterKey = null;

// Every entry point routes through here, so a run can be started without a key
// and simply resumes once one is given.
function ensureKey(then) {
  if (getKey()) return true;
  _afterKey = then || null;
  openKeyModal();
  return false;
}

function openKeyModal() {
  $('aiKeyInput').value = getKey();
  $('aiKeyOverlay').classList.add('open');
  setTimeout(() => $('aiKeyInput').focus(), 30);
}

function closeKeyModal() {
  $('aiKeyOverlay').classList.remove('open');
  _afterKey = null;
}

function saveKeyFromModal() {
  const v = $('aiKeyInput').value.trim();
  const msg = $('aiKeyMsg');
  if (!v) { msg.textContent = 'Paste a key, or press Cancel.'; return; }
  if (!/^sk-ant-/.test(v)) {
    msg.textContent = 'That does not look like an Anthropic key (they start "sk-ant-"). Saving anyway.';
  }
  setKey(v);
  const then = _afterKey;
  $('aiKeyOverlay').classList.remove('open');
  _afterKey = null;
  if (then) then();
}

function forgetKey() {
  setKey('');
  setStatus('aiQAStatus', 'Key cleared for this tab.', false);
}

function syncKeyState() {
  const has = !!getKey();
  const dot = $('aiKeyDot'), txt = $('aiKeyTxt');
  if (dot) dot.className = 'ai-key-dot' + (has ? ' ok' : '');
  if (txt) txt.textContent = has ? 'Key set for this tab' : 'No API key set';
}

/* ── Panel plumbing ────────────────────────────────────────────────── */

function setStatus(id, msg, busy) {
  const n = $(id);
  if (!n) return;
  n.innerHTML = '';
  if (busy) n.appendChild(el('div', 'ai-spin'));
  if (msg) n.appendChild(el('span', null, msg));
}

function openPanel(tab) {
  $('aiPanel').classList.add('open');
  document.body.classList.add('ai-open');
  syncKeyState();
  if (tab) showTab(tab);
}

function closePanel() {
  $('aiPanel').classList.remove('open');
  document.body.classList.remove('ai-open');
  clearHighlights();
}

function showTab(tab) {
  ['qa', 'chat'].forEach(t => {
    $('aiTab_' + t).classList.toggle('on', t === tab);
    $('ai' + (t === 'qa' ? 'QA' : 'Chat') + 'View').classList.toggle('on', t === tab);
  });
  $('aiCompose').style.display = tab === 'chat' ? '' : 'none';
  if (tab === 'chat') setTimeout(() => $('aiChatInput').focus(), 40);
}

function stopRun() {
  if (state.running) state.running.abort();
}

/* ── Building the UI ───────────────────────────────────────────────── */

const SPARK = '<svg viewBox="0 0 24 24"><path d="M12 2l2.2 6.1L20.5 10l-6.3 1.9L12 18l-2.2-6.1L3.5 10l6.3-1.9z"/><path d="M18.5 14l.9 2.4 2.6.8-2.6.8-.9 2.4-.9-2.4-2.6-.8 2.6-.8z"/></svg>';

function modelOptions() {
  return Object.keys(MODELS)
    .map(id => '<option value="' + esc(id) + '">' + esc(MODELS[id].label) + '</option>')
    .join('');
}

function buildUI() {
  // ── Toolbar buttons, in front of the syntax hints ──
  const bar = document.querySelector('.toolbar');
  const hints = bar && bar.querySelector('.hints');
  if (bar) {
    const sep = el('div', 'sep');
    const qa = el('button', 'ai-tbtn');
    qa.innerHTML = SPARK + '<span>AI QA</span>';
    qa.title = 'Check the document for flow, spelling and grammar';
    qa.onclick = () => openPanel('qa');

    const chat = el('button', 'ai-tbtn');
    chat.innerHTML = SPARK + '<span>Assistant</span>';
    chat.title = 'Describe a calculation and have it drafted';
    chat.onclick = () => openPanel('chat');

    if (hints) { bar.insertBefore(sep, hints); bar.insertBefore(qa, hints); bar.insertBefore(chat, hints); }
    else { bar.appendChild(sep); bar.appendChild(qa); bar.appendChild(chat); }
  }

  // ── Beta marker in the alpha banner ──
  const banner = $('alphaBanner');
  if (banner) banner.insertBefore(el('span', 'ai-beta-pill', 'BETA'), banner.firstChild);

  // ── The panel ──
  const panel = el('div', 'ai-panel');
  panel.id = 'aiPanel';
  panel.innerHTML =
    '<div class="ai-head">' +
      '<span class="ai-head-title">' + SPARK + 'CalcPad AI<span class="ai-head-badge">BETA</span></span>' +
      '<button class="ai-close" id="aiClose" title="Close">&times;</button>' +
    '</div>' +

    '<div class="ai-tabs">' +
      '<button class="ai-tab on" id="aiTab_qa">QA review</button>' +
      '<button class="ai-tab" id="aiTab_chat">Assistant</button>' +
    '</div>' +

    // ── QA view ──
    '<div class="ai-view on" id="aiQAView">' +
      '<div class="ai-note">Reads the whole sheet and reports on how it <b>flows</b>, ' +
        'plus <b>spelling</b> and <b>grammar</b>. Click a finding to highlight it in the document. ' +
        'Nothing is changed for you.</div>' +
      '<div class="ai-row">' +
        '<button class="ai-btn" id="aiQARun">Run QA review</button>' +
        '<button class="ai-btn-ghost" id="aiQAStop" style="display:none">Stop</button>' +
      '</div>' +
      '<div class="ai-status" id="aiQAStatus"></div>' +
      '<div class="ai-think" id="aiQAThink"></div>' +
      '<div id="aiQAResults" style="display:flex;flex-direction:column;gap:8px"></div>' +
    '</div>' +

    // ── Chat view ──
    '<div class="ai-view" id="aiChatView">' +
      '<div class="ai-note" id="aiChatHint">Describe the calculation you want. ' +
        'Replies come back as CalcPad lines you can drop straight into the sheet — ' +
        'review them before you do. <br>Try: <code>a simply supported RC beam, 6 m span, ' +
        '12 kN/m, check bending to EC2</code></div>' +
      '<div class="ai-chat-log" id="aiChatLog"></div>' +
    '</div>' +

    '<div class="ai-compose" id="aiCompose" style="display:none">' +
      '<textarea id="aiChatInput" placeholder="Describe the calculation, or ask for a change…"></textarea>' +
      '<div class="ai-status" id="aiChatStatus"></div>' +
      '<div class="ai-compose-row">' +
        '<label class="ai-chk"><input type="checkbox" id="aiChatCtx" checked> Send the sheet as context</label>' +
        '<button class="ai-btn-ghost" id="aiChatStop" style="display:none">Stop</button>' +
        '<button class="ai-btn ai-send" id="aiChatSend">Send</button>' +
      '</div>' +
    '</div>' +

    '<div class="ai-foot">' +
      '<span class="ai-key-dot" id="aiKeyDot"></span>' +
      '<span id="aiKeyTxt">No API key set</span>' +
      '<button class="ai-btn-ghost" id="aiKeyBtn" style="padding:2px 8px">Key…</button>' +
      '<select class="ai-sel" id="aiModel" title="Model">' + modelOptions() + '</select>' +
      '<select class="ai-sel" id="aiEffort" title="How hard the model works">' +
        '<option value="low">Quick</option>' +
        '<option value="high" selected>Standard</option>' +
        '<option value="max">Thorough</option>' +
      '</select>' +
      '<span class="ai-usage" id="aiUsage"></span>' +
    '</div>';
  document.body.appendChild(panel);

  // ── Key modal ──
  const ov = el('div', 'ai-overlay');
  ov.id = 'aiKeyOverlay';
  ov.innerHTML =
    '<div class="ai-modal">' +
      '<h3>Anthropic API key</h3>' +
      '<div class="ai-warn"><b>Beta testing only.</b> The key is held in this browser tab ' +
        'for this session and sent straight from your browser to api.anthropic.com. It is ' +
        'never saved into your document, an autosave, or a <code>.perega</code> file, and it ' +
        'is gone when you close the tab. Use a key you are happy to rotate, and set a spend ' +
        'limit on it.</div>' +
      '<input type="password" id="aiKeyInput" placeholder="sk-ant-…" autocomplete="off" spellcheck="false">' +
      '<div class="ai-note" id="aiKeyMsg" style="min-height:15px"></div>' +
      '<div class="ai-modal-acts">' +
        '<button class="ai-btn-ghost" id="aiKeyForget">Forget key</button>' +
        '<button class="ai-btn-ghost" id="aiKeyCancel">Cancel</button>' +
        '<button class="ai-btn" id="aiKeySave">Save for this session</button>' +
      '</div>' +
    '</div>';
  document.body.appendChild(ov);
}

function wireUI() {
  $('aiClose').onclick    = closePanel;
  $('aiTab_qa').onclick   = () => showTab('qa');
  $('aiTab_chat').onclick = () => showTab('chat');

  $('aiQARun').onclick  = runQA;
  $('aiQAStop').onclick = stopRun;

  $('aiChatSend').onclick = sendChat;
  $('aiChatStop').onclick = stopRun;
  $('aiChatInput').addEventListener('keydown', e => {
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); sendChat(); }
  });

  $('aiKeyBtn').onclick    = openKeyModal;
  $('aiKeySave').onclick   = saveKeyFromModal;
  $('aiKeyCancel').onclick = closeKeyModal;
  $('aiKeyForget').onclick = () => { forgetKey(); $('aiKeyInput').value = ''; $('aiKeyMsg').textContent = 'Cleared.'; };
  $('aiKeyInput').addEventListener('keydown', e => { if (e.key === 'Enter') saveKeyFromModal(); });

  $('aiModel').onchange  = e => { state.model = e.target.value; syncEffortEnabled(); };
  $('aiEffort').onchange = e => { state.effort = e.target.value; };

  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    if ($('aiKeyOverlay').classList.contains('open')) closeKeyModal();
    else if ($('aiPanel').classList.contains('open')) closePanel();
  });

  syncEffortEnabled();
  syncKeyState();
}

// Haiku 4.5 rejects an effort setting, so the control goes away rather than
// producing a 400 the tester has to decode.
function syncEffortEnabled() {
  const caps = MODELS[state.model] || {};
  $('aiEffort').disabled = !caps.effort;
  $('aiEffort').style.opacity = caps.effort ? '' : '.45';
}

/* ── Start ─────────────────────────────────────────────────────────── */

function init() {
  if ($('aiPanel')) return;          // already initialised
  buildUI();
  wireUI();
  // Highlights are positioned in page coordinates, so they stay put as the
  // page scrolls and only a reflow can strand them. Clearing on scroll would
  // wipe the highlight the moment flashRange() scrolled it into view.
  window.addEventListener('resize', clearHighlights);
  console.log('CalcPad Beta — AI features loaded');
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

})();

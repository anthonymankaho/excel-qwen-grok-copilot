const API = '';
const CONTEXT_CAP = 45_000; // Relay rejects messages over 50,000 chars; leave room for the prompt itself.
const KEY_STORE = 'eqgc.keys.v1';
const state = { messages: [], lastAnswer: '', busy: false, officeReady: false, forceSheet: false };
const $ = id => document.getElementById(id);

function storedKeys() {
  try { return JSON.parse(localStorage.getItem(KEY_STORE) || '{}'); } catch { return {}; }
}
function rememberKey(provider, key) {
  const all = storedKeys();
  if (key) all[provider] = key; else delete all[provider];
  try { localStorage.setItem(KEY_STORE, JSON.stringify(all)); } catch { /* storage unavailable: session-only */ }
}
let keysReady = null;
function ensureKeys() {
  if (!keysReady) keysReady = (async () => {
    const stored = storedKeys();
    let allOk = true;
    for (const provider of ['qwen', 'grok']) {
      if (stored[provider]) { try { await api('configure', { provider, apiKey: stored[provider] }); } catch { allOk = false; } }
    }
    if (!allOk) keysReady = null; // relay was offline; retry on next question
  })();
  return keysReady;
}

async function api(path, body) {
  const response = await fetch(`${API}/api/${path}`, { method: body ? 'POST' : 'GET', headers: body ? { 'Content-Type': 'application/json' } : {}, body: body ? JSON.stringify(body) : undefined });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}

function toast(message) {
  const el = $('toast'); el.textContent = message; el.classList.remove('hidden');
  clearTimeout(toast.timer); toast.timer = setTimeout(() => el.classList.add('hidden'), 3200);
}

async function readSelection() {
  return Excel.run(async context => {
    const range = context.workbook.getSelectedRange();
    range.load(['address', 'text', 'formulas', 'rowCount', 'columnCount']);
    await context.sync();
    return { address: range.address, text: range.text, formulas: range.formulas, rowCount: range.rowCount, columnCount: range.columnCount };
  });
}

async function readUsedRange() {
  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRangeOrNullObject();
    range.load(['address', 'text', 'formulas']);
    await context.sync();
    return range.isNullObject ? null : { address: range.address, text: range.text, formulas: range.formulas };
  });
}

function annotateRows(rows, formulas) {
  return (rows || []).map((row, i) => row.map((value, j) => {
    const formula = formulas?.[i]?.[j];
    return formula && String(formula).startsWith('=') ? `${value} [${formula}]` : value;
  }).join('\t')).join('\n');
}

async function buildContext() {
  if (!state.officeReady || !$('includeContext').checked) return null;
  let source = null, kind = 'selection';
  if (!state.forceSheet) {
    const selection = await readSelection();
    if (selection.rowCount * selection.columnCount > 1) source = selection;
  }
  if (!source) { source = await readUsedRange(); kind = 'used range'; }
  if (!source) return null;
  let text = annotateRows(source.text, source.formulas);
  if (!text) return null;
  let truncated = false;
  if (text.length > CONTEXT_CAP) {
    const marker = `\n[...truncated: ${text.length - CONTEXT_CAP} chars omitted — ask about a smaller range for full detail]`;
    text = text.slice(0, CONTEXT_CAP - marker.length) + marker;
    truncated = true;
  }
  return { header: `[Excel ${kind} ${source.address}]`, text, address: source.address, kind, truncated, chars: text.length };
}

async function refreshContextCard() {
  if (!state.officeReady) { $('selectionName').textContent = 'Open Excel to read context'; return; }
  try {
    const ctx = await buildContext();
    if (!$('includeContext').checked) {
      $('selectionName').textContent = 'Context off';
      $('selectionPreview').textContent = 'Questions are sent without workbook data.';
      $('contextHint').textContent = 'Nothing will be sent';
      return;
    }
    if (!ctx) {
      $('selectionName').textContent = 'No workbook data';
      $('selectionPreview').textContent = '(The active sheet and selection are empty)';
      $('contextHint').textContent = 'Nothing will be sent';
      return;
    }
    $('selectionName').textContent = ctx.address;
    $('selectionPreview').textContent = ctx.text;
    $('contextHint').textContent = `Next question sends ${ctx.kind} · ${ctx.chars.toLocaleString()} chars${ctx.truncated ? ' · truncated' : ''} · formulas included`;
  } catch (error) { toast(`Could not read workbook context: ${error.message}`); }
}

function renderMessage(role, content) {
  const conversation = $('conversation');
  conversation.querySelector('.welcome')?.remove();
  const bubble = document.createElement('div'); bubble.className = `message ${role}`;
  if (role === 'assistant') {
    const label = document.createElement('div'); label.className = 'message-label'; label.textContent = $('activeProvider').value.toUpperCase(); bubble.append(label);
  }
  const body = document.createElement('div'); body.textContent = content; bubble.append(body); conversation.append(bubble); conversation.scrollTop = conversation.scrollHeight;
  return bubble;
}

function parseEdits(text) {
  const blocks = [...text.matchAll(/```json\s*([\s\S]*?)```/g)];
  for (let i = blocks.length - 1; i >= 0; i--) {
    try {
      const parsed = JSON.parse(blocks[i][1]);
      const edits = parsed?.edits;
      if (!Array.isArray(edits) || edits.length < 1 || edits.length > 200) continue;
      if (!edits.every(e => e && typeof e.sheet === 'string' && e.sheet.length <= 100 && typeof e.cell === 'string' && /^[A-Z]{1,3}[1-9][0-9]{0,6}$/i.test(e.cell) && (typeof e.value === 'string' || typeof e.value === 'number') && String(e.value).length <= 5000)) continue;
      const prose = (text.slice(0, blocks[i].index) + text.slice(blocks[i].index + blocks[i][0].length)).trim();
      return { prose, edits };
    } catch { /* not an edits block; keep looking */ }
  }
  return { prose: text, edits: null };
}

async function renderEditsCard(edits) {
  const conversation = $('conversation');
  const card = document.createElement('div'); card.className = 'message assistant edits-card';
  const label = document.createElement('div'); label.className = 'message-label'; label.textContent = `PROPOSED EDITS · ${edits.length} · REVIEW OLD → NEW · APPLY ONLY WITH YOUR CLICK`;
  card.append(label);
  let olds = null;
  try {
    olds = await Excel.run(async context => {
      const ranges = edits.map(e => context.workbook.worksheets.getItem(e.sheet).getRange(e.cell));
      ranges.forEach(r => r.load(['values', 'formulas']));
      await context.sync();
      return ranges.map(r => ({ value: r.values?.[0]?.[0], formula: r.formulas?.[0]?.[0] }));
    });
  } catch { olds = null; }
  edits.forEach((edit, i) => {
    const row = document.createElement('div'); row.className = 'edit-row';
    const old = olds?.[i];
    const oldText = !old ? '(current unreadable)'
      : String(old.formula || '').startsWith('=') ? old.formula
      : (old.value === '' || old.value == null ? '(empty)' : String(old.value));
    row.textContent = `${edit.sheet}!${edit.cell}: ${oldText} → ${String(edit.value)}`;
    card.append(row);
  });
  const actions = document.createElement('div'); actions.className = 'button-row';
  const apply = document.createElement('button'); apply.className = 'primary'; apply.textContent = 'Apply edits';
  const dismiss = document.createElement('button'); dismiss.className = 'secondary'; dismiss.textContent = 'Dismiss';
  apply.addEventListener('click', async () => {
    apply.disabled = true; dismiss.disabled = true;
    try {
      const recs = await Excel.run(async context => {
        const ranges = edits.map(e => context.workbook.worksheets.getItem(e.sheet).getRange(e.cell));
        ranges.forEach(r => r.load(['values', 'formulas']));
        await context.sync();
        const before = ranges.map((r, i) => ({ sheet: edits[i].sheet, cell: edits[i].cell, value: r.values?.[0]?.[0], formula: r.formulas?.[0]?.[0] }));
        ranges.forEach((r, i) => {
          const v = edits[i].value;
          if (typeof v === 'string' && v.startsWith('=')) r.formulas = [[v]];
          else r.values = [[v]];
        });
        await context.sync();
        return before;
      });
      label.textContent = `APPLIED ${edits.length} EDITS`;
      const undoBtn = document.createElement('button'); undoBtn.className = 'secondary'; undoBtn.textContent = 'Undo last apply';
      undoBtn.addEventListener('click', async () => {
        undoBtn.disabled = true;
        try {
          await Excel.run(async context => {
            for (const rec of recs) {
              const r = context.workbook.worksheets.getItem(rec.sheet).getRange(rec.cell);
              if (String(rec.formula || '').startsWith('=')) r.formulas = [[rec.formula]];
              else if (rec.value === '' || rec.value == null) r.clear();
              else r.values = [[rec.value]];
            }
            await context.sync();
          });
          label.textContent = `UNDONE — RESTORED ${recs.length} CELLS`;
          toast('Restored previous cell values.');
        } catch (error) { undoBtn.disabled = false; toast(`Could not undo: ${error.message}`); }
      });
      actions.append(undoBtn);
      toast(`Applied ${edits.length} edits to the workbook.`);
    } catch (error) {
      apply.disabled = false; dismiss.disabled = false;
      toast(`Could not apply edits: ${error.message}`);
    }
  });
  dismiss.addEventListener('click', () => card.remove());
  actions.append(apply, dismiss); card.append(actions);
  conversation.append(card); conversation.scrollTop = conversation.scrollHeight;
}

function dedupContext(messages) {
  // Older turns keep their prose only; re-sending every past 45k context block would multiply input tokens each turn.
  return messages.map((m, i) => (m.role === 'user' && i < messages.length - 1)
    ? { ...m, content: m.content.replace(/\n\n\[Excel (?:selection|used range) [^\]]*\]\n[\s\S]*$/, '') }
    : m);
}

async function streamChat(payload) {
  const response = await fetch(`${API}/api/chat-stream`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || `Request failed (${response.status})`);
  }
  const bubble = renderMessage('assistant', '');
  const bodyEl = bubble.children[bubble.children.length - 1]; // HTMLCollection has no .at()
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buf = '', text = '', model = '';
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      let idx;
      while ((idx = buf.indexOf('\n\n')) >= 0) {
        const frame = buf.slice(0, idx); buf = buf.slice(idx + 2);
        for (const line of frame.split('\n')) {
          if (!line.startsWith('data: ')) continue;
          const raw = line.slice(6);
          if (raw === '[DONE]') continue;
          let j; try { j = JSON.parse(raw); } catch { continue; }
          if (j.error) throw new Error(j.error);
          if (j.model) model = j.model;
          if (j.status === 'thinking' && !text) { bodyEl.textContent = 'Thinking…'; continue; }
          const delta = j.choices?.[0]?.delta?.content;
          if (delta) {
            text += delta; bodyEl.textContent = text;
            $('conversation').scrollTop = $('conversation').scrollHeight;
          }
        }
      }
    }
  } catch (error) { error.bubble = bubble; throw error; }
  return { text, model, bubble };
}

async function sendPrompt(prompt) {
  const clean = prompt.trim(); if (!clean || state.busy) return;
  await ensureKeys();
  const provider = $('activeProvider').value;
  let ctx = null;
  try { ctx = await buildContext(); } catch (error) { toast(`Could not read workbook context: ${error.message}`); }
  state.forceSheet = false;
  const content = clean + (ctx ? `\n\n${ctx.header}\n${ctx.text}` : '');
  if (content.length > 49_500) { toast('Message plus workbook context is too large — shorten the message or turn context off.'); return; }
  state.messages.push({ role: 'user', content }); renderMessage('user', clean + (ctx ? `\n\nContext: ${ctx.address}${ctx.truncated ? ' (truncated)' : ''}` : ''));
  state.busy = true; $('send').disabled = true; $('send').textContent = '…';
  try {
    const result = await streamChat({ provider, model: provider === 'qwen' ? $('qwenModel').value.trim() : $('grokModel').value.trim(), endpoint: $('qwenEndpoint').value.trim(), messages: dedupContext(state.messages.slice(-24)), fast: $('fastMode').checked });
    const { prose, edits } = parseEdits(result.text);
    state.messages.push({ role: 'assistant', content: result.text });
    state.lastAnswer = prose; $('insertAnswer').disabled = !prose; $('lastAnswerState').textContent = result.model ? `via ${result.model}` : '';
    result.bubble.children[result.bubble.children.length - 1].textContent = prose;
    if (edits) await renderEditsCard(edits);
    refreshContextCard();
  } catch (error) {
    if (error.bubble) error.bubble.children[error.bubble.children.length - 1].textContent = `Error: ${error.message}`;
    else renderMessage('assistant', `Error: ${error.message}`);
    toast(error.message);
  }
  finally { state.busy = false; $('send').disabled = false; $('send').textContent = '↑'; }
}

async function refreshSettings() {
  try {
    const data = await api('health');
    const provider = $('provider').value;
    const remembered = Boolean(storedKeys()[provider]);
    $('keyStatus').textContent = data.configured[provider] ? (remembered ? 'Key active · remembered on this PC' : 'Key active for this session') : '';
    $('keyStatus').style.color = data.configured[provider] ? '#168251' : '#777';
  } catch { $('keyStatus').textContent = 'Relay unavailable'; $('keyStatus').style.color = '#b23b36'; }
}

function toggleProviderFields() {
  const isQwen = $('provider').value === 'qwen';
  $('qwenFields').classList.toggle('hidden', !isQwen); $('grokFields').classList.toggle('hidden', isQwen);
  $('activeProvider').value = $('provider').value;
  refreshSettings();
}

async function insertAnswer() {
  if (!state.lastAnswer || !state.officeReady) return;
  try {
    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange().getCell(0, 0); range.values = [[state.lastAnswer]]; await context.sync();
    });
    toast('Answer inserted into the selected cell.');
  } catch (error) { toast(`Could not insert answer: ${error.message}`); }
}

$('settingsToggle').addEventListener('click', () => { $('settings').classList.toggle('hidden'); refreshSettings(); });
$('closeSettings').addEventListener('click', () => $('settings').classList.add('hidden'));
$('provider').addEventListener('change', toggleProviderFields);
$('activeProvider').addEventListener('change', () => { $('provider').value = $('activeProvider').value; toggleProviderFields(); });
$('refreshSelection').addEventListener('click', refreshContextCard);
$('includeContext').addEventListener('change', refreshContextCard);
$('addUsedRange').addEventListener('click', () => { state.forceSheet = true; refreshContextCard(); toast('Next question will use the whole active sheet.'); });
$('saveKey').addEventListener('click', async () => {
  const apiKey = $('apiKey').value.trim(); if (!apiKey) return toast('Enter an API key first.');
  const provider = $('provider').value;
  try {
    await api('configure', { provider, apiKey });
    rememberKey(provider, $('rememberKey').checked ? apiKey : '');
    keysReady = null;
    $('apiKey').value = ''; await refreshSettings();
    toast($('rememberKey').checked ? 'Key loaded and remembered on this PC.' : 'Key loaded into local relay memory.');
  }
  catch (error) { toast(error.message); }
});
$('clearKey').addEventListener('click', async () => {
  const provider = $('provider').value;
  try { await api('clear-key', { provider }); rememberKey(provider, ''); keysReady = null; await refreshSettings(); toast('Key cleared.'); }
  catch (error) { toast(error.message); }
});
$('composer').addEventListener('submit', async event => { event.preventDefault(); const value = $('prompt').value; $('prompt').value = ''; await sendPrompt(value); });
$('prompt').addEventListener('keydown', event => { if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); $('composer').requestSubmit(); } });
$('insertAnswer').addEventListener('click', insertAnswer);
document.querySelectorAll('.action').forEach(button => button.addEventListener('click', () => { $('prompt').value = button.dataset.prompt; $('prompt').focus(); }));

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) { state.officeReady = true; refreshContextCard(); }
  refreshSettings();
  ensureKeys().then(refreshSettings);
});

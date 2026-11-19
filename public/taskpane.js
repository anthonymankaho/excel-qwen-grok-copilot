const API = '';
const state = { messages: [], sharedSelection: '', lastAnswer: '', busy: false, officeReady: false };
const $ = id => document.getElementById(id);

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

function safeText(node, value) { node.textContent = value; }

async function readSelection() {
  if (!state.officeReady) { toast('Open this task pane inside Excel first.'); return; }
  try {
    const selection = await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'text', 'formulas']);
      await context.sync();
      return { address: range.address, text: range.text, formulas: range.formulas };
    });
    $('selectionName').textContent = selection.address || 'Selected cells';
    const rows = (selection.text || []).map((row, i) => row.map((value, j) => {
      const formula = selection.formulas?.[i]?.[j];
      return formula && String(formula).startsWith('=') ? `${value} [${formula}]` : value;
    }).join('\t')).join('\n');
    $('selectionPreview').textContent = rows || '(Selection is empty)';
    state.currentSelection = { address: selection.address, rows: selection.text || [], formulas: selection.formulas || [] };
  } catch (error) { toast(`Could not read selection: ${error.message}`); }
}

function renderMessage(role, content) {
  const conversation = $('conversation');
  conversation.querySelector('.welcome')?.remove();
  const bubble = document.createElement('div'); bubble.className = `message ${role}`;
  if (role === 'assistant') {
    const label = document.createElement('div'); label.className = 'message-label'; label.textContent = $('activeProvider').value.toUpperCase(); bubble.append(label);
  }
  const body = document.createElement('div'); body.textContent = content; bubble.append(body); conversation.append(bubble); conversation.scrollTop = conversation.scrollHeight;
}

function selectionContext() {
  if (!state.sharedSelection || !state.currentSelection) return '';
  return `\n\n[Excel selection ${state.currentSelection.address}]\n${state.sharedSelection}`;
}

async function sendPrompt(prompt) {
  const clean = prompt.trim(); if (!clean || state.busy) return;
  const provider = $('activeProvider').value;
  const content = clean + selectionContext();
  state.messages.push({ role: 'user', content }); renderMessage('user', clean + (state.sharedSelection ? `\n\nAttached selection: ${state.currentSelection.address}` : ''));
  state.busy = true; $('send').disabled = true; $('send').textContent = '…';
  try {
    const result = await api('chat', { provider, model: provider === 'qwen' ? $('qwenModel').value.trim() : $('grokModel').value.trim(), endpoint: $('qwenEndpoint').value.trim(), messages: state.messages.slice(-24) });
    state.messages.push({ role: 'assistant', content: result.text });
    state.lastAnswer = result.text; $('insertAnswer').disabled = false; $('lastAnswerState').textContent = result.model ? `via ${result.model}` : '';
    renderMessage('assistant', result.text);
  } catch (error) { renderMessage('assistant', `Error: ${error.message}`); toast(error.message); }
  finally { state.busy = false; $('send').disabled = false; $('send').textContent = '↑'; }
}

async function refreshSettings() {
  try {
    const data = await api('health');
    const provider = $('provider').value;
    $('keyStatus').textContent = data.configured[provider] ? 'Key active for this session' : '';
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
$('refreshSelection').addEventListener('click', readSelection);
$('addSelection').addEventListener('click', async () => {
  await readSelection();
  if (!state.currentSelection) return;
  state.sharedSelection = (state.currentSelection.rows || []).map(row => row.join('\t')).join('\n');
  $('contextHint').textContent = 'Selection will be shared with provider';
  toast('Selection attached to your next messages.');
});
$('saveKey').addEventListener('click', async () => {
  const apiKey = $('apiKey').value.trim(); if (!apiKey) return toast('Enter an API key first.');
  try { await api('configure', { provider: $('provider').value, apiKey }); $('apiKey').value = ''; await refreshSettings(); toast('Key loaded into local relay memory.'); }
  catch (error) { toast(error.message); }
});
$('clearKey').addEventListener('click', async () => {
  try { await api('clear-key', { provider: $('provider').value }); await refreshSettings(); toast('Key cleared.'); }
  catch (error) { toast(error.message); }
});
$('composer').addEventListener('submit', async event => { event.preventDefault(); const value = $('prompt').value; $('prompt').value = ''; await sendPrompt(value); });
$('prompt').addEventListener('keydown', event => { if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); $('composer').requestSubmit(); } });
$('insertAnswer').addEventListener('click', insertAnswer);
document.querySelectorAll('.action').forEach(button => button.addEventListener('click', () => { $('prompt').value = button.dataset.prompt; $('prompt').focus(); }));

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) { state.officeReady = true; readSelection(); }
  refreshSettings();
});

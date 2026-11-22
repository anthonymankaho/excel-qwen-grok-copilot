const QWEN_PATH = '/compatible-mode/v1/chat/completions';
const MODEL_RE = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,79}$/;
const SYSTEM_PROMPT = `You are an assistant inside Microsoft Office, revising the open Excel workbook or Word document the way Copilot and Claude for Excel do. Be accurate and concise. Distinguish values from formulas. The user message may include an [Excel selection ...], [Excel used range ...], [Word selection], or [Word document] block.

When the user asks to change the file, reply with a short explanation plus one fenced json block. Do not claim the change is already saved; the pane applies the actions. Prefer {"actions":[...]} with at most 40 actions. A legacy {"edits":[{"sheet":"Sheet1","cell":"D5","value":"=SUM(A1:A4)"}]} block is still accepted for single cells.

Excel actions:
- {"op":"set","sheet":"Sheet1","range":"A1:B2","values":[["Name","Total"],["Ada",100]]} — a formula is a string that starts with =. A single cell may use "value" instead of "values".
- {"op":"format","sheet":"Sheet1","range":"A1:B1","bold":true,"italic":false,"fill":"#1F4E79","fontColor":"#FFFFFF","numberFormat":"#,##0","align":"center","wrap":true}
- {"op":"clear","sheet":"Sheet1","range":"C2:C20"}
- {"op":"insertRows","sheet":"Sheet1","row":5,"count":1}
- {"op":"insertColumns","sheet":"Sheet1","column":"C","count":1}
- {"op":"deleteRows","sheet":"Sheet1","row":5,"count":1}
- {"op":"addSheet","name":"Summary"}
- {"op":"renameSheet","sheet":"Sheet1","name":"Data"}
- {"op":"autofit","sheet":"Sheet1","range":"A:D"}
- {"op":"sort","sheet":"Sheet1","range":"A1:B20","column":"B","ascending":false,"hasHeaders":true}
- {"op":"chart","sheet":"Sheet1","range":"A1:B20","type":"column","title":"Sales"} — type is column, bar, line, or pie
- {"op":"table","sheet":"Sheet1","range":"A1:B20","name":"SalesTable"}
- {"op":"pivot","sheet":"Data","range":"A1:Q753","name":"Aging Pivot","destinationSheet":"Pivot - Aging","destinationCell":"A3","rows":["Entity"],"columns":["Bucket"],"values":[{"field":"Total USD","summarizeBy":"Sum"}],"filters":["Currency"],"layout":"Tabular","grandTotals":true} — a real native PivotTable, not a formula imitation. The source range must include the header row and every field name must match a header exactly. The destination sheet is created when it does not exist. rows, columns and filters are arrays of header names; values is an array of {field, summarizeBy, name}. summarizeBy is Sum, Count, Average, Max, Min, CountNumbers, Product, StandardDeviation or Variance. layout is Tabular, Compact or Outline. At least one row or column field and at least one value field are required.

Word actions:
- {"op":"replaceSelection","text":"Revised paragraph."}
- {"op":"insertText","text":"Added at the end.","where":"end"} — where is start or end
- {"op":"formatSelection","bold":true,"italic":false}
- {"op":"styleSelection","style":"Heading 1"} — Heading 1, Heading 2, Heading 3, Title, or Normal

Use only actions for the open host. Do not propose WEBSERVICE, FILTERXML, or RTD unless the user explicitly asks. Do not invent a sheet or range you have not seen, except addSheet. For accounting work, journal entries are drafts: do not silently clear a reconciliation difference, do not change source figures to force a total to match, and do not claim that debits equal credits unless a pane validation says so. Do not assume the first columns are debit and credit.`;

export function qwenEndpointAllowed(endpoint) {
  if (typeof endpoint !== 'string' || !endpoint.trim()) return null;
  let url;
  try { url = new URL(endpoint.trim()); } catch { return null; }
  if (url.protocol !== 'https:' || url.username || url.password) return null;
  const hostname = url.hostname.toLowerCase().replace(/\.$/, '');
  const trusted = hostname === 'qwencloudapi.com' || hostname.endsWith('.qwencloudapi.com') || hostname === 'aliyuncs.com' || hostname.endsWith('.aliyuncs.com');
  if (!trusted) return null;
  const pathname = url.pathname.replace(/\/+$/, '') || '/';
  if (pathname !== QWEN_PATH) return null;
  return url;
}

export function requestHostAllowed(header, boundPort) {
  if (typeof header !== 'string' || !header.trim()) return false;
  const value = header.trim().toLowerCase();
  const suffix = `:${boundPort}`;
  return value === `127.0.0.1${suffix}` || value === `localhost${suffix}`;
}

export function originAllowed(origin, boundPort) {
  if (origin == null || origin === '') return true;
  return origin === `http://127.0.0.1:${boundPort}` || origin === `http://localhost:${boundPort}`;
}

export function takeSseDataLines(buffer) {
  const data = [];
  let rest = buffer;
  let idx;
  while ((idx = rest.indexOf('\n')) >= 0) {
    const line = rest.slice(0, idx).replace(/\r$/, '').trim();
    rest = rest.slice(idx + 1);
    if (!line.startsWith('data:')) continue;
    const raw = line.slice(5).trim();
    if (!raw || raw === '[DONE]') continue;
    data.push(raw);
  }
  return { data, rest };
}

export function providerDelta(provider, event) {
  if (!event || typeof event !== 'object') return { content: '', model: '', error: '' };
  if (provider === 'qwen') {
    if (event.error) {
      const message = event.error.message || event.error;
      return { content: '', model: '', error: typeof message === 'string' ? message : 'The provider failed while streaming.' };
    }
    const delta = event.choices?.[0]?.delta?.content;
    return { content: typeof delta === 'string' ? delta : '', model: typeof event.model === 'string' ? event.model : '', error: '' };
  }
  if (event.type === 'response.output_text.delta' && typeof event.delta === 'string') return { content: event.delta, model: '', error: '' };
  if (event.type === 'response.completed') return { content: '', model: typeof event.response?.model === 'string' ? event.response.model : '', error: '' };
  if (event.type === 'response.incomplete') {
    const reason = event.response?.incomplete_details?.reason || 'incomplete';
    return { content: '', model: typeof event.response?.model === 'string' ? event.response.model : '', error: `The provider stopped early (${reason}).` };
  }
  if (event.type === 'response.failed' || event.type === 'error') {
    const message = event.error?.message || event.response?.error?.message || event.message || 'The provider failed while streaming.';
    return { content: '', model: '', error: typeof message === 'string' ? message : 'The provider failed while streaming.' };
  }
  return { content: '', model: '', error: '' };
}

function hostNote(host) {
  return host === 'word'
    ? '\nThe open file is a Word document. Use only Word actions.'
    : '\nThe open file is an Excel workbook. Use only Excel actions.';
}

const PROFILE_NOTE = ' To understand a sheet with many rows, set profileRequest to {sheet, range} instead of contextRequest: you receive each column\'s header, filled and blank counts, distinct count, sample values with counts, and numeric sum, min and max. Prefer it before choosing pivot or SUMIFS fields. A read counts as successful only when cell values or a profile actually reach you; if the pane says nothing was read, request a smaller range rather than assuming success.';

function modeNote(mode) {
  if (mode === 'chat') return '\nMODE CHAT: Answer from the context. Do not emit an actions or edits block.';
  if (mode === 'plan') return '\nMODE PLAN: Return one json object with a plan and no claim that the file changed. Include goal, steps, assumptions, requiredInputs, risks, and finalValidation. Put actions in the same object only as a proposal. Do not invent validation pass or fail.' + PROFILE_NOTE;
  if (mode === 'agent') return '\nMODE AGENT: Return one json object with plan, optional actions for this step only, and decision set to COMPLETE, CONTINUE, REPLAN, NEEDS_USER, or FAILED. If you need a range, set contextRequest to {sheet, range} and do not guess the cells. Include the header row in a finance range. Set finalValidation columns only after those headers were read. Finance checks are calculated by the pane from the workbook. A failed check stops the loop. A check without an identified column is not a pass. Destructive actions wait for the user.' + PROFILE_NOTE;
  return '';
}

function resolveModel(model, fallback) {
  if (model == null || (typeof model === 'string' && model.trim() === '')) return { model: fallback };
  if (typeof model !== 'string' || !MODEL_RE.test(model.trim())) return { error: 'Invalid model name.' };
  return { model: model.trim() };
}

export function prepareChat(body, secrets) {
  const provider = body?.provider;
  const messages = body?.messages;
  if (!['qwen', 'grok'].includes(provider)) return { status: 400, error: 'Choose Qwen or Grok.' };
  if (!secrets[provider]) return { status: 401, error: `Add your ${provider === 'qwen' ? 'Qwen' : 'Grok'} API key in Settings first.` };
  if (!Array.isArray(messages) || messages.length < 1 || messages.length > 24) return { status: 400, error: 'Conversation must contain 1 to 24 messages.' };
  if (messages.some(m => !m || !['user', 'assistant'].includes(m.role) || typeof m.content !== 'string' || m.content.length > 50_000)) return { status: 400, error: 'Invalid conversation message.' };
  let url;
  let payload;
  if (provider === 'qwen') {
    url = qwenEndpointAllowed(body.endpoint);
    if (!url) return { status: 400, error: 'Qwen endpoint must be an HTTPS QwenCloud chat completions URL.' };
    const model = resolveModel(body.model, 'qwen4-max');
    if (model.error) return { status: 400, error: model.error };
    payload = { model: model.model, messages: [{ role: 'system', content: SYSTEM_PROMPT + hostNote(body.host) + modeNote(body.mode) }, ...messages], stream: false };
    if (body.fast) payload.enable_thinking = false;
  } else {
    url = new URL('https://api.x.ai/v1/responses');
    const model = resolveModel(body.model, 'grok-4.7');
    if (model.error) return { status: 400, error: model.error };
    payload = { model: model.model, instructions: SYSTEM_PROMPT + hostNote(body.host) + modeNote(body.mode), input: messages.map(({ role, content }) => ({ role, content })), store: false };
    if (body.fast) payload.reasoning = { effort: 'low' };
  }
  return { provider, model: payload.model, url, payload };
}

export function extractGrokText(data) {
  if (typeof data?.output_text === 'string' && data.output_text.trim()) return data.output_text;
  for (const item of data?.output || []) {
    if (item.type === 'message') {
      const text = (item.content || []).filter(part => part.type === 'output_text').map(part => part.text).join('');
      if (text) return text;
    }
  }
  return '';
}

export function providerErrorMessage(data, status) {
  return data?.error?.message || data?.message || `Provider returned HTTP ${status}.`;
}

export function relayTokenOk(presented, expected) {
  if (typeof expected !== 'string' || expected.length < 16) return false;
  if (typeof presented !== 'string' || presented.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= presented.charCodeAt(i) ^ expected.charCodeAt(i);
  return diff === 0;
}

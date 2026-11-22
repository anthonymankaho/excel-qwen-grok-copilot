(function (root) {
  const MAX_LOOPS = 4;
  const DECISIONS = ['COMPLETE', 'CONTINUE', 'REPLAN', 'NEEDS_USER', 'FAILED'];
  const CHECKS = ['debitEqualsCredit', 'formulaErrors', 'blankRequired', 'duplicateKeys', 'rowCount', 'totalDelta'];
  const DESTRUCTIVE = new Set(['clear', 'deleteRows', 'deleteColumns', 'renameSheet']);

  function cleanStrings(list, max) {
    if (!Array.isArray(list)) return [];
    return list.filter(item => typeof item === 'string' && item.trim()).map(item => item.trim().slice(0, 240)).slice(0, max);
  }

  function normalizeDecision(value) {
    const decision = String(value || '').trim().toUpperCase();
    return DECISIONS.includes(decision) ? decision : '';
  }

  function normalizeStep(raw, index) {
    const objective = typeof raw?.objective === 'string' ? raw.objective.trim() : '';
    if (!objective) return null;
    const actionType = ['inspect', 'edit', 'validate', 'report'].includes(raw.actionType) ? raw.actionType : 'inspect';
    return {
      id: Number.isInteger(raw.id) ? raw.id : index + 1,
      objective: objective.slice(0, 240),
      actionType,
      requiresApproval: raw.requiresApproval === true || actionType === 'edit',
      validation: typeof raw.validation === 'string' ? raw.validation.trim().slice(0, 240) : ''
    };
  }

  function normalizeCheck(raw, index) {
    const check = typeof raw?.check === 'string' ? raw.check : '';
    if (!CHECKS.includes(check)) return null;
    const columns = Array.isArray(raw.columns) ? raw.columns.filter(n => Number.isInteger(n) && n >= 1 && n <= 20).slice(0, 6) : [];
    return {
      id: index + 1,
      check,
      sheet: typeof raw.sheet === 'string' ? raw.sheet.trim().slice(0, 31) : '',
      range: typeof raw.range === 'string' ? raw.range.trim().slice(0, 32) : '',
      columns,
      tolerance: Number.isFinite(raw.tolerance) ? raw.tolerance : 0.005
    };
  }

  function normalizePlan(raw) {
    if (!raw || typeof raw !== 'object' || typeof raw.goal !== 'string' || !raw.goal.trim()) return null;
    const steps = Array.isArray(raw.steps) ? raw.steps.map(normalizeStep).filter(Boolean).slice(0, 12) : [];
    return {
      goal: raw.goal.trim().slice(0, 500),
      mode: ['chat', 'plan', 'agent'].includes(raw.mode) ? raw.mode : 'agent',
      steps,
      assumptions: cleanStrings(raw.assumptions, 8),
      requiredInputs: cleanStrings(raw.requiredInputs, 8),
      risks: cleanStrings(raw.risks, 8),
      finalValidation: Array.isArray(raw.finalValidation) ? raw.finalValidation.map(normalizeCheck).filter(Boolean).slice(0, 8) : []
    };
  }

  function normalizeContextRequest(raw) {
    if (!raw || typeof raw.sheet !== 'string' || typeof raw.range !== 'string') return null;
    const sheet = raw.sheet.trim();
    const range = raw.range.trim();
    if (!sheet || !range || sheet.length > 31 || range.length > 32) return null;
    return { sheet, range };
  }

  function parseEnvelope(text) {
    const source = String(text || '');
    const blocks = [...source.matchAll(/```(?:json)?\s*([\s\S]*?)```/gi)];
    for (let i = blocks.length - 1; i >= 0; i--) {
      try {
        const parsed = JSON.parse(blocks[i][1]);
        if (!parsed || typeof parsed !== 'object') continue;
        if (!parsed.plan && !parsed.decision && !parsed.contextRequest && !parsed.actions) continue;
        const prose = (source.slice(0, blocks[i].index) + source.slice(blocks[i].index + blocks[i][0].length)).trim();
        return {
          prose,
          plan: normalizePlan(parsed.plan),
          actions: Array.isArray(parsed.actions) ? parsed.actions : null,
          decision: normalizeDecision(parsed.decision),
          contextRequest: normalizeContextRequest(parsed.contextRequest)
        };
      } catch { /* not this block */ }
    }
    return { prose: source, plan: null, actions: null, decision: '', contextRequest: null };
  }

  function routeProvider({ selected, mode, fast, text, ready }) {
    const available = name => Boolean(ready?.[name]);
    if (selected === 'qwen' || selected === 'grok') {
      if (!available(selected)) return { error: `Add your ${selected === 'qwen' ? 'Qwen' : 'Grok'} API key before sending.`, provider: selected };
      return { provider: selected, reason: 'manual', fallback: '' };
    }
    if (selected !== 'auto') return { error: 'Choose Qwen, Grok, or Auto.', provider: '' };
    const finance = /reconcil|journal|debit|credit|month-end|month end|trial balance|ledger|subledger|accrual/i.test(String(text || ''));
    const complex = mode === 'agent' || mode === 'plan' || finance || String(text || '').length > 500;
    const prefer = complex && !fast ? ['grok', 'qwen'] : ['qwen', 'grok'];
    const provider = prefer.find(available);
    if (!provider) return { error: 'Add a Qwen or Grok API key before Auto can choose a model.', provider: '' };
    const fallback = prefer.find(name => name !== provider && available(name)) || '';
    return { provider, reason: `auto:${complex && !fast ? 'complex' : 'quick'}`, fallback };
  }

  function isDestructive(action) {
    return DESTRUCTIVE.has(action?.op);
  }

  function isExcelError(value) {
    if (typeof value === 'string') return /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NULL!|NUM!|GETTING_DATA)$/i.test(value.trim());
    return Boolean(value && typeof value === 'object' && (value.type === 'Error' || value.errorType));
  }

  function asNumber(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (typeof value !== 'string') return null;
    const text = value.trim();
    if (!text || text.startsWith('=')) return null;
    const n = Number(text.replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  function sumNumbers(values) {
    let total = 0;
    for (const row of values || []) {
      for (const value of (Array.isArray(row) ? row : [row])) {
        const n = asNumber(value);
        if (n != null) total += n;
      }
    }
    return Math.round(total * 100) / 100;
  }

  function evaluateGridCheck(check, grid, prior) {
    const values = grid?.values || [];
    const formulas = grid?.formulas || [];
    const base = { check: check.check, before: prior?.before ?? '', after: '', difference: '', details: [] };
    if (check.check === 'formulaErrors') {
      const details = [];
      values.forEach((row, r) => (row || []).forEach((value, c) => {
        if (isExcelError(value) || isExcelError(formulas?.[r]?.[c])) details.push(`R${r + 1}C${c + 1}`);
      }));
      return { ...base, status: details.length ? 'FAIL' : 'PASS', after: String(details.length), details };
    }
    if (check.check === 'debitEqualsCredit') {
      let debit = 0;
      let credit = 0;
      let seen = 0;
      for (const row of values) {
        const left = asNumber(row?.[0]);
        const right = asNumber(row?.[1]);
        if (left == null && right == null) continue;
        seen += 1;
        debit += left || 0;
        credit += right || 0;
      }
      const difference = Math.round((debit - credit) * 100) / 100;
      if (!seen) return { ...base, status: 'WARNING', before: '0', after: '0', difference: '0', details: ['No numeric debit or credit rows.'] };
      return { ...base, status: Math.abs(difference) <= 0.005 ? 'PASS' : 'FAIL', before: String(debit), after: String(credit), difference: String(difference), details: [`Debit ${debit}`, `Credit ${credit}`] };
    }
    if (check.check === 'blankRequired') {
      const cols = check.columns?.length ? check.columns : [1];
      const details = [];
      values.forEach((row, r) => {
        if (!(row || []).some(value => String(value ?? '').trim())) return;
        for (const col of cols) {
          if (String(row?.[col - 1] ?? '').trim() === '') details.push(`R${r + 1}C${col}`);
        }
      });
      return { ...base, status: details.length ? 'FAIL' : 'PASS', after: String(details.length), details };
    }
    if (check.check === 'duplicateKeys') {
      const cols = check.columns?.length ? check.columns : [1];
      const seen = new Set();
      const details = [];
      values.forEach(row => {
        const key = cols.map(col => String(row?.[col - 1] ?? '').trim()).join('|');
        if (!key.replace(/\|/g, '')) return;
        if (seen.has(key)) details.push(key);
        else seen.add(key);
      });
      return { ...base, status: details.length ? 'FAIL' : 'PASS', after: String(details.length), details };
    }
    if (check.check === 'rowCount') {
      const after = values.length;
      const before = Number(prior?.before ?? after);
      const difference = after - before;
      return { ...base, status: difference === 0 ? 'PASS' : 'WARNING', before: String(before), after: String(after), difference: String(difference), details: [] };
    }
    if (check.check === 'totalDelta') {
      const after = sumNumbers(values);
      const before = Number(prior?.before ?? after);
      const difference = Math.round((after - before) * 100) / 100;
      const tolerance = Number.isFinite(check.tolerance) ? check.tolerance : 0.005;
      return { ...base, status: Math.abs(difference) <= tolerance ? 'PASS' : 'FAIL', before: String(before), after: String(after), difference: String(difference), details: [] };
    }
    return { ...base, status: 'WARNING', details: ['Unsupported check.'] };
  }

  function explainFailure(result) {
    const guidance = {
      debitEqualsCredit: ['A journal line may be missing, duplicated, or signed the wrong way.', 'Treat the entry as a draft. Do not post it until debit equals credit.'],
      formulaErrors: ['A formula references a missing cell, a bad value, or a broken range.', 'Open the listed cells and correct the formula before relying on the total.'],
      blankRequired: ['A required field is empty on a row that otherwise has data.', 'Fill the blank or exclude the row. Do not guess the missing value.'],
      duplicateKeys: ['The same key appears on more than one row.', 'Confirm whether this is a duplicate posting before keeping both rows.'],
      totalDelta: ['The total changed by more than the allowed tolerance.', 'Trace the source rows to the output. Do not force the totals to match.'],
      rowCount: ['The number of rows changed.', 'Confirm the inserted or deleted rows are the ones you intended.']
    };
    const pair = guidance[result.check] || ['The check did not pass.', 'Review the evidence before changing the workbook again.'];
    return { issue: result.check, evidence: result.details || [], possibleCause: pair[0], requiredCheck: result.check, recommendedAction: pair[1] };
  }

  function decide({ decision, validations, iteration, maxLoops }) {
    if ((validations || []).some(item => item.status === 'FAIL')) return 'FAILED';
    const next = normalizeDecision(decision) || 'COMPLETE';
    if ((next === 'CONTINUE' || next === 'REPLAN') && iteration + 1 >= maxLoops) return 'FAILED';
    return next;
  }

  function audit(entry) {
    const copy = {};
    for (const [key, value] of Object.entries(entry || {})) {
      if (/key|token|secret|authorization/i.test(key)) continue;
      copy[key] = value;
    }
    return copy;
  }

  function summarize(entries) {
    return (entries || []).map(entry => {
      if (entry.kind === 'request') return `Request (${entry.mode}) via ${entry.provider} · ${entry.reason}`;
      if (entry.kind === 'context') return `Context ${entry.address || 'none'} · ${entry.chars || 0} chars`;
      if (entry.kind === 'model') return `Model ${entry.model || entry.provider || ''} · ${entry.decision || 'no decision'}${entry.plan ? ` · ${entry.plan}` : ''}`;
      if (entry.kind === 'execute') return `Execute ${entry.status} · ${entry.count || 0} actions · blocked ${entry.blocked || 0}`;
      if (entry.kind === 'validation') return `${entry.status} ${entry.check} · before ${entry.before} after ${entry.after} diff ${entry.difference}`;
      if (entry.kind === 'blocked') return `Blocked · ${entry.reason}`;
      if (entry.kind === 'stopped') return `Stopped · ${entry.reason}`;
      if (entry.kind === 'fallback') return `Fallback to ${entry.provider} · ${entry.reason}`;
      return entry.kind || 'event';
    }).join('\n');
  }

  root.AgentCore = {
    MAX_LOOPS, parseEnvelope, normalizePlan, routeProvider, isDestructive, isExcelError,
    evaluateGridCheck, explainFailure, decide, audit, summarize, sumNumbers
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

(function (root) {
  const MAX_LOOPS = 10;
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
      tolerance: Number.isFinite(raw.tolerance) && raw.tolerance >= 0 ? raw.tolerance : null
    };
  }

  function normalizePlan(raw) {
    // Models often send plan as prose rather than an object; keep it instead of dropping it.
    if (typeof raw === 'string' && raw.trim()) raw = { goal: raw };
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

  // A profile asks for column statistics instead of raw cells, so a 753-row sheet
  // can be understood without shipping every cell into the context window.
  function normalizeProfileRequest(raw) {
    if (!raw || typeof raw.sheet !== 'string') return null;
    const sheet = raw.sheet.trim();
    if (!sheet || sheet.length > 31) return null;
    const range = typeof raw.range === 'string' ? raw.range.trim() : '';
    if (range.length > 32) return null;
    return { sheet, range };
  }

  function parseEnvelope(text) {
    const source = String(text || '');
    const blocks = [...source.matchAll(/```(?:json)?\s*([\s\S]*?)```/gi)];
    for (let i = blocks.length - 1; i >= 0; i--) {
      try {
        const parsed = JSON.parse(blocks[i][1]);
        if (!parsed || typeof parsed !== 'object') continue;
        if (!parsed.plan && !parsed.decision && !parsed.contextRequest && !parsed.profileRequest && !parsed.actions) continue;
        const prose = (source.slice(0, blocks[i].index) + source.slice(blocks[i].index + blocks[i][0].length)).trim();
        return {
          prose,
          plan: normalizePlan(parsed.plan),
          actions: Array.isArray(parsed.actions) ? parsed.actions : null,
          decision: normalizeDecision(parsed.decision),
          contextRequest: normalizeContextRequest(parsed.contextRequest),
          profileRequest: normalizeProfileRequest(parsed.profileRequest)
        };
      } catch { /* not this block */ }
    }
    return { prose: source, plan: null, actions: null, decision: '', contextRequest: null, profileRequest: null };
  }

  function isFinancePrompt(text) {
    return /\b(?:finance|financial|accounting|reconcil(?:e|iation|ing)|journal|ledger|subledger|trial balance|debit|credit|month[- ]end|accrual|cash flow|profit and loss|income statement|balance sheet|working capital|accounts receivable|accounts payable|inventory|depreciation|tax|audit|hkfrs|hkas|ifrs|ias)\b|\b(?:p&l|ar|ap|a\/r|a\/p)\b|財務|會計|對賬|對帳|月結|月尾|分錄|借方|貸方|應收|應付|試算表|總賬|總帳|現金流|損益|資產負債|利潤|收入|成本|費用|稅務|稅項|審計|存貨|固定資產|折舊|租賃|預算|預測/i.test(String(text || ''));
  }

  function routeProvider({ selected, mode, fast, text, ready }) {
    const available = name => Boolean(ready?.[name]);
    if (selected === 'qwen' || selected === 'grok') {
      if (!available(selected)) return { error: `Add your ${selected === 'qwen' ? 'Qwen' : 'Grok'} API key before sending.`, provider: selected };
      return { provider: selected, reason: 'manual', fallback: '' };
    }
    if (selected !== 'auto') return { error: 'Choose Qwen, Grok, or Auto.', provider: '' };
    const finance = isFinancePrompt(text);
    const complex = mode === 'agent' || mode === 'plan' || finance || String(text || '').length > 500;
    const prefer = complex && !fast ? ['grok', 'qwen'] : ['qwen', 'grok'];
    const provider = prefer.find(available);
    if (!provider) return { error: 'Add a Qwen or Grok API key before Auto can choose a model.', provider: '' };
    const fallback = prefer.find(name => name !== provider && available(name)) || '';
    return { provider, reason: `auto:${complex && !fast ? 'complex' : 'quick'}`, fallback };
  }

  function isDestructive(action) {
    // bulk marks a change Undo cannot reverse, so it waits for a click like a delete.
    return DESTRUCTIVE.has(action?.op) || action?.bulk === true;
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

  const HEADER_ALIASES = {
    debit: ['debit', 'debits', 'dr', 'debit amount'],
    credit: ['credit', 'credits', 'cr', 'credit amount'],
    amount: ['amount', 'amt', 'net amount'],
    key: ['invoice', 'invoice no', 'invoice number', 'voucher', 'voucher no', 'document no', 'document number', 'reference']
  };

  function columnSpec(value) {
    if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 50) return value;
    if (typeof value !== 'string') return null;
    const text = value.trim();
    if (!text) return null;
    if (/^[1-9][0-9]?$/.test(text)) return Number(text);
    return text.slice(0, 40);
  }

  function columnList(value) {
    const source = Array.isArray(value) ? value : typeof value === 'string' ? value.split(',') : [];
    return source.map(columnSpec).filter(item => item != null).slice(0, 6);
  }

  function normalizeFinanceProfile(raw) {
    const source = raw && typeof raw === 'object' ? raw : {};
    const tolerance = Number(source.tolerance);
    return {
      debitColumn: columnSpec(source.debitColumn),
      creditColumn: columnSpec(source.creditColumn),
      amountColumn: columnSpec(source.amountColumn),
      keyColumns: columnList(source.keyColumns),
      requiredColumns: columnList(source.requiredColumns),
      entityColumn: columnSpec(source.entityColumn),
      currencyColumn: columnSpec(source.currencyColumn),
      dateColumn: columnSpec(source.dateColumn),
      tolerance: Number.isFinite(tolerance) && tolerance >= 0 && tolerance <= 1000 ? tolerance : 0.01
    };
  }

  function headerText(value) {
    return String(value ?? '').trim().toLowerCase().replace(/[_./]+/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function findAlias(header, aliases) {
    const found = [];
    (header || []).forEach((cell, index) => {
      if (aliases.includes(headerText(cell))) found.push(index + 1);
    });
    return found;
  }

  function resolveSpec(spec, header) {
    if (typeof spec === 'number') return spec >= 1 && spec <= 50 ? spec : null;
    if (typeof spec !== 'string') return null;
    const want = headerText(spec);
    const index = (header || []).findIndex(cell => headerText(cell) === want);
    return index >= 0 ? index + 1 : null;
  }

  function resolveMany(specs, header) {
    if (!specs?.length) return null;
    const columns = [];
    for (const spec of specs) {
      const index = resolveSpec(spec, header);
      if (!index) return null;
      columns.push(index);
    }
    return columns;
  }

  function headerWasUsed(header) {
    const aliases = Object.values(HEADER_ALIASES).flat();
    return (header || []).some(cell => aliases.includes(headerText(cell)));
  }

  function bodyRows(values, header) {
    return headerWasUsed(header) ? values.slice(1) : values;
  }

  function needsUser(base, detail) {
    return { ...base, status: 'NEEDS_USER', before: '', after: '', difference: '', details: [detail] };
  }

  function toleranceOf(check, profile) {
    if (Number.isFinite(check.tolerance) && check.tolerance >= 0) return check.tolerance;
    return profile.tolerance;
  }

  function columnTotal(rows, column) {
    let total = 0;
    let seen = 0;
    for (const row of rows) {
      const amount = asNumber(row?.[column - 1]);
      if (amount == null) continue;
      seen += 1;
      total += amount;
    }
    return { total: Math.round(total * 100) / 100, seen };
  }

  function debitColumns(check, header, profile) {
    if (profile.debitColumn != null || profile.creditColumn != null) {
      const debit = resolveSpec(profile.debitColumn, header);
      const credit = resolveSpec(profile.creditColumn, header);
      return debit && credit && debit !== credit ? [debit, credit] : null;
    }
    if (check.columns?.length >= 2 && check.columns[0] !== check.columns[1]) return [check.columns[0], check.columns[1]];
    const debit = findAlias(header, HEADER_ALIASES.debit);
    const credit = findAlias(header, HEADER_ALIASES.credit);
    if (debit.length === 1 && credit.length === 1 && debit[0] !== credit[0]) return [debit[0], credit[0]];
    return null;
  }

  function evaluateGridCheck(check, grid, prior) {
    const values = grid?.values || [];
    const formulas = grid?.formulas || [];
    const profile = normalizeFinanceProfile(check?.profile);
    const header = values[0] || [];
    const rows = bodyRows(values, header);
    const base = { check: check.check, before: prior?.before ?? '', after: '', difference: '', details: [] };
    if (check.check === 'formulaErrors') {
      const details = [];
      values.forEach((row, r) => (row || []).forEach((value, c) => {
        if (isExcelError(value) || isExcelError(formulas?.[r]?.[c])) details.push(`R${r + 1}C${c + 1}`);
      }));
      return { ...base, status: details.length ? 'FAIL' : 'PASS', after: String(details.length), details };
    }
    if (check.check === 'debitEqualsCredit') {
      const columns = debitColumns(check, header, profile);
      if (!columns) return needsUser(base, 'Debit and credit columns were not identified.');
      const debit = columnTotal(rows, columns[0]);
      const credit = columnTotal(rows, columns[1]);
      const difference = Math.round((debit.total - credit.total) * 100) / 100;
      if (!debit.seen && !credit.seen) return { ...base, status: 'WARNING', before: '0', after: '0', difference: '0', details: ['No numeric debit or credit rows.'] };
      const tolerance = toleranceOf(check, profile);
      return { ...base, status: Math.abs(difference) <= tolerance ? 'PASS' : 'FAIL', before: String(debit.total), after: String(credit.total), difference: String(difference), details: [`Debit ${debit.total}`, `Credit ${credit.total}`] };
    }
    if (check.check === 'blankRequired') {
      const columns = profile.requiredColumns.length ? resolveMany(profile.requiredColumns, header) : (check.columns?.length ? check.columns : null);
      if (!columns) return needsUser(base, 'Required columns were not identified.');
      const details = [];
      rows.forEach((row, r) => {
        if (!(row || []).some(value => String(value ?? '').trim())) return;
        for (const col of columns) {
          if (String(row?.[col - 1] ?? '').trim() === '') details.push(`R${r + 1}C${col}`);
        }
      });
      return { ...base, status: details.length ? 'FAIL' : 'PASS', after: String(details.length), details };
    }
    if (check.check === 'duplicateKeys') {
      let columns = profile.keyColumns.length ? resolveMany(profile.keyColumns, header) : (check.columns?.length ? check.columns : null);
      if (!columns) {
        const found = findAlias(header, HEADER_ALIASES.key);
        columns = found.length === 1 ? found : null;
      }
      if (!columns) return needsUser(base, 'The business-key column was not identified.');
      const seen = new Set();
      const details = [];
      rows.forEach(row => {
        const key = columns.map(col => String(row?.[col - 1] ?? '').trim()).join('|');
        if (!key.replace(/\|/g, '')) return;
        if (seen.has(key)) details.push(key);
        else seen.add(key);
      });
      return { ...base, status: details.length ? 'FAIL' : 'PASS', after: String(details.length), details };
    }
    if (check.check === 'rowCount') {
      const after = rows.length;
      const before = Number(prior?.before ?? after);
      const difference = after - before;
      return { ...base, status: difference === 0 ? 'PASS' : 'WARNING', before: String(before), after: String(after), difference: String(difference), details: [] };
    }
    if (check.check === 'totalDelta') {
      let column = null;
      if (profile.amountColumn != null) column = resolveSpec(profile.amountColumn, header);
      else if (check.columns?.length === 1) column = check.columns[0];
      else {
        const amounts = findAlias(header, HEADER_ALIASES.amount);
        if (amounts.length === 1) column = amounts[0];
        else {
          const numeric = new Set();
          rows.forEach(row => (row || []).forEach((value, index) => { if (asNumber(value) != null) numeric.add(index + 1); }));
          if (numeric.size === 1) column = [...numeric][0];
        }
      }
      if (!column) return needsUser(base, 'The amount column was not identified.');
      const after = columnTotal(rows, column).total;
      const before = Number(prior?.before ?? after);
      const difference = Math.round((after - before) * 100) / 100;
      const tolerance = toleranceOf(check, profile);
      return { ...base, status: Math.abs(difference) <= tolerance ? 'PASS' : 'FAIL', before: String(before), after: String(after), difference: String(difference), details: [] };
    }
    return { ...base, status: 'WARNING', details: ['Unsupported check.'] };
  }

  function explainFailure(result) {
    if (result.status === 'NEEDS_USER') {
      return {
        issue: result.check,
        evidence: result.details || [],
        possibleCause: 'The columns for this check were not identified from headers or Finance settings.',
        requiredCheck: result.check,
        recommendedAction: 'Name the columns in Finance checks, or include a header row the pane recognizes. This result is not a pass or a fail.'
      };
    }
    const guidance = {
      debitEqualsCredit: ['A journal line may be missing, duplicated, or signed the wrong way.', 'Treat the entry as a draft. Do not post it until debit equals credit. Do not change source figures to force the totals together.'],
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
    const items = validations || [];
    if (items.some(item => item.status === 'FAIL')) return 'FAILED';
    if (items.some(item => item.status === 'NEEDS_USER')) return 'NEEDS_USER';
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
      if (entry.kind === 'execute') {
        const skipped = entry.skipped ? ` · ${entry.skipped} rejected as unusable` : '';
        const reason = entry.reason ? ` · ${entry.reason}` : '';
        return `Execute ${entry.status} · ${entry.count || 0} actions · blocked ${entry.blocked || 0}${skipped}${reason}`;
      }
      if (entry.kind === 'validation') return `${entry.status} ${entry.check}${entry.address ? ` @ ${entry.address}` : ''} · before ${entry.before} after ${entry.after} diff ${entry.difference}`;
      if (entry.kind === 'blocked') return `Blocked · ${entry.reason}`;
      if (entry.kind === 'stopped') return `Stopped · ${entry.reason}`;
      if (entry.kind === 'fallback') return `Fallback to ${entry.provider} · ${entry.reason}`;
      return entry.kind || 'event';
    }).join('\n');
  }

  root.AgentCore = {
    MAX_LOOPS, parseEnvelope, normalizePlan, normalizeFinanceProfile, isFinancePrompt, routeProvider, isDestructive, isExcelError,
    evaluateGridCheck, explainFailure, decide, audit, summarize, sumNumbers
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

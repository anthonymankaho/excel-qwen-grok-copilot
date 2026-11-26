import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(root, 'public', 'agent-core.js'), 'utf8'), sandbox);
const AgentCore = sandbox.AgentCore;

test('a plan keeps steps, assumptions, and finance checks, and drops an unnamed goal', () => {
  const envelope = AgentCore.parseEnvelope('Draft only.\n```json\n{"plan":{"goal":"Draft the accrual","mode":"agent","steps":[{"id":1,"objective":"Inspect the source","actionType":"inspect","requiresApproval":false,"validation":"headers exist"},{"objective":""}],"assumptions":["Amount is the source of truth"],"requiredInputs":["Entity"],"risks":["Do not post"],"finalValidation":[{"check":"debitEqualsCredit","sheet":"Journal","range":"C2:D20"},{"check":"nope"}]},"decision":"CONTINUE"}\n```');
  assert.equal(envelope.prose, 'Draft only.');
  assert.equal(envelope.plan.goal, 'Draft the accrual');
  assert.equal(envelope.plan.steps.length, 1);
  assert.equal(envelope.plan.steps[0].actionType, 'inspect');
  assert.equal(envelope.plan.steps[0].requiresApproval, false);
  assert.equal(envelope.plan.finalValidation.length, 1);
  assert.equal(envelope.decision, 'CONTINUE');
  assert.equal(AgentCore.normalizePlan({ steps: [] }), null);
});

test('routing is Qwen-only and never sends without a key', () => {
  assert.equal(AgentCore.MAX_LOOPS, 25);
  const manual = AgentCore.routeProvider({ selected: 'qwen', mode: 'agent', fast: false, text: 'Reconcile debits', ready: { qwen: true } });
  assert.equal(manual.provider, 'qwen');
  assert.equal(manual.reason, 'manual');
  const missing = AgentCore.routeProvider({ selected: 'qwen', mode: 'chat', fast: false, text: 'Hi', ready: { qwen: false } });
  assert.match(missing.error, /Qwen/);
  const auto = AgentCore.routeProvider({ selected: 'auto', mode: 'agent', fast: false, text: 'Reconcile the ledger', ready: { qwen: true } });
  assert.equal(auto.provider, 'qwen');
  assert.equal(auto.fallback, '');
  const other = AgentCore.routeProvider({ selected: 'grok', mode: 'chat', fast: false, text: 'Hi', ready: { qwen: true } });
  assert.match(other.error, /Qwen only/);
});

test('a context request keeps a paging offset and normalizes invalid ones to zero', () => {
  const envelope = AgentCore.parseEnvelope('```json\n{"contextRequest":{"sheet":"G. GL 2026","range":"A1:AD41773","offset":2000},"decision":"CONTINUE"}\n```');
  assert.equal(envelope.contextRequest.offset, 2000);
  const bad = AgentCore.parseEnvelope('```json\n{"contextRequest":{"sheet":"S","range":"A1:B2","offset":-5},"decision":"CONTINUE"}\n```');
  assert.equal(bad.contextRequest.offset, 0);
  const huge = AgentCore.parseEnvelope('```json\n{"contextRequest":{"sheet":"S","range":"A1:B2","offset":99999999},"decision":"CONTINUE"}\n```');
  assert.equal(huge.contextRequest.offset, 1048575);
});

test('debit, blanks, duplicates, and formula errors are calculated from identified columns', () => {
  const headers = ['Debit', 'Credit'];
  const balanced = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit' }, { values: [headers, [100, 40], [20, 80]] });
  assert.equal(balanced.status, 'PASS');
  assert.equal(balanced.difference, '0');
  const open = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit' }, { values: [headers, [100, 40]] });
  assert.equal(open.status, 'FAIL');
  assert.equal(open.difference, '60');
  const unnamed = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit' }, { values: [[100, 40], [20, 80]] });
  assert.equal(unnamed.status, 'NEEDS_USER');
  assert.equal(unnamed.difference, '');
  const wrongHeaders = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit' }, { values: [['Amount', 'Tax'], [100, 40]] });
  assert.equal(wrongHeaders.status, 'NEEDS_USER');
  const named = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit', columns: [1, 2] }, { values: [[100, 40], [20, 80]] });
  assert.equal(named.status, 'PASS');
  const profile = AgentCore.normalizeFinanceProfile({ debitColumn: '1', creditColumn: '2', tolerance: 0.02 });
  const within = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit', profile }, { values: [[100, 100.01]] });
  assert.equal(within.status, 'PASS');
  const beyond = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit', profile: { ...profile, tolerance: 0.01 } }, { values: [[100, 99.98]] });
  assert.equal(beyond.status, 'FAIL');
  assert.equal(beyond.difference, '0.02');
  const blanks = AgentCore.evaluateGridCheck({ check: 'blankRequired', columns: [2] }, { values: [['Acme', ''], ['', '']] });
  assert.equal(blanks.status, 'FAIL');
  assert.equal(blanks.details.length, 1);
  assert.equal(blanks.details[0], 'R1C2');
  const blankUnknown = AgentCore.evaluateGridCheck({ check: 'blankRequired' }, { values: [['Acme', '']] });
  assert.equal(blankUnknown.status, 'NEEDS_USER');
  const dupes = AgentCore.evaluateGridCheck({ check: 'duplicateKeys', columns: [1] }, { values: [['A-1', 5], ['A-1', 9], ['B-2', 1]] });
  assert.equal(dupes.status, 'FAIL');
  const dupeHeader = AgentCore.evaluateGridCheck({ check: 'duplicateKeys' }, { values: [['Invoice', 'Amount'], ['A-1', 5], ['A-1', 9]] });
  assert.equal(dupeHeader.status, 'FAIL');
  const dupeUnknown = AgentCore.evaluateGridCheck({ check: 'duplicateKeys' }, { values: [['A-1', 5], ['A-1', 9]] });
  assert.equal(dupeUnknown.status, 'NEEDS_USER');
  const errors = AgentCore.evaluateGridCheck({ check: 'formulaErrors' }, { values: [[10, '#REF!']], formulas: [['=A1', '#REF!']] });
  assert.equal(errors.status, 'FAIL');
  const unread = AgentCore.evaluateGridCheck({ check: 'formulaErrors' }, { values: [], formulas: [], readError: 'The requested resource does not exist.' });
  assert.equal(unread.status, 'NEEDS_USER');
  assert.match(unread.details[0], /could not be fully read/);
  const mixedTotal = AgentCore.evaluateGridCheck({ check: 'totalDelta' }, { values: [[10, 20]] });
  assert.equal(mixedTotal.status, 'NEEDS_USER');
  const explained = AgentCore.explainFailure(open);
  assert.equal(explained.issue, 'debitEqualsCredit');
  assert.match(explained.recommendedAction, /draft/i);
  const asked = AgentCore.explainFailure(unnamed);
  assert.match(asked.recommendedAction, /not a pass or a fail/i);
});

test('a failed finance check stops the loop even if the model says continue', () => {
  const stopped = AgentCore.decide({ decision: 'CONTINUE', validations: [{ status: 'FAIL' }], iteration: 0, maxLoops: 4 });
  assert.equal(stopped, 'FAILED');
  const limited = AgentCore.decide({ decision: 'REPLAN', validations: [], iteration: 3, maxLoops: 4 });
  assert.equal(limited, 'FAILED');
  const done = AgentCore.decide({ decision: '', validations: [{ status: 'PASS' }], iteration: 0, maxLoops: 4 });
  assert.equal(done, 'COMPLETE');
  const ask = AgentCore.decide({ decision: 'COMPLETE', validations: [{ status: 'NEEDS_USER' }], iteration: 0, maxLoops: 4 });
  assert.equal(ask, 'NEEDS_USER');
});

test('the audit summary keeps the trail and drops secrets', () => {
  const entry = AgentCore.audit({ kind: 'request', mode: 'agent', provider: 'qwen', reason: 'manual', apiKey: 'sk-secret', text: 'Reconcile' });
  assert.equal(entry.apiKey, undefined);
  const summary = AgentCore.summarize([entry, AgentCore.audit({ kind: 'validation', status: 'FAIL', check: 'debitEqualsCredit', before: '100', after: '40', difference: '60' })]);
  assert.match(summary, /Request \(agent\) via qwen/);
  assert.match(summary, /FAIL debitEqualsCredit/);
  assert.equal(summary.includes('sk-secret'), false);
});

test('groupBalance balances each voucher and crossCheck compares two ranges', () => {
  const headers = ['Voucher', 'Debit', 'Credit'];
  const balanced = AgentCore.evaluateGridCheck({ check: 'groupBalance' }, { values: [headers, ['GJ001', 100, 100], ['GJ002', 50, 50]] });
  assert.equal(balanced.status, 'PASS');
  assert.match(balanced.details[0], /2 groups balanced/);
  const open = AgentCore.evaluateGridCheck({ check: 'groupBalance' }, { values: [headers, ['GJ001', 100, 99.99], ['GJ002', 50, 50]] });
  assert.equal(open.status, 'FAIL');
  assert.match(open.details[0], /GJ001: DR 100 CR 99.99 diff 0.01/);
  const unnamed = AgentCore.evaluateGridCheck({ check: 'groupBalance' }, { values: [['A', 'B', 'C'], ['x', 1, 1]] });
  assert.equal(unnamed.status, 'NEEDS_USER');
  const cross = AgentCore.evaluateGridCheck({ check: 'crossCheck', columns: [2] }, { values: [headers, ['GJ001', 100, 100], ['GJ002', 50, 50]] }, { values: [headers, ['GJ001', 100, 100]] });
  assert.equal(cross.status, 'FAIL');
  assert.match(cross.details[0], /rows 2 vs 1/);
  const crossOk = AgentCore.evaluateGridCheck({ check: 'crossCheck', columns: [2] }, { values: [headers, ['GJ001', 100, 100]] }, { values: [headers, ['GJ001', 100, 100]] });
  assert.equal(crossOk.status, 'PASS');
  const crossUnread = AgentCore.evaluateGridCheck({ check: 'crossCheck' }, { values: [headers] }, { values: [], readError: 'gone' });
  assert.equal(crossUnread.status, 'NEEDS_USER');
  const scan = AgentCore.parseEnvelope('```json\n{"scanRequest":true,"decision":"CONTINUE"}\n```');
  assert.equal(JSON.stringify(scan.scanRequest), '{}');
});

test('a OneDrive source request keeps search and one-file read, and drops a blank query', () => {
  const search = AgentCore.parseEnvelope('```json\n{"sourceRequest":{"action":"search","query":"October aging"},"decision":"CONTINUE"}\n```');
  assert.equal(search.sourceRequest.action, 'search');
  assert.equal(search.sourceRequest.query, 'October aging');
  const read = AgentCore.parseEnvelope('```json\n{"sourceRequest":{"action":"read","id":"item-1","name":"aging.xlsx"},"decision":"CONTINUE"}\n```');
  assert.equal(read.sourceRequest.action, 'read');
  assert.equal(read.sourceRequest.id, 'item-1');
  assert.equal(read.sourceRequest.name, 'aging.xlsx');
  const blank = AgentCore.parseEnvelope('```json\n{"sourceRequest":{"action":"search","query":"  "},"decision":"CONTINUE"}\n```');
  assert.equal(blank.sourceRequest, null);
  const github = AgentCore.parseEnvelope('```json\n{"githubRequest":{"action":"read","repo":"owner/repo","path":"README.md"},"decision":"CONTINUE"}\n```');
  assert.equal(github.githubRequest.path, 'README.md');
  const write = AgentCore.parseEnvelope('```json\n{"githubWrite":{"repo":"owner/repo","title":"Update notes","files":[{"path":"README.md","content":"Hello"}]}}\n```');
  assert.equal(write.githubWrite.files.length, 1);
  assert.equal(write.githubWrite.files[0].content, 'Hello');
  assert.equal(AgentCore.parseGitCommand('open src/app.js').action, 'open');
  assert.equal(AgentCore.parseGitCommand('git pull').action, 'pull');
  assert.equal(AgentCore.parseGitCommand('git commit -m "Update the note"').message, 'Update the note');
  assert.equal(AgentCore.parseGitCommand('git commit and push Update the note').action, 'commit-push');
  assert.equal(AgentCore.parseGitCommand('git commit and push Update the note').message, 'Update the note');
  assert.equal(AgentCore.parseGitCommand('git push').action, 'push');
  assert.equal(AgentCore.parseGitCommand('git push').rejected, false);
  assert.equal(AgentCore.parseGitCommand('git push --force').rejected, true);
  assert.equal(AgentCore.parseGitCommand('please git commit'), null);
});

test('a prose plan string is kept and a wide context request survives parsing', () => {
  const envelope = AgentCore.parseEnvelope('Reading first.\n```json\n{"plan":"1) Read the header row of Combined Aging (USD). 2) Add a Pivot - Aging tab. 3) Build a SUMIFS cross-tab.","actions":[],"contextRequest":{"sheet":"Combined Aging (USD)","range":"A1:Q6"},"decision":"CONTINUE"}\n```');
  assert.equal(envelope.plan.goal.startsWith('1) Read the header row'), true);
  assert.equal(envelope.contextRequest.sheet, 'Combined Aging (USD)');
  assert.equal(envelope.contextRequest.range, 'A1:Q6');
  assert.equal(envelope.actions.length, 0);
  assert.equal(envelope.decision, 'CONTINUE');
  assert.equal(AgentCore.normalizePlan('   '), null);
});

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

test('Auto never routes workbook data to a provider without a key', () => {
  const manual = AgentCore.routeProvider({ selected: 'qwen', mode: 'agent', fast: false, text: 'Reconcile debits', ready: { qwen: true, grok: true } });
  assert.equal(manual.provider, 'qwen');
  assert.equal(manual.reason, 'manual');
  const missing = AgentCore.routeProvider({ selected: 'grok', mode: 'chat', fast: false, text: 'Hi', ready: { qwen: true, grok: false } });
  assert.equal(missing.provider, 'grok');
  assert.match(missing.error, /Grok/);
  const finance = AgentCore.routeProvider({ selected: 'auto', mode: 'agent', fast: false, text: 'Reconcile the ledger', ready: { qwen: true, grok: false } });
  assert.equal(finance.provider, 'qwen');
  assert.equal(finance.fallback, '');
  const both = AgentCore.routeProvider({ selected: 'auto', mode: 'agent', fast: false, text: 'Reconcile the ledger', ready: { qwen: true, grok: true } });
  assert.equal(both.provider, 'grok');
  assert.equal(both.fallback, 'qwen');
  const quick = AgentCore.routeProvider({ selected: 'auto', mode: 'chat', fast: true, text: 'What is in A1?', ready: { qwen: true, grok: true } });
  assert.equal(quick.provider, 'qwen');
});

test('debit, blanks, duplicates, and formula errors are calculated from the grid', () => {
  const balanced = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit' }, { values: [[100, 40], [20, 80]] });
  assert.equal(balanced.status, 'PASS');
  assert.equal(balanced.difference, '0');
  const open = AgentCore.evaluateGridCheck({ check: 'debitEqualsCredit' }, { values: [[100, 40]] });
  assert.equal(open.status, 'FAIL');
  assert.equal(open.difference, '60');
  const blanks = AgentCore.evaluateGridCheck({ check: 'blankRequired', columns: [2] }, { values: [['Acme', ''], ['', '']] });
  assert.equal(blanks.status, 'FAIL');
  assert.equal(blanks.details.length, 1);
  assert.equal(blanks.details[0], 'R1C2');
  const dupes = AgentCore.evaluateGridCheck({ check: 'duplicateKeys', columns: [1] }, { values: [['A-1', 5], ['A-1', 9], ['B-2', 1]] });
  assert.equal(dupes.status, 'FAIL');
  const errors = AgentCore.evaluateGridCheck({ check: 'formulaErrors' }, { values: [[10, '#REF!']], formulas: [['=A1', '#REF!']] });
  assert.equal(errors.status, 'FAIL');
  const explained = AgentCore.explainFailure(open);
  assert.equal(explained.issue, 'debitEqualsCredit');
  assert.match(explained.recommendedAction, /draft/i);
});

test('a failed finance check stops the loop even if the model says continue', () => {
  const stopped = AgentCore.decide({ decision: 'CONTINUE', validations: [{ status: 'FAIL' }], iteration: 0, maxLoops: 4 });
  assert.equal(stopped, 'FAILED');
  const limited = AgentCore.decide({ decision: 'REPLAN', validations: [], iteration: 3, maxLoops: 4 });
  assert.equal(limited, 'FAILED');
  const done = AgentCore.decide({ decision: '', validations: [{ status: 'PASS' }], iteration: 0, maxLoops: 4 });
  assert.equal(done, 'COMPLETE');
});

test('the audit summary keeps the trail and drops secrets', () => {
  const entry = AgentCore.audit({ kind: 'request', mode: 'agent', provider: 'qwen', reason: 'manual', apiKey: 'sk-secret', text: 'Reconcile' });
  assert.equal(entry.apiKey, undefined);
  const summary = AgentCore.summarize([entry, AgentCore.audit({ kind: 'validation', status: 'FAIL', check: 'debitEqualsCredit', before: '100', after: '40', difference: '60' })]);
  assert.match(summary, /Request \(agent\) via qwen/);
  assert.match(summary, /FAIL debitEqualsCredit/);
  assert.equal(summary.includes('sk-secret'), false);
});

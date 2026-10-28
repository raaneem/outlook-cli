// Unit tests for the shared machinery behind the message-changing commands:
// id validation, id input parsing, and the preflight-then-apply runner.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  validateMessageId,
  parseMessageIds,
  isMissingItem,
  runMessageBatch,
  MAX_BATCH_IDS,
} from '../src/actions.mjs';
import { AppError, E } from '../src/errors.mjs';

const isArgsError = (e) => e instanceof AppError && e.code === E.ARGS;
const httpError = (status, serverCode) =>
  new AppError({ code: E.HTTP, message: `HTTP ${status}`, status, serverCode });

// ---------------------------------------------------------------------------
// ids

test('validateMessageId accepts standard and URL-safe base64 ids, trimmed', () => {
  assert.equal(validateMessageId(' AAMkAGI2/abc+def= '), 'AAMkAGI2/abc+def=');
  assert.equal(validateMessageId('AAkALgAA_x-y'), 'AAkALgAA_x-y');
});

test('validateMessageId rejects garbage, and explains an InternetMessageId', () => {
  for (const bad of ['', '   ', 'has space', 'semi;colon', 42, null]) {
    assert.throws(() => validateMessageId(bad), isArgsError, `expected E_ARGS for ${bad}`);
  }
  assert.throws(
    () => validateMessageId('<CAF123@mail.example.com>'),
    (e) => isArgsError(e) && /InternetMessageId/.test(e.hint),
  );
});

test('parseMessageIds prefers arguments and drops duplicates', () => {
  assert.deepEqual(parseMessageIds(['A1', 'B2', 'A1'], '["IGNORED"]'), ['A1', 'B2']);
});

test('parseMessageIds reads a JSON array or whitespace-separated ids from STDIN', () => {
  assert.deepEqual(parseMessageIds([], '["A1","B2"]\n'), ['A1', 'B2']);
  assert.deepEqual(parseMessageIds([], 'A1\nB2  C3\n'), ['A1', 'B2', 'C3']);
});

test('parseMessageIds refuses no ids, broken JSON, and oversized batches', () => {
  assert.throws(() => parseMessageIds([], ''), (e) => isArgsError(e) && /No message ids/.test(e.message));
  assert.throws(() => parseMessageIds([], '["A1",'), isArgsError);
  const many = Array.from({ length: MAX_BATCH_IDS + 1 }, (_, i) => `ID${i}`);
  assert.throws(() => parseMessageIds(many), (e) => isArgsError(e) && /Too many/.test(e.message));
});

test('isMissingItem recognises not-found and wrong-mailbox errors only', () => {
  assert.equal(isMissingItem(httpError(404, 'ErrorItemNotFound')), true);
  assert.equal(isMissingItem(httpError(400, 'ErrorInvalidMailboxItemId')), true);
  assert.equal(isMissingItem(httpError(500, 'ErrorInternalServerError')), false);
  assert.equal(isMissingItem(new AppError({ code: E.NETWORK, message: 'down' })), false);
  assert.equal(isMissingItem(new Error('plain')), false);
});

// ---------------------------------------------------------------------------
// runMessageBatch

/** A batch over in-memory messages where `plan` sets Flag = true. */
function flagBatch(messages, { dryRun = false, applyFails = {}, loadFails = {} } = {}) {
  const applied = [];
  const run = runMessageBatch({
    command: 'test',
    ids: Object.keys(messages),
    mailbox: null,
    dryRun,
    load: async (id) => {
      if (loadFails[id]) throw loadFails[id];
      return messages[id];
    },
    plan: (message, id) =>
      message.Flag
        ? { skip: 'already-flagged' }
        : {
            before: { flag: false },
            after: { flag: true },
            apply: async () => {
              if (applyFails[id]) throw applyFails[id];
              applied.push(id);
              return { newId: `${id}-new` };
            },
          },
  });
  return run.then((result) => ({ result, applied }));
}

test('runMessageBatch changes, skips, and reports new ids', async () => {
  const { result, applied } = await flagBatch({
    A: { Subject: 'a' },
    B: { Subject: 'b', Flag: true },
  });
  assert.equal(result.ok, true);
  assert.equal(result.aborted, false);
  assert.deepEqual(applied, ['A']);
  assert.deepEqual(result.items, [
    { id: 'A', subject: 'a', status: 'changed', before: { flag: false }, after: { flag: true }, newId: 'A-new' },
    { id: 'B', subject: 'b', status: 'skipped', reason: 'already-flagged' },
  ]);
  assert.equal(result.counts.changed, 1);
  assert.equal(result.counts.skipped, 1);
});

test('runMessageBatch --dry-run plans every message but applies none', async () => {
  const { result, applied } = await flagBatch({ A: {}, B: {} }, { dryRun: true });
  assert.equal(result.mode, 'dry-run');
  assert.deepEqual(applied, []);
  assert.deepEqual(result.items.map((i) => i.status), ['would-change', 'would-change']);
  assert.equal(result.counts.wouldChange, 2);
});

test('runMessageBatch writes nothing when any preflight fails', async () => {
  const { result, applied } = await flagBatch(
    { A: {}, B: {}, C: {} },
    { loadFails: { B: httpError(404, 'ErrorItemNotFound') } },
  );
  assert.deepEqual(applied, []);
  assert.equal(result.ok, false);
  assert.equal(result.aborted, true);
  assert.deepEqual(result.items.map((i) => i.status), ['not-run', 'not-found', 'not-run']);
  assert.equal(result.items[1].error.serverCode, 'ErrorItemNotFound');
});

test('runMessageBatch keeps going after an HTTP failure but stops after a network one', async () => {
  const http = await flagBatch({ A: {}, B: {}, C: {} }, { applyFails: { A: httpError(500, 'Boom') } });
  assert.deepEqual(http.applied, ['B', 'C']);
  assert.deepEqual(http.result.items.map((i) => i.status), ['failed', 'changed', 'changed']);

  const down = new AppError({ code: E.NETWORK, message: 'down' });
  const net = await flagBatch({ A: {}, B: {}, C: {} }, { applyFails: { A: down } });
  assert.deepEqual(net.applied, []);
  assert.deepEqual(net.result.items.map((i) => i.status), ['failed', 'not-run', 'not-run']);
  assert.equal(net.result.counts.notRun, 2);
});

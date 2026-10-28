// Shared machinery for the commands that change existing messages
// (mark-read, mark-unread, move, archive, categorize).
//
// Each runs in two phases:
//   1. Preflight: GET every id in the target mailbox. That proves the id
//      exists *in that mailbox* (ids are mailbox-bound, LEARNINGS.md
//      section 16) and captures the state the change is planned from. If
//      any id fails, nothing is written -- a caller that approved a set of
//      messages should not get a silent partial run.
//   2. Apply: each message changes on its own. A failure is recorded and
//      the rest still run, unless the failure means every remaining write
//      would fail the same way (auth, network, writes disabled). Messages
//      already in the requested state are skipped, so re-running is safe.
//
// --dry-run stops after planning: only GETs reach the server.

import { AppError, E } from './errors.mjs';

/** Upper bound per call, so a runaway pipeline cannot touch a whole mailbox. */
export const MAX_BATCH_IDS = 100;

// Exchange REST ids are base64, in the standard or URL-safe alphabet.
const ID_RE = /^[A-Za-z0-9+/=_-]+$/;
const MAX_ID_LENGTH = 1024;

// Server codes that mean "no such item in this mailbox". ErrorInvalidMailboxItemId
// in particular means the id belongs to a *different* mailbox.
const MISSING_CODES = new Set([
  'ErrorItemNotFound',
  'ErrorInvalidMailboxItemId',
  'ErrorInvalidIdMalformed',
  'ErrorFolderNotFound',
]);

// After one of these, every remaining write would fail the same way.
const STOP_CODES = new Set([
  E.AUTH_REQUIRED,
  E.AUTH_BLOCKED,
  E.NETWORK,
  E.PRECONDITION,
  E.UNEXPECTED,
]);

/** Validate one message id; returns it trimmed. */
export function validateMessageId(raw) {
  const id = typeof raw === 'string' ? raw.trim() : '';
  if (/^<.+@.+>$/.test(id)) {
    throw new AppError({
      code: E.ARGS,
      message: `Not a message Id: ${id}`,
      hint: 'That is an InternetMessageId. Pass the Id field from `outlook list` or `outlook read`.',
    });
  }
  if (!id || id.length > MAX_ID_LENGTH || !ID_RE.test(id)) {
    throw new AppError({
      code: E.ARGS,
      message: `Invalid message Id: ${JSON.stringify(String(raw ?? '').slice(0, 80))}`,
      hint: 'Pass the Id field from `outlook list` or `outlook read`.',
    });
  }
  return id;
}

/**
 * Collect ids from positional arguments or, when there are none, from
 * STDIN: a JSON array of strings, or ids separated by whitespace/newlines
 * (so `jq -r '.value[].Id'` pipes straight in). Duplicates are dropped.
 */
export function parseMessageIds(args = [], stdinText = '') {
  let raw = args;
  if (raw.length === 0) {
    const text = stdinText.trim();
    if (text.startsWith('[')) {
      try {
        raw = JSON.parse(text);
      } catch (cause) {
        throw new AppError({
          code: E.ARGS,
          message: 'STDIN started like a JSON array but did not parse.',
          hint: 'Pipe a JSON array of id strings, or one id per line.',
          cause,
        });
      }
    } else {
      raw = text ? text.split(/\s+/) : [];
    }
  }
  if (raw.length === 0) {
    throw new AppError({
      code: E.ARGS,
      message: 'No message ids provided.',
      hint: 'Pass ids as arguments, or pipe them via STDIN (a JSON array or one per line).',
    });
  }
  const ids = [...new Set(raw.map(validateMessageId))];
  if (ids.length > MAX_BATCH_IDS) {
    throw new AppError({
      code: E.ARGS,
      message: `Too many message ids (${ids.length}); the limit is ${MAX_BATCH_IDS} per call.`,
      hint: 'Split the work into smaller batches.',
    });
  }
  return ids;
}

/** True when an error from a message or folder GET means "not in this mailbox". */
export function isMissingItem(err) {
  return (
    err instanceof AppError &&
    err.code === E.HTTP &&
    (err.status === 404 || MISSING_CODES.has(err.serverCode))
  );
}

function asAppError(err) {
  return err instanceof AppError
    ? err
    : new AppError({ code: E.UNEXPECTED, message: err?.message ?? String(err) });
}

function errorDetail(err) {
  return {
    code: err.code,
    status: err.status ?? null,
    serverCode: err.serverCode ?? null,
    message: err.message,
    hint: err.hint ?? null,
  };
}

/**
 * Run a per-message change over `ids` and return the JSON result.
 *
 * `plan(message, id)` returns `{ skip: reason }` for a no-op, or
 * `{ before, after, apply }` where `apply()` performs the write and may
 * resolve to `{ newId }` (a move changes the Id).
 *
 * Item statuses: changed | would-change (dry run) | skipped | not-found |
 * failed | not-run. `ok` is true only when nothing failed or was missing;
 * `aborted` is true when preflight failed and therefore nothing was written.
 *
 * @param {object} o
 * @param {string} o.command
 * @param {string[]} o.ids
 * @param {string|null} o.mailbox
 * @param {boolean} o.dryRun
 * @param {(id: string) => Promise<object>} o.load
 * @param {(message: object, id: string) => object} o.plan
 * @param {object} [o.extra] additional top-level fields (destination, change, ...)
 */
export async function runMessageBatch({ command, ids, mailbox, dryRun, load, plan, extra = {} }) {
  const items = new Map();
  const loaded = new Map();

  for (const id of ids) {
    try {
      loaded.set(id, await load(id));
    } catch (raw) {
      const err = asAppError(raw);
      items.set(id, {
        id,
        status: isMissingItem(err) ? 'not-found' : 'failed',
        error: errorDetail(err),
      });
      if (STOP_CODES.has(err.code)) break;
    }
  }
  const aborted = items.size > 0;
  let stopped = aborted;

  for (const id of ids) {
    if (items.has(id)) continue;
    const message = loaded.get(id);
    const base = { id, subject: message?.Subject ?? null };
    if (stopped || !message) {
      items.set(id, { ...base, status: 'not-run' });
      continue;
    }
    const step = plan(message, id);
    if (step.skip) {
      items.set(id, { ...base, status: 'skipped', reason: step.skip });
      continue;
    }
    if (dryRun) {
      items.set(id, { ...base, status: 'would-change', before: step.before, after: step.after });
      continue;
    }
    try {
      const result = await step.apply();
      items.set(id, {
        ...base,
        status: 'changed',
        before: step.before,
        after: step.after,
        ...(result?.newId ? { newId: result.newId } : {}),
      });
    } catch (raw) {
      const err = asAppError(raw);
      items.set(id, { ...base, status: 'failed', before: step.before, error: errorDetail(err) });
      if (STOP_CODES.has(err.code)) stopped = true;
    }
  }

  const list = ids.map((id) => items.get(id));
  const count = (status) => list.filter((i) => i.status === status).length;
  const counts = {
    requested: ids.length,
    changed: count('changed'),
    wouldChange: count('would-change'),
    skipped: count('skipped'),
    notFound: count('not-found'),
    failed: count('failed'),
    notRun: count('not-run'),
  };
  return {
    ok: counts.notFound === 0 && counts.failed === 0,
    command,
    mode: dryRun ? 'dry-run' : 'apply',
    mailbox: mailbox ?? null,
    ...extra,
    aborted,
    counts,
    items: list,
  };
}

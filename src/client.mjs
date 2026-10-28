// Pure-fetch wrapper around the Outlook REST v2 API.
//
// All Playwright-dependent code lives in capture.mjs. Keeping this module
// dependency-free makes it cheap to unit-test and means cache-hit calls
// don't pay the Playwright cold-load cost.
//
// `OUTLOOK_API_BASE` overrides the base URL — useful for tests pointing at
// a local mock server. `OUTLOOK_HTTP_TIMEOUT_MS` bounds each request so a
// hung connection can't stall the CLI forever.

import { AppError, E } from './errors.mjs';
import { debug } from './output.mjs';
import { DEFAULT_RESOURCE, resourceBase } from './resources.mjs';
import { anchorHeaders, mailboxBase } from './mailbox.mjs';

const DEFAULT_TIMEOUT_MS = 30_000;

// Statuses where the server explicitly rejected the request without
// processing it — safe to retry for any HTTP method.
const RETRYABLE_STATUSES = new Set([429, 503]);
const MAX_ATTEMPTS = 3;
const MAX_RETRY_WAIT_MS = 30_000;

function timeoutMs() {
  const v = Number(process.env.OUTLOOK_HTTP_TIMEOUT_MS);
  return Number.isFinite(v) && v > 0 ? v : DEFAULT_TIMEOUT_MS;
}

/** Milliseconds to wait before the next attempt, honouring Retry-After. */
function retryDelayMs(res, attempt) {
  const header = res?.headers?.get('retry-after');
  const secs = Number(header);
  if (Number.isFinite(secs) && secs >= 0) {
    return Math.min(secs * 1000, MAX_RETRY_WAIT_MS);
  }
  return 500 * 2 ** attempt; // 500ms, 1s
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** An env switch is on unless unset or "", "0", "false", "no". */
function envFlag(name) {
  const v = process.env[name];
  return !!v && !/^(0|false|no)$/i.test(v.trim());
}

// POSTs that only read: getSchedule (free-busy) is a lookup with a body.
const READ_ONLY_POSTS = /\/getSchedule$/i;
// Mail submission: /messages/{id}/send and /sendmail (any casing).
const SEND_ROUTES = /\/(send|sendmail)$/i;

/**
 * Process-wide brakes, enforced here so no command -- or helper script
 * that imports call() -- can forget them. OUTLOOK_READ_ONLY refuses every
 * request that could change a mailbox (dry runs of automation);
 * OUTLOOK_NO_SEND refuses only mail submission, so unattended automation
 * can organise mail but never send it.
 */
function assertWriteAllowed(method, path) {
  if (method === 'GET') return;
  const route = path.split('?')[0];
  if (envFlag('OUTLOOK_READ_ONLY') && !(method === 'POST' && READ_ONLY_POSTS.test(route))) {
    throw new AppError({
      code: E.PRECONDITION,
      message: `Refusing ${method} ${route}: OUTLOOK_READ_ONLY is set.`,
      hint: 'Unset OUTLOOK_READ_ONLY to allow changes, or preview with --dry-run.',
    });
  }
  if (envFlag('OUTLOOK_NO_SEND') && method === 'POST' && SEND_ROUTES.test(route)) {
    throw new AppError({
      code: E.PRECONDITION,
      message: `Refusing to send mail (POST ${route}): OUTLOOK_NO_SEND is set.`,
      hint: 'Unset OUTLOOK_NO_SEND to allow sending. Drafts can still be created.',
    });
  }
}

function networkError(cause, url) {
  const detail =
    cause?.name === 'TimeoutError'
      ? `Request timed out after ${timeoutMs()}ms`
      : (cause?.cause?.message ?? cause?.message ?? String(cause));
  return new AppError({
    code: E.NETWORK,
    message: `Could not reach ${new URL(url).host}: ${detail}`,
    hint: 'Check your network connection (VPN/proxy included) and retry. Increase OUTLOOK_HTTP_TIMEOUT_MS for slow links.',
    cause,
  });
}

/**
 * Call a Microsoft API with the captured headers.
 *
 * The base URL is chosen by `init.resource` (default "outlook"); pass e.g.
 * `{ resource: 'graph' }` to hit Graph with a Graph-audience token. `path`
 * is appended to that resource's base. `init.mailbox` (an SMTP address)
 * targets a shared mailbox instead of /me -- Outlook resource only.
 *
 * Transient failures are retried up to MAX_ATTEMPTS with backoff:
 *   - 429/503 responses (throttled/unavailable — not processed) for any method.
 *     `init.retryStatuses` narrows that set for a call that must not repeat
 *     (send-draft retries only 429).
 *   - Network-level errors (DNS, refused, timeout) for GET only, since a
 *     write that timed out mid-flight may already have been applied.
 * Anything that still fails maps to AppError(E.NETWORK); HTTP statuses are
 * returned to the caller untouched.
 */
export async function call(auth, path, init = {}) {
  // `resource`, `mailbox` and `retryStatuses` steer routing and retries;
  // they are not fetch options -- pull them out.
  const { resource: res_, mailbox, retryStatuses = RETRYABLE_STATUSES, ...fetchInit } = init;
  const key = res_ ?? DEFAULT_RESOURCE;
  if (mailbox && key !== DEFAULT_RESOURCE) {
    // Graph and Substrate address other users differently; silently
    // dropping the mailbox would act on the wrong one.
    throw new AppError({
      code: E.ARGS,
      message: `--mailbox applies to Outlook mail and calendar only, not ${key}.`,
      hint: 'Drop --mailbox for Graph, Teams and Substrate commands.',
    });
  }
  const base = mailbox ? mailboxBase(mailbox) : resourceBase(key);
  const url = `${base}${path}`;
  const method = fetchInit.method ?? 'GET';
  assertWriteAllowed(method, path);
  const headers = mailbox ? anchorHeaders(auth, mailbox) : auth;

  for (let attempt = 0; ; attempt++) {
    debug('fetch', method, url, attempt > 0 ? `(attempt ${attempt + 1})` : '');
    let res;
    try {
      res = await fetch(url, {
        ...fetchInit,
        headers: {
          ...headers,
          Accept: 'application/json',
          ...(fetchInit.headers ?? {}),
        },
        signal: AbortSignal.timeout(timeoutMs()),
      });
    } catch (cause) {
      if (method === 'GET' && attempt < MAX_ATTEMPTS - 1) {
        const wait = 500 * 2 ** attempt;
        debug(`network error (${cause?.name ?? 'Error'}); retrying in ${wait}ms`);
        await sleep(wait);
        continue;
      }
      throw networkError(cause, url);
    }

    if (retryStatuses.has(res.status) && attempt < MAX_ATTEMPTS - 1) {
      const wait = retryDelayMs(res, attempt);
      debug(`HTTP ${res.status}; retrying in ${wait}ms`);
      await res.text().catch(() => {}); // drain so the connection is reusable
      await sleep(wait);
      continue;
    }

    const text = await res.text();
    let body;
    try {
      body = JSON.parse(text);
    } catch {
      body = text;
    }
    return { status: res.status, body };
  }
}

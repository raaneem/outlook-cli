// Test helpers — JWT minting, fake token-cache seeding, mock HTTP server,
// and CLI subprocess runner.
//
// Everything here is deliberately small and dep-free. We rely on Node's
// built-in `node:test`, `node:http`, and `node:child_process`.

import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { mkdtempSync, writeFileSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
export const CLI_PATH = resolve(HERE, '..', 'src', 'cli.mjs');

/**
 * Build a JWT-shaped string. The signature segment is placeholder text —
 * the CLI never validates the signature, only decodes the payload to read
 * `exp`.
 */
export function makeFakeJwt({ expSeconds, claims = {} } = {}) {
  const enc = (obj) =>
    Buffer.from(JSON.stringify(obj))
      .toString('base64')
      .replace(/=+$/, '')
      .replace(/\+/g, '-')
      .replace(/\//g, '_');
  const header = enc({ typ: 'JWT', alg: 'none' });
  const payload = enc({
    aud: 'https://outlook.office.com/',
    exp: expSeconds,
    ...claims,
  });
  return `${header}.${payload}.test-signature-not-validated`;
}

/**
 * Create a fresh tmp directory and write a token cache file inside it whose
 * Bearer JWT expires `secondsFromNow` from now. Returns the file path.
 */
export function seedTokenCache({
  secondsFromNow = 3600,
  headers = {},
  claims = {},
} = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-test-'));
  const file = join(dir, 'auth.json');
  const exp = Math.floor(Date.now() / 1000) + secondsFromNow;
  const jwt = makeFakeJwt({ expSeconds: exp, claims });
  mkdirSync(dir, { recursive: true });
  writeFileSync(
    file,
    JSON.stringify({
      headers: {
        authorization: `Bearer ${jwt}`,
        'x-anchormailbox': 'PUID:test@test',
        ...headers,
      },
      expiresAt: exp * 1000,
      savedAt: Date.now(),
    }),
  );
  return file;
}

/**
 * Write a captured-headers JSON file for the OUTLOOK_CAPTURE_FIXTURE seam —
 * what capture.mjs would have produced from a real browser run. Lets tests
 * exercise the automatic 401-recapture path without Playwright.
 */
export function seedCaptureFixture({ secondsFromNow = 3600, claims = {} } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-fixture-'));
  const file = join(dir, 'captured-headers.json');
  const exp = Math.floor(Date.now() / 1000) + secondsFromNow;
  writeFileSync(
    file,
    JSON.stringify({
      authorization: `Bearer ${makeFakeJwt({ expSeconds: exp, claims })}`,
      'x-anchormailbox': 'PUID:fixture@test',
    }),
  );
  return file;
}

/** Return a unique tmp path for an isolated learnings file. */
export function freshLearningsPath() {
  return join(mkdtempSync(join(tmpdir(), 'outlook-learn-')), 'learnings.md');
}

/**
 * Run an HTTP server on a random port. `handler` receives (req, res) and is
 * responsible for the entire response. Returns `{ url, close }`.
 */
export function startMockServer(handler) {
  return new Promise((resolveStart) => {
    const server = createServer(handler);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      resolveStart({
        url: `http://127.0.0.1:${port}`,
        close: () =>
          new Promise((r) => {
            server.closeAllConnections?.();
            server.close(() => r());
          }),
      });
    });
  });
}

/**
 * In-memory stand-in for Exchange mailboxes, served over HTTP like
 * startMockServer, for the commands that change messages. `mailboxes` maps
 * "me" and SMTP addresses to `{ folders, messages, categories }`; requests
 * under /api/v2.0/me/... and /api/v2.0/users/{smtp}/... hit the matching
 * mailbox. Like Exchange, an id is bound to its mailbox (a foreign id gets
 * 400 ErrorInvalidMailboxItemId) and a moved or sent message gets a new
 * Id. It accepts any write it can perform -- refusing is the CLI's job, and
 * tests prove it by checking `writes()`.
 *
 * `faults` maps "METHOD /path" (path after the mailbox, ids decoded) to a
 * canned `{ status, body, headers }`, or to 'drop' to cut the connection.
 *
 * Folder:  { Id, DisplayName, ParentId?, WellKnownName? }
 * Message: { Id, Subject, ParentFolderId, IsRead?, Categories?, IsDraft?,
 *            ChangeKey?, ToRecipients?, Attachments?, ... }
 */
export async function startFakeMailbox(mailboxes, { faults = {} } = {}) {
  const boxes = structuredClone(mailboxes);
  const requests = [];
  let seq = 0;

  const reply = (res, status, body) => {
    res.writeHead(status, { 'Content-Type': 'application/json' });
    res.end(body === undefined ? '' : JSON.stringify(body));
  };
  const fail = (res, status, code) => reply(res, status, { error: { code, message: code } });
  const findMessage = (box, id) => box.messages.find((m) => m.Id === id);
  const folderBy = (box, ref) =>
    box.folders.find(
      (f) => f.Id === ref || f.WellKnownName?.toLowerCase() === String(ref).toLowerCase(),
    );
  const page = (list, query) => {
    const skip = Number(query.get('$skip') ?? 0);
    const top = Number(query.get('$top') ?? 10);
    return { value: list.slice(skip, skip + top) };
  };

  const server = await startMockServer((req, res) => {
    let raw = '';
    req.on('data', (d) => (raw += d));
    req.on('end', () => {
      const url = new URL(req.url, 'http://fake');
      const match = url.pathname.match(/^\/api\/v2\.0\/(me|users\/([^/]+))(\/.*)?$/);
      if (!match) return fail(res, 404, 'FakeUnknownBase');
      const key = match[1] === 'me' ? 'me' : decodeURIComponent(match[2]);
      const segs = (match[3] ?? '').split('/').filter(Boolean).map(decodeURIComponent);
      const path = `/${segs.join('/')}`;
      let body = null;
      try {
        body = raw ? JSON.parse(raw) : null;
      } catch {
        body = raw;
      }
      requests.push({
        method: req.method,
        mailbox: key,
        path,
        query: Object.fromEntries(url.searchParams),
        anchor: req.headers['x-anchormailbox'],
        body,
      });

      const fault = faults[`${req.method} ${path}`];
      if (fault === 'drop') return req.socket.destroy();
      if (fault) {
        res.writeHead(fault.status, { 'Content-Type': 'application/json', ...(fault.headers ?? {}) });
        return res.end(fault.body === undefined ? '' : JSON.stringify(fault.body));
      }

      const box = boxes[key];
      if (!box) return fail(res, 404, 'ErrorNonExistentMailbox');

      if (segs[0] === 'messages' && segs[1]) {
        const msg = findMessage(box, segs[1]);
        if (!msg) {
          const elsewhere = Object.values(boxes).some((b) => findMessage(b, segs[1]));
          return elsewhere
            ? fail(res, 400, 'ErrorInvalidMailboxItemId')
            : fail(res, 404, 'ErrorItemNotFound');
        }
        const action = segs[2];
        if (req.method === 'GET' && !action) return reply(res, 200, msg);
        if (req.method === 'GET' && action === 'attachments') {
          return reply(res, 200, { value: msg.Attachments ?? [] });
        }
        if (req.method === 'PATCH' && !action) {
          Object.assign(msg, body);
          msg.ChangeKey = `CK${++seq}`;
          return reply(res, 200, msg);
        }
        if (req.method === 'POST' && action === 'move') {
          const dest = folderBy(box, body?.DestinationId);
          if (!dest) return fail(res, 404, 'ErrorFolderNotFound');
          msg.ParentFolderId = dest.Id;
          msg.Id = `${msg.Id}M${++seq}`;
          return reply(res, 201, msg);
        }
        if (req.method === 'POST' && action === 'send') {
          msg.ParentFolderId = folderBy(box, 'SentItems')?.Id ?? null;
          msg.IsDraft = false;
          msg.Id = `${msg.Id}S${++seq}`;
          return reply(res, 202);
        }
      }

      if (req.method === 'GET' && segs[0] === 'mailfolders') {
        if (!segs[1]) return reply(res, 200, page(box.folders.filter((f) => !f.ParentId), url.searchParams));
        const folder = folderBy(box, segs[1]);
        if (!folder) return fail(res, 404, 'ErrorFolderNotFound');
        if (!segs[2]) return reply(res, 200, folder);
        if (segs[2] === 'childfolders') {
          return reply(res, 200, page(box.folders.filter((f) => f.ParentId === folder.Id), url.searchParams));
        }
      }

      if (req.method === 'GET' && path === '/MasterCategories') {
        return reply(res, 200, { value: box.categories ?? [] });
      }

      return fail(res, 501, 'FakeRouteNotImplemented');
    });
  });

  return {
    ...server,
    requests,
    boxes,
    /** Every request that was not a GET -- what a dry run must leave empty. */
    writes: () => requests.filter((r) => r.method !== 'GET'),
  };
}

/**
 * Environment every CLI run starts from: layered over process.env, under
 * the test's own `env`. A developer machine has a real token cache and a
 * real network, so a test that forgets to pin a cache path or an API base
 * must land on an empty tmp dir and a closed local port -- never read the
 * real token, and never reach Microsoft with it.
 */
function isolatedEnv() {
  const root = mkdtempSync(join(tmpdir(), 'outlook-run-'));
  const dead = 'http://127.0.0.1:9';
  return {
    OUTLOOK_NO_CAPTURE: '1',
    OUTLOOK_CAPTURE_FIXTURE: '',
    XDG_CACHE_HOME: join(root, 'cache'),
    XDG_DATA_HOME: join(root, 'data'),
    OUTLOOK_TOKEN_CACHE: join(root, 'cache', 'auth.json'),
    OUTLOOK_TOKEN_CACHE_GRAPH: join(root, 'cache', 'auth-graph.json'),
    OUTLOOK_TOKEN_CACHE_SUBSTRATE: join(root, 'cache', 'auth-substrate.json'),
    OUTLOOK_LEARNINGS: join(root, 'data', 'learnings.md'),
    OUTLOOK_PROFILE: join(root, 'data', 'browser-profile'),
    OUTLOOK_API_BASE: `${dead}/api/v2.0/me`,
    OUTLOOK_GRAPH_BASE: `${dead}/graph/v1.0`,
    OUTLOOK_SUBSTRATE_BASE: `${dead}/substrate`,
    OUTLOOK_READ_ONLY: '',
    OUTLOOK_NO_SEND: '',
    OUTLOOK_DEBUG: '',
    DEBUG: '',
  };
}

/**
 * Spawn the CLI as a subprocess. Returns `{ code, stdout, stderr }`.
 * Pass `env` to add/override environment variables for that run.
 *
 * OUTLOOK_NO_CAPTURE is set by default so no test can ever fall through to
 * a real Playwright/Chromium launch; override it (or set
 * OUTLOOK_CAPTURE_FIXTURE) in tests that exercise the recapture path.
 */
export function runCli(args, { env = {}, stdin = null, timeoutMs = 15_000 } = {}) {
  return new Promise((resolveRun, reject) => {
    const child = spawn(process.execPath, [CLI_PATH, ...args], {
      env: { ...process.env, ...isolatedEnv(), ...env },
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d) => (stdout += d));
    child.stderr.on('data', (d) => (stderr += d));

    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`CLI run timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    child.on('close', (code) => {
      clearTimeout(timer);
      resolveRun({ code, stdout, stderr });
    });
    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });

    if (stdin !== null) {
      child.stdin.write(stdin);
    }
    child.stdin.end();
  });
}

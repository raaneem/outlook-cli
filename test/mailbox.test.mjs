// Shared-mailbox routing (`--mailbox`). These pin the wire shape the
// shared-mailbox workflows rely on -- /users/{smtp} paths and exactly one
// SMTP: anchor -- so a refactor or an upstream merge cannot silently send
// shared-mailbox calls to /me.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startMockServer } from './helpers.mjs';
import { parseMailbox, mailboxBase, anchorHeaders } from '../src/mailbox.mjs';

const SHARED = 'shared@example.com';
const SHARED_PATH = '/api/v2.0/users/shared%40example.com';

/** Mock that records every request; `reply(entry)` may override the answer. */
async function recordingMock(reply = () => ({ status: 200, body: { value: [] } })) {
  const seen = [];
  const mock = await startMockServer((req, res) => {
    let body = '';
    req.on('data', (d) => (body += d));
    req.on('end', () => {
      const entry = {
        method: req.method,
        url: req.url,
        anchor: req.headers['x-anchormailbox'],
        body,
      };
      seen.push(entry);
      const { status = 200, body: out = {} } = reply(entry) ?? {};
      res.writeHead(status, { 'Content-Type': 'application/json' });
      res.end(out === null ? '' : JSON.stringify(out));
    });
  });
  return { ...mock, seen };
}

function envFor(mock) {
  return {
    OUTLOOK_TOKEN_CACHE: seedTokenCache(),
    OUTLOOK_API_BASE: `${mock.url}/api/v2.0/me`,
  };
}

// ---------------------------------------------------------------------------
// Unit

test('parseMailbox lower-cases a valid address and returns null when absent', () => {
  assert.equal(parseMailbox(undefined), null);
  assert.equal(parseMailbox(' Shared@Example.COM '), 'shared@example.com');
});

test('parseMailbox rejects values that are not an address', () => {
  for (const bad of ['', 'shared', 'shared@', '@example.com', 'a b@example.com']) {
    assert.throws(
      () => parseMailbox(bad),
      (e) => e.code === 'E_ARGS',
      `expected E_ARGS for ${JSON.stringify(bad)}`,
    );
  }
});

test('mailboxBase swaps /me for /users/{smtp} on the configured base', () => {
  const prev = process.env.OUTLOOK_API_BASE;
  try {
    process.env.OUTLOOK_API_BASE = 'http://127.0.0.1:9/api/v2.0/me';
    assert.equal(mailboxBase(SHARED), `http://127.0.0.1:9${SHARED_PATH}`);
  } finally {
    if (prev === undefined) delete process.env.OUTLOOK_API_BASE;
    else process.env.OUTLOOK_API_BASE = prev;
  }
});

test('anchorHeaders replaces the captured anchor whatever its casing', () => {
  const out = anchorHeaders(
    { authorization: 'Bearer x', 'x-anchormailbox': 'PUID:me', 'X-AnchorMailbox': 'PUID:me' },
    SHARED,
  );
  const anchors = Object.keys(out).filter((k) => k.toLowerCase() === 'x-anchormailbox');
  assert.deepEqual(anchors, ['x-anchormailbox']);
  assert.equal(out['x-anchormailbox'], 'SMTP:shared@example.com');
  assert.equal(out.authorization, 'Bearer x');
});

// ---------------------------------------------------------------------------
// Integration

test('--mailbox routes a shared-mailbox list to /users/{smtp} with one SMTP anchor', async () => {
  const mock = await recordingMock();
  try {
    // The shape a triage bot uses to read a shared inbox.
    const { code } = await runCli(
      [
        '--mailbox', SHARED,
        'list', '--folder', 'Inbox', '--top', '25',
        '--order-by', 'ReceivedDateTime desc',
        '--select', 'Id,Subject,From',
      ],
      { env: envFor(mock) },
    );
    assert.equal(code, 0);
    assert.equal(mock.seen.length, 1);
    const [req] = mock.seen;
    assert.ok(req.url.startsWith(`${SHARED_PATH}/mailfolders/Inbox/messages?`), req.url);
    assert.match(req.url, /\$select=Id%2CSubject%2CFrom/);
    // A duplicated header would arrive comma-joined, so equality also
    // proves the captured PUID anchor was replaced, not added to.
    assert.equal(req.anchor, 'SMTP:shared@example.com');
  } finally {
    await mock.close();
  }
});

test('--mailbox works after the subcommand and in --mailbox= form', async () => {
  const mock = await recordingMock();
  try {
    const after = await runCli(['list', '--mailbox', SHARED], { env: envFor(mock) });
    const inline = await runCli(['list', `--mailbox=${SHARED}`], { env: envFor(mock) });
    assert.equal(after.code, 0);
    assert.equal(inline.code, 0);
    for (const req of mock.seen) {
      assert.ok(req.url.startsWith(`${SHARED_PATH}/messages?`), req.url);
      assert.equal(req.anchor, 'SMTP:shared@example.com');
    }
  } finally {
    await mock.close();
  }
});

test('without --mailbox calls stay on /me with the captured anchor', async () => {
  const mock = await recordingMock();
  try {
    const { code } = await runCli(['list'], { env: envFor(mock) });
    assert.equal(code, 0);
    assert.ok(mock.seen[0].url.startsWith('/api/v2.0/me/messages?'));
    assert.equal(mock.seen[0].anchor, 'PUID:test@test');
  } finally {
    await mock.close();
  }
});

test('read with --mailbox fetches the message from the shared mailbox', async () => {
  const mock = await recordingMock(() => ({ body: { Id: 'AAMk/abc' } }));
  try {
    const { code } = await runCli(['--mailbox', SHARED, 'read', 'AAMk/abc'], { env: envFor(mock) });
    assert.equal(code, 0);
    assert.equal(mock.seen[0].url, `${SHARED_PATH}/messages/AAMk%2Fabc`);
  } finally {
    await mock.close();
  }
});

test('draft-reply with --mailbox keeps every call in the shared mailbox and says so', async () => {
  const mock = await recordingMock((req) =>
    req.url.endsWith('/createReply')
      ? { status: 201, body: { Id: 'DRAFT-1', WebLink: 'https://outlook.office.com/owa/?ItemID=DRAFT-1' } }
      : { body: {} },
  );
  try {
    const override = JSON.stringify({
      Body: { ContentType: 'Text', Content: 'on it' },
      CcRecipients: [{ EmailAddress: { Address: 'cc@example.com' } }],
    });
    const { code, stdout } = await runCli(['--mailbox', SHARED, 'draft-reply', 'ORIG-1', override], {
      env: envFor(mock),
    });
    assert.equal(code, 0);
    assert.deepEqual(
      mock.seen.map((r) => `${r.method} ${r.url}`),
      [`POST ${SHARED_PATH}/messages/ORIG-1/createReply`, `PATCH ${SHARED_PATH}/messages/DRAFT-1`],
    );
    const out = JSON.parse(stdout);
    assert.equal(out.DraftId, 'DRAFT-1');
    assert.equal(out.Mailbox, SHARED);
    assert.match(out.message, /Drafts folder of shared@example\.com/);
  } finally {
    await mock.close();
  }
});

test('send with --mailbox posts to the shared mailbox sendmail', async () => {
  const mock = await recordingMock(() => ({ status: 202, body: null }));
  try {
    const message = JSON.stringify({
      Subject: 'hi',
      Body: { ContentType: 'Text', Content: 'x' },
      ToRecipients: [{ EmailAddress: { Address: 'to@example.com' } }],
    });
    const { code } = await runCli(['--mailbox', SHARED, 'send', message], { env: envFor(mock) });
    assert.equal(code, 0);
    assert.equal(mock.seen[0].method, 'POST');
    assert.equal(mock.seen[0].url, `${SHARED_PATH}/sendmail`);
  } finally {
    await mock.close();
  }
});

test('graph --resource outlook honours --mailbox', async () => {
  const mock = await recordingMock();
  try {
    const { code } = await runCli(
      ['--mailbox', SHARED, 'graph', '/mailfolders', '--resource', 'outlook'],
      { env: envFor(mock) },
    );
    assert.equal(code, 0);
    assert.equal(mock.seen[0].url, `${SHARED_PATH}/mailfolders`);
  } finally {
    await mock.close();
  }
});

test('an invalid --mailbox exits 64 before any request', async () => {
  const mock = await recordingMock();
  try {
    const { code, stderr } = await runCli(['--mailbox', 'not-an-address', 'list'], {
      env: envFor(mock),
    });
    assert.equal(code, 64);
    assert.match(stderr, /E_ARGS/);
    assert.match(stderr, /Invalid --mailbox/);
    assert.equal(mock.seen.length, 0);
  } finally {
    await mock.close();
  }
});

test('--mailbox with a Graph or Teams command exits 64 instead of being ignored', async () => {
  for (const args of [['graph', '/me'], ['teams']]) {
    const { code, stderr } = await runCli(['--mailbox', SHARED, ...args], {
      env: { OUTLOOK_TOKEN_CACHE_GRAPH: seedTokenCache({ claims: { aud: 'https://graph.microsoft.com' } }) },
    });
    assert.equal(code, 64, `expected exit 64 for ${args.join(' ')}`);
    assert.match(stderr, /--mailbox applies to Outlook mail and calendar only/);
  }
});

// mark-read / mark-unread, driven through the CLI against the in-memory
// fake mailbox. These also pin the behaviour every message-changing
// command shares: preflight, --dry-run, partial failure, STDIN ids,
// shared-mailbox routing, and the OUTLOOK_READ_ONLY brake.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startFakeMailbox } from './helpers.mjs';

const SHARED = 'shared@example.com';

function mailbox() {
  return {
    folders: [{ Id: 'F-INBOX', DisplayName: 'Inbox', WellKnownName: 'Inbox' }],
    messages: [
      { Id: 'MSG1', Subject: 'one', ParentFolderId: 'F-INBOX', IsRead: false },
      { Id: 'MSG2', Subject: 'two', ParentFolderId: 'F-INBOX', IsRead: true },
      { Id: 'MSG3', Subject: 'three', ParentFolderId: 'F-INBOX', IsRead: false },
    ],
  };
}

function sharedMailbox() {
  return {
    folders: [{ Id: 'S-INBOX', DisplayName: 'Inbox', WellKnownName: 'Inbox' }],
    messages: [{ Id: 'SHARED1', Subject: 'shared', ParentFolderId: 'S-INBOX', IsRead: false }],
  };
}

async function withFake(boxes, fn, opts) {
  const fake = await startFakeMailbox(boxes, opts);
  try {
    return await fn(fake, {
      OUTLOOK_TOKEN_CACHE: seedTokenCache(),
      OUTLOOK_API_BASE: `${fake.url}/api/v2.0/me`,
    });
  } finally {
    await fake.close();
  }
}

test('mark-read marks unread messages and skips ones already read', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout, stderr } = await runCli(['mark-read', 'MSG1', 'MSG2'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.equal(out.ok, true);
    assert.equal(out.command, 'mark-read');
    assert.equal(out.mode, 'apply');
    assert.equal(out.mailbox, null);
    assert.deepEqual(
      out.items.map((i) => [i.id, i.status]),
      [['MSG1', 'changed'], ['MSG2', 'skipped']],
    );
    assert.deepEqual(out.items[0].after, { isRead: true });
    assert.deepEqual(fake.writes().map((w) => [w.method, w.path, w.body]), [
      ['PATCH', '/messages/MSG1', { IsRead: true }],
    ]);
    assert.equal(fake.boxes.me.messages[0].IsRead, true);
  });
});

test('mark-unread is the mirror image', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['mark-unread', 'MSG1', 'MSG2'], { env });
    assert.equal(code, 0);
    const out = JSON.parse(stdout);
    assert.deepEqual(out.items.map((i) => i.status), ['skipped', 'changed']);
    assert.deepEqual(fake.writes().map((w) => w.body), [{ IsRead: false }]);
  });
});

test('mark-read --dry-run only reads', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['mark-read', 'MSG1', 'MSG3', '--dry-run'], { env });
    assert.equal(code, 0);
    const out = JSON.parse(stdout);
    assert.equal(out.mode, 'dry-run');
    assert.deepEqual(out.items.map((i) => i.status), ['would-change', 'would-change']);
    assert.equal(fake.writes().length, 0);
    assert.equal(fake.boxes.me.messages[0].IsRead, false);
  });
});

test('mark-read --mailbox acts on the shared mailbox only', async () => {
  await withFake({ me: mailbox(), [SHARED]: sharedMailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['--mailbox', SHARED, 'mark-read', 'SHARED1'], { env });
    assert.equal(code, 0);
    assert.equal(JSON.parse(stdout).mailbox, SHARED);
    assert.ok(fake.requests.every((r) => r.mailbox === SHARED && r.anchor === `SMTP:${SHARED}`));
    assert.equal(fake.boxes[SHARED].messages[0].IsRead, true);
  });
});

test('an id from another mailbox exits 4, changes nothing, and points at --mailbox', async () => {
  await withFake({ me: mailbox(), [SHARED]: sharedMailbox() }, async (fake, env) => {
    // Shared-mailbox id, but --mailbox forgotten.
    const { code, stdout, stderr } = await runCli(['mark-read', 'MSG1', 'SHARED1'], { env });
    assert.equal(code, 4);
    assert.match(stderr, /E_PRECONDITION/);
    assert.match(stderr, /nothing was changed/);
    assert.match(stderr, /--mailbox/);
    const out = JSON.parse(stdout);
    assert.equal(out.aborted, true);
    assert.deepEqual(out.items.map((i) => i.status), ['not-run', 'not-found']);
    assert.equal(out.items[1].error.serverCode, 'ErrorInvalidMailboxItemId');
    assert.equal(fake.writes().length, 0, 'the valid id must not be changed either');
  });
});

test('a failed write is reported per message and exits 3; the others still change', async () => {
  const faults = { 'PATCH /messages/MSG1': { status: 500, body: { error: { code: 'ErrorInternalServerError', message: 'boom' } } } };
  await withFake(
    { me: mailbox() },
    async (fake, env) => {
      const { code, stdout, stderr } = await runCli(['mark-read', 'MSG1', 'MSG3'], { env });
      assert.equal(code, 3);
      assert.match(stderr, /1 of 2 message\(s\) failed/);
      const out = JSON.parse(stdout);
      assert.equal(out.ok, false);
      assert.deepEqual(out.items.map((i) => i.status), ['failed', 'changed']);
      assert.equal(out.items[0].error.status, 500);
      assert.equal(fake.boxes.me.messages[2].IsRead, true);
    },
    { faults },
  );
});

test('mark-read takes ids from STDIN, as a JSON array or one per line', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const json = await runCli(['mark-read', '--dry-run'], { env, stdin: '["MSG1","MSG3"]' });
    const lines = await runCli(['mark-read', '--dry-run'], { env, stdin: 'MSG1\nMSG3\n' });
    for (const run of [json, lines]) {
      assert.equal(run.code, 0);
      assert.deepEqual(JSON.parse(run.stdout).items.map((i) => i.id), ['MSG1', 'MSG3']);
    }
  });
});

test('mark-read with no ids or a malformed id exits 64 before any request', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const none = await runCli(['mark-read'], { env });
    const bad = await runCli(['mark-read', 'not an id'], { env });
    assert.equal(none.code, 64);
    assert.match(none.stderr, /No message ids/);
    assert.equal(bad.code, 64);
    assert.match(bad.stderr, /Invalid message Id/);
    assert.equal(fake.requests.length, 0);
  });
});

test('OUTLOOK_READ_ONLY lets preflight read but refuses the write (exit 4)', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout, stderr } = await runCli(['mark-read', 'MSG1', 'MSG3'], {
      env: { ...env, OUTLOOK_READ_ONLY: '1' },
    });
    assert.equal(code, 4);
    assert.match(stderr, /OUTLOOK_READ_ONLY/);
    assert.equal(fake.writes().length, 0);
    assert.deepEqual(JSON.parse(stdout).items.map((i) => i.status), ['failed', 'not-run']);
  });
});

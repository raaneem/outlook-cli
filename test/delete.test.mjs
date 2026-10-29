// delete (a move into Deleted Items) and the discard-draft guard, driven
// through the CLI against the in-memory fake mailbox. The point of both:
// received mail is only ever deleted recoverably; the permanent DELETE is
// reserved for unsent drafts.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startFakeMailbox } from './helpers.mjs';

const SHARED = 'shared@example.com';

function mailbox() {
  return {
    folders: [
      { Id: 'F-INBOX', DisplayName: 'Inbox', WellKnownName: 'Inbox' },
      { Id: 'F-DELETED', DisplayName: 'Deleted Items', WellKnownName: 'DeletedItems' },
      { Id: 'F-DRAFTS', DisplayName: 'Drafts', WellKnownName: 'Drafts' },
    ],
    messages: [
      { Id: 'MSG1', Subject: 'one', ParentFolderId: 'F-INBOX', IsDraft: false },
      { Id: 'MSG2', Subject: 'two', ParentFolderId: 'F-DELETED', IsDraft: false },
      { Id: 'MSG3', Subject: 'three', ParentFolderId: 'F-INBOX', IsDraft: false },
      { Id: 'DRAFT1', Subject: 'draft', ParentFolderId: 'F-DRAFTS', IsDraft: true },
    ],
  };
}

async function withFake(boxes, fn) {
  const fake = await startFakeMailbox(boxes);
  try {
    return await fn(fake, {
      OUTLOOK_TOKEN_CACHE: seedTokenCache(),
      OUTLOOK_API_BASE: `${fake.url}/api/v2.0/me`,
    });
  } finally {
    await fake.close();
  }
}

test('delete moves messages to Deleted Items, returns new ids, and skips ones already there', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout, stderr } = await runCli(['delete', 'MSG1', 'MSG2'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.equal(out.command, 'delete');
    assert.deepEqual(out.destination, {
      id: 'F-DELETED',
      displayName: 'Deleted Items',
      wellKnownName: 'DeletedItems',
    });
    assert.deepEqual(out.items.map((i) => [i.id, i.status]), [['MSG1', 'changed'], ['MSG2', 'skipped']]);
    const moved = fake.boxes.me.messages.find((m) => m.Subject === 'one');
    assert.equal(moved.ParentFolderId, 'F-DELETED');
    assert.equal(out.items[0].newId, moved.Id);
    // Recoverable: a move, never an HTTP DELETE.
    assert.deepEqual(fake.writes().map((w) => [w.method, w.path, w.body]), [
      ['POST', '/messages/MSG1/move', { DestinationId: 'F-DELETED' }],
    ]);
  });
});

test('delete --dry-run plans without moving', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['delete', 'MSG1', 'MSG3', '--dry-run'], { env });
    assert.equal(code, 0);
    assert.deepEqual(JSON.parse(stdout).items.map((i) => i.status), ['would-change', 'would-change']);
    assert.equal(fake.writes().length, 0);
  });
});

test('delete --mailbox deletes inside the shared mailbox only', async () => {
  const shared = mailbox();
  shared.messages = [{ Id: 'SHARED1', Subject: 'shared', ParentFolderId: 'F-INBOX' }];
  await withFake({ me: mailbox(), [SHARED]: shared }, async (fake, env) => {
    const { code, stdout } = await runCli(['--mailbox', SHARED, 'delete', 'SHARED1'], { env });
    assert.equal(code, 0);
    assert.equal(JSON.parse(stdout).mailbox, SHARED);
    assert.ok(fake.requests.every((r) => r.mailbox === SHARED && r.anchor === `SMTP:${SHARED}`));
    assert.equal(fake.boxes[SHARED].messages[0].ParentFolderId, 'F-DELETED');
    assert.equal(fake.boxes.me.messages[0].ParentFolderId, 'F-INBOX');
  });
});

test('discard-draft removes an unsent draft', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['discard-draft', 'DRAFT1'], { env });
    assert.equal(code, 0);
    assert.deepEqual(JSON.parse(stdout), { discarded: true, DraftId: 'DRAFT1' });
    assert.equal(fake.boxes.me.messages.some((m) => m.Id === 'DRAFT1'), false);
  });
});

test('discard-draft refuses a received message (exit 4) and points at delete', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stderr } = await runCli(['discard-draft', 'MSG1'], { env });
    assert.equal(code, 4);
    assert.match(stderr, /Not a draft: MSG1/);
    assert.match(stderr, /outlook delete/);
    assert.equal(fake.writes().length, 0);
  });
});

test('discard-draft with a missing or foreign id exits 4 without deleting', async () => {
  const shared = mailbox();
  shared.messages = [{ Id: 'SHAREDDRAFT1', Subject: 'x', ParentFolderId: 'F-DRAFTS', IsDraft: true }];
  await withFake({ me: mailbox(), [SHARED]: shared }, async (fake, env) => {
    const missing = await runCli(['discard-draft', 'NOPE1'], { env });
    assert.equal(missing.code, 4);
    assert.match(missing.stderr, /Draft not found in your mailbox/);

    const foreign = await runCli(['discard-draft', 'SHAREDDRAFT1'], { env });
    assert.equal(foreign.code, 4);
    assert.match(foreign.stderr, /different mailbox/);
    assert.equal(fake.writes().length, 0);
  });
});

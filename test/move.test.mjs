// move / archive (and `folders --parent`), driven through the CLI against
// the in-memory fake mailbox: destination resolution, new ids after a
// move, skips, dry runs, and shared-mailbox routing.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startFakeMailbox } from './helpers.mjs';

const SHARED = 'shared@example.com';
// Long enough to be taken for a folder Id rather than a name.
const PROJECTS_ID = 'AAMkAGFakeProjectsFolderId00000000000000=';

function mailbox({ withArchive = true } = {}) {
  return {
    folders: [
      { Id: 'F-INBOX', DisplayName: 'Inbox', WellKnownName: 'Inbox' },
      ...(withArchive ? [{ Id: 'F-ARCHIVE', DisplayName: 'Archive', WellKnownName: 'Archive' }] : []),
      { Id: 'F-DELETED', DisplayName: 'Deleted Items', WellKnownName: 'DeletedItems' },
      { Id: PROJECTS_ID, DisplayName: 'Projects', ParentId: 'F-INBOX' },
      { Id: 'F-2026', DisplayName: '2026', ParentId: PROJECTS_ID },
      { Id: 'F-CLIENTS-1', DisplayName: 'Clients' },
      { Id: 'F-CLIENTS-2', DisplayName: 'clients' },
    ],
    messages: [
      { Id: 'MSG1', Subject: 'one', ParentFolderId: 'F-INBOX' },
      { Id: 'MSG2', Subject: 'two', ParentFolderId: 'F-ARCHIVE' },
      { Id: 'MSG3', Subject: 'three', ParentFolderId: 'F-INBOX' },
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

test('archive moves to the well-known Archive folder, returns the new id, skips ones already there', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout, stderr } = await runCli(['archive', 'MSG1', 'MSG2'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.deepEqual(out.destination, { id: 'F-ARCHIVE', displayName: 'Archive', wellKnownName: 'Archive' });
    assert.deepEqual(out.items.map((i) => [i.id, i.status]), [['MSG1', 'changed'], ['MSG2', 'skipped']]);
    assert.deepEqual(out.items[0].before, { folderId: 'F-INBOX' });
    assert.deepEqual(out.items[0].after, { folderId: 'F-ARCHIVE' });
    // The fake, like Exchange, gives a moved message a new Id.
    const moved = fake.boxes.me.messages.find((m) => m.Subject === 'one');
    assert.equal(out.items[0].newId, moved.Id);
    assert.notEqual(moved.Id, 'MSG1');
    assert.ok(fake.requests.some((r) => r.method === 'GET' && r.path === '/mailfolders/Archive'));
    assert.deepEqual(fake.writes().map((w) => [w.method, w.path, w.body]), [
      ['POST', '/messages/MSG1/move', { DestinationId: 'F-ARCHIVE' }],
    ]);
  });
});

test('archive in a mailbox without an Archive folder exits 64 before reading any message', async () => {
  await withFake({ me: mailbox({ withArchive: false }) }, async (fake, env) => {
    const { code, stderr } = await runCli(['archive', 'MSG1'], { env });
    assert.equal(code, 64);
    assert.match(stderr, /has no Archive folder/);
    assert.deepEqual(fake.requests.map((r) => r.path), ['/mailfolders/Archive']);
  });
});

test('move --to walks a folder path one level at a time, case-insensitively', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['move', 'MSG1', '--to', 'inbox/projects/2026'], { env });
    assert.equal(code, 0);
    assert.deepEqual(JSON.parse(stdout).destination, { id: 'F-2026', displayName: '2026', wellKnownName: null });
    assert.deepEqual(fake.writes()[0].body, { DestinationId: 'F-2026' });
  });
});

test('move --to accepts a folder Id', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['move', 'MSG1', '--to', PROJECTS_ID], { env });
    assert.equal(code, 0);
    assert.equal(JSON.parse(stdout).destination.displayName, 'Projects');
  });
});

test('move --to Deleted is a soft delete into Deleted Items', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code } = await runCli(['move', 'MSG1', '--to', 'Deleted'], { env });
    assert.equal(code, 0);
    assert.deepEqual(fake.writes()[0].body, { DestinationId: 'F-DELETED' });
  });
});

test('move --to an unknown or ambiguous folder exits 64 and changes nothing', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const unknown = await runCli(['move', 'MSG1', '--to', 'Nope'], { env });
    assert.equal(unknown.code, 64);
    assert.match(unknown.stderr, /Folder not found in your mailbox: Nope/);
    assert.match(unknown.stderr, /Inbox\/Nope/);

    const ambiguous = await runCli(['move', 'MSG1', '--to', 'clients'], { env });
    assert.equal(ambiguous.code, 64);
    assert.match(ambiguous.stderr, /ambiguous/);
    assert.match(ambiguous.stderr, /F-CLIENTS-1, F-CLIENTS-2/);

    assert.equal(fake.writes().length, 0);
  });
});

test('move --dry-run plans without moving', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['move', 'MSG1', 'MSG3', '--to', 'Archive', '--dry-run'], { env });
    assert.equal(code, 0);
    const out = JSON.parse(stdout);
    assert.equal(out.mode, 'dry-run');
    assert.deepEqual(out.items.map((i) => i.status), ['would-change', 'would-change']);
    assert.equal(fake.writes().length, 0);
  });
});

test('move without --to exits 64', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stderr } = await runCli(['move', 'MSG1'], { env });
    assert.equal(code, 64);
    assert.match(stderr, /--to/);
    assert.equal(fake.requests.length, 0);
  });
});

test('archive --mailbox resolves the folder and moves inside the shared mailbox', async () => {
  const shared = mailbox();
  shared.messages = [{ Id: 'SHARED1', Subject: 'shared', ParentFolderId: 'F-INBOX' }];
  await withFake({ me: mailbox(), [SHARED]: shared }, async (fake, env) => {
    const { code, stdout } = await runCli(['--mailbox', SHARED, 'archive', 'SHARED1'], { env });
    assert.equal(code, 0);
    assert.equal(JSON.parse(stdout).mailbox, SHARED);
    assert.ok(fake.requests.every((r) => r.mailbox === SHARED && r.anchor === `SMTP:${SHARED}`));
    assert.equal(fake.boxes[SHARED].messages[0].ParentFolderId, 'F-ARCHIVE');
    assert.equal(fake.boxes.me.messages[0].ParentFolderId, 'F-INBOX');
  });
});

test('folders --parent lists the subfolders of a resolved folder', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['folders', '--parent', 'Inbox'], { env });
    assert.equal(code, 0);
    assert.deepEqual(JSON.parse(stdout).value.map((f) => f.DisplayName), ['Projects']);
    const listing = fake.requests.at(-1);
    assert.equal(listing.path, '/mailfolders/F-INBOX/childfolders');
    assert.equal(listing.query.$top, '100');
  });
});

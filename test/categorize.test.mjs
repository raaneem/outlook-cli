// categorize / categories: master-list validation and canonical spelling,
// add / remove / set / clear semantics, and the CLI wiring against the
// in-memory fake mailbox.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startFakeMailbox } from './helpers.mjs';
import {
  parseCategoryChange,
  resolveCategoryNames,
  applyCategoryChange,
  sameCategories,
} from '../src/categories.mjs';
import { AppError, E } from '../src/errors.mjs';

const SHARED = 'shared@example.com';
const MASTER = [
  { DisplayName: 'Done', Color: 'Preset4' },
  { DisplayName: 'For information', Color: 'Preset7' },
  { DisplayName: 'In progress', Color: 'Preset1' },
];
const isArgsError = (e) => e instanceof AppError && e.code === E.ARGS;

// ---------------------------------------------------------------------------
// Unit

test('parseCategoryChange accepts add/remove together, or set, or clear', () => {
  assert.deepEqual(parseCategoryChange({ add: [' Done '], remove: ['x'] }), {
    add: ['Done'],
    remove: ['x'],
    set: null,
    clear: false,
  });
  assert.deepEqual(parseCategoryChange({ set: ['A', 'a', 'B'] }).set, ['A', 'B']);
  assert.equal(parseCategoryChange({ clear: true }).clear, true);
});

test('parseCategoryChange refuses mixes, empty names, and no-ops', () => {
  assert.throws(() => parseCategoryChange({ set: ['A'], add: ['B'] }), isArgsError);
  assert.throws(() => parseCategoryChange({ set: ['A'], clear: true }), isArgsError);
  assert.throws(() => parseCategoryChange({ add: ['  '] }), isArgsError);
  assert.throws(() => parseCategoryChange({}), (e) => isArgsError(e) && /Nothing to change/.test(e.message));
});

test('resolveCategoryNames uses the master spelling and refuses unknown names', () => {
  const change = parseCategoryChange({ add: ['for INFORMATION'] });
  assert.deepEqual(resolveCategoryNames(change, MASTER, 'x').add, ['For information']);
  assert.throws(
    () => resolveCategoryNames(parseCategoryChange({ add: ['Urgent'] }), MASTER, 'x'),
    (e) => isArgsError(e) && /Known categories: Done, For information, In progress/.test(e.hint),
  );
  // --allow-unlisted: no master list, names pass through as typed.
  assert.deepEqual(resolveCategoryNames(parseCategoryChange({ add: ['Urgent'] }), null, 'x').add, ['Urgent']);
});

test('applyCategoryChange adds, removes, replaces casing variants, sets and clears', () => {
  const add = { add: ['In progress'], remove: [], set: null, clear: false };
  assert.deepEqual(applyCategoryChange(['Done'], add), ['Done', 'In progress']);
  assert.deepEqual(applyCategoryChange(['in Progress'], add), ['In progress']);
  const remove = { add: [], remove: ['done'], set: null, clear: false };
  assert.deepEqual(applyCategoryChange(['Done', 'Stray'], remove), ['Stray']);
  assert.deepEqual(applyCategoryChange(['Done'], { add: [], remove: [], set: ['A'], clear: false }), ['A']);
  assert.deepEqual(applyCategoryChange(['Done'], { add: [], remove: [], set: null, clear: true }), []);
});

test('sameCategories ignores order but not spelling', () => {
  assert.equal(sameCategories(['A', 'B'], ['B', 'A']), true);
  assert.equal(sameCategories(['a'], ['A']), false);
  assert.equal(sameCategories(['A'], ['A', 'B']), false);
});

// ---------------------------------------------------------------------------
// CLI

function mailbox() {
  return {
    folders: [{ Id: 'F-INBOX', DisplayName: 'Inbox', WellKnownName: 'Inbox' }],
    categories: MASTER,
    messages: [
      { Id: 'MSG1', Subject: 'one', ParentFolderId: 'F-INBOX', Categories: ['Done'] },
      { Id: 'MSG2', Subject: 'two', ParentFolderId: 'F-INBOX', Categories: ['for information'] },
      { Id: 'MSG3', Subject: 'three', ParentFolderId: 'F-INBOX', Categories: [] },
    ],
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

test('categorize --add writes the master spelling and keeps existing categories', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout, stderr } = await runCli(['categorize', 'MSG1', 'MSG2', 'MSG3', '--add', 'FOR information'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.deepEqual(out.change.add, ['For information']);
    assert.deepEqual(out.items.map((i) => [i.id, i.status]), [
      ['MSG1', 'changed'],
      ['MSG2', 'changed'], // casing variant fixed in place
      ['MSG3', 'changed'],
    ]);
    assert.deepEqual(fake.writes().map((w) => [w.path, w.body.Categories]), [
      ['/messages/MSG1', ['Done', 'For information']],
      ['/messages/MSG2', ['For information']],
      ['/messages/MSG3', ['For information']],
    ]);
  });
});

test('categorize refuses an unknown category before reading any message (exit 64)', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stderr } = await runCli(['categorize', 'MSG1', '--add', 'Urgent'], { env });
    assert.equal(code, 64);
    assert.match(stderr, /Unknown category in your mailbox: Urgent/);
    assert.match(stderr, /Known categories: Done, For information, In progress/);
    assert.deepEqual(fake.requests.map((r) => r.path), ['/MasterCategories']);
  });
});

test('categorize --allow-unlisted skips the master list and writes the name as typed', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code } = await runCli(['categorize', 'MSG3', '--add', 'Urgent', '--allow-unlisted'], { env });
    assert.equal(code, 0);
    assert.ok(!fake.requests.some((r) => r.path === '/MasterCategories'));
    assert.deepEqual(fake.writes()[0].body, { Categories: ['Urgent'] });
  });
});

test('categorize --remove matches case-insensitively without needing the master list', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['categorize', 'MSG1', 'MSG2', '--remove', 'FOR INFORMATION'], { env });
    assert.equal(code, 0);
    assert.deepEqual(JSON.parse(stdout).items.map((i) => i.status), ['skipped', 'changed']);
    assert.ok(!fake.requests.some((r) => r.path === '/MasterCategories'));
    assert.deepEqual(fake.writes().map((w) => w.body), [{ Categories: [] }]);
  });
});

test('categorize --set replaces, --clear empties, and no-ops are skipped', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const set = await runCli(['categorize', 'MSG1', '--set', 'In progress', '--set', 'done'], { env });
    assert.equal(set.code, 0);
    assert.deepEqual(fake.writes()[0].body, { Categories: ['In progress', 'Done'] });

    const again = await runCli(['categorize', 'MSG1', '--set', 'Done', '--set', 'In progress'], { env });
    assert.equal(JSON.parse(again.stdout).items[0].reason, 'no-change');

    const clear = await runCli(['categorize', 'MSG1', '--clear'], { env });
    assert.equal(clear.code, 0);
    assert.deepEqual(fake.writes().at(-1).body, { Categories: [] });
    assert.equal(fake.writes().length, 2);
  });
});

test('categorize with a bad option mix exits 64 before any request', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const mix = await runCli(['categorize', 'MSG1', '--set', 'Done', '--add', 'In progress'], { env });
    const none = await runCli(['categorize', 'MSG1'], { env });
    assert.equal(mix.code, 64);
    assert.equal(none.code, 64);
    assert.match(none.stderr, /Nothing to change/);
    assert.equal(fake.requests.length, 0);
  });
});

test('an unreadable master list exits 3 and suggests --allow-unlisted', async () => {
  const faults = { 'GET /MasterCategories': { status: 403, body: { error: { code: 'ErrorAccessDenied', message: 'denied' } } } };
  await withFake(
    { me: mailbox() },
    async (fake, env) => {
      const { code, stderr } = await runCli(['categorize', 'MSG1', '--add', 'Done'], { env });
      assert.equal(code, 3);
      assert.match(stderr, /Could not read the master category list .*ErrorAccessDenied/);
      assert.match(stderr, /--allow-unlisted/);
      assert.equal(fake.writes().length, 0);
    },
    { faults },
  );
});

test('categorize --dry-run plans without writing', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['categorize', 'MSG3', '--add', 'Done', '--dry-run'], { env });
    assert.equal(code, 0);
    const item = JSON.parse(stdout).items[0];
    assert.equal(item.status, 'would-change');
    assert.deepEqual(item.after, { categories: ['Done'] });
    assert.equal(fake.writes().length, 0);
  });
});

test('categories and categorize --mailbox use the shared mailbox master list', async () => {
  const shared = mailbox();
  shared.categories = [{ DisplayName: 'Team only', Color: 'Preset2' }];
  await withFake({ me: mailbox(), [SHARED]: shared }, async (fake, env) => {
    const list = await runCli(['--mailbox', SHARED, 'categories'], { env });
    assert.equal(list.code, 0);
    assert.deepEqual(JSON.parse(list.stdout).value.map((c) => c.DisplayName), ['Team only']);

    const ok = await runCli(['--mailbox', SHARED, 'categorize', 'MSG3', '--add', 'team ONLY'], { env });
    assert.equal(ok.code, 0);
    assert.deepEqual(fake.boxes[SHARED].messages[2].Categories, ['Team only']);
    assert.deepEqual(fake.boxes.me.messages[2].Categories, []);

    // A name from the user's own master list is not valid in the shared mailbox.
    const wrong = await runCli(['--mailbox', SHARED, 'categorize', 'MSG3', '--add', 'Done'], { env });
    assert.equal(wrong.code, 64);
    assert.match(wrong.stderr, /Unknown category in shared@example\.com: Done/);
  });
});

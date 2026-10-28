// send-draft: the two-step preview / confirm flow, every refusal path, and
// the no-double-send guarantees (no 503 retry, no resend after a dropped
// connection). Driven through the CLI against the in-memory fake mailbox,
// which -- unlike Exchange -- would send anything: every refusal here is
// the CLI's own.

import { test } from 'node:test';
import { strict as assert } from 'node:assert';

import { runCli, seedTokenCache, startFakeMailbox } from './helpers.mjs';
import { draftPreview, assertSendable } from '../src/send-draft.mjs';
import { call } from '../src/client.mjs';
import { AppError, E } from '../src/errors.mjs';

const SHARED = 'shared@example.com';
const to = (address, name = null) => ({ EmailAddress: { Name: name, Address: address } });

function mailbox() {
  return {
    folders: [
      { Id: 'F-INBOX', DisplayName: 'Inbox', WellKnownName: 'Inbox' },
      { Id: 'F-DRAFTS', DisplayName: 'Drafts', WellKnownName: 'Drafts' },
      { Id: 'F-SENT', DisplayName: 'Sent Items', WellKnownName: 'SentItems' },
      { Id: 'F-OTHER', DisplayName: 'Other' },
    ],
    messages: [
      {
        Id: 'DRAFT1',
        Subject: 'Re: plan',
        ParentFolderId: 'F-DRAFTS',
        IsDraft: true,
        ChangeKey: 'CK-A',
        ToRecipients: [to('alice@example.com', 'Alice')],
        HasAttachments: true,
        BodyPreview: 'Sounds good',
        Attachments: [{ Name: 'plan.pdf', Size: 1234, ContentType: 'application/pdf', IsInline: false }],
      },
      { Id: 'INBOX1', Subject: 'plan', ParentFolderId: 'F-INBOX', IsDraft: false, ChangeKey: 'CK-I', ToRecipients: [to('me@example.com')] },
      { Id: 'EMPTY1', Subject: 'nobody', ParentFolderId: 'F-DRAFTS', IsDraft: true, ChangeKey: 'CK-E', ToRecipients: [] },
      { Id: 'ASIDE1', Subject: 'filed', ParentFolderId: 'F-OTHER', IsDraft: true, ChangeKey: 'CK-O', ToRecipients: [to('a@example.com')] },
      {
        Id: 'WIDE1',
        Subject: 'everyone',
        ParentFolderId: 'F-DRAFTS',
        IsDraft: true,
        ChangeKey: 'CK-W',
        ToRecipients: [to('a@example.com'), to('b@example.com')],
        CcRecipients: [to('c@example.com')],
        BccRecipients: [to('d@example.com')],
      },
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

const sends = (fake) => fake.writes().filter((w) => w.path.endsWith('/send'));

// ---------------------------------------------------------------------------
// Unit

test('draftPreview lists every recipient and attachment', () => {
  const p = draftPreview(mailbox().messages[4], [{ Name: 'a.txt', Size: 3 }]);
  assert.equal(p.recipientCount, 4);
  assert.deepEqual(p.bcc, [{ name: null, address: 'd@example.com' }]);
  assert.deepEqual(p.attachments, [{ name: 'a.txt', size: 3, contentType: null, isInline: false }]);
});

test('assertSendable refuses with E_PRECONDITION, never sends a non-draft', () => {
  const ok = { draftsFolderId: 'F-DRAFTS', where: 'x' };
  const isPrecondition = (e) => e instanceof AppError && e.code === E.PRECONDITION;
  const [draft, received] = mailbox().messages;
  assert.doesNotThrow(() => assertSendable(draft, { ...ok, expectChangeKey: 'CK-A' }));
  assert.throws(() => assertSendable(received, ok), isPrecondition);
  assert.throws(() => assertSendable(draft, { ...ok, expectChangeKey: 'CK-OLD' }), isPrecondition);
  assert.throws(() => assertSendable(draft, { ...ok, draftsFolderId: undefined }), isPrecondition);
});

test('the brakes in call() hold for any method casing', async () => {
  // Helper scripts import call() directly and may pass lower-case verbs;
  // fetch accepts them, so the brakes must too. Refusal happens before any
  // request, so no server is needed.
  const isPrecondition = (e) => e instanceof AppError && e.code === E.PRECONDITION;
  const keys = ['OUTLOOK_NO_SEND', 'OUTLOOK_READ_ONLY'];
  const saved = Object.fromEntries(keys.map((k) => [k, process.env[k]]));
  try {
    process.env.OUTLOOK_NO_SEND = '1';
    await assert.rejects(call({}, '/messages/D1/send', { method: 'post' }), isPrecondition);
    await assert.rejects(call({}, '/sendMail', { method: 'PUT' }), isPrecondition);
    delete process.env.OUTLOOK_NO_SEND;
    process.env.OUTLOOK_READ_ONLY = 'true';
    await assert.rejects(call({}, '/messages/D1', { method: 'patch' }), isPrecondition);
  } finally {
    for (const k of keys) {
      if (saved[k] === undefined) delete process.env[k];
      else process.env[k] = saved[k];
    }
  }
});

// ---------------------------------------------------------------------------
// CLI

test('send-draft --dry-run previews recipients, attachments and changeKey, and sends nothing', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout, stderr } = await runCli(['send-draft', 'DRAFT1', '--dry-run'], { env });
    assert.equal(code, 0, stderr);
    const out = JSON.parse(stdout);
    assert.equal(out.mode, 'dry-run');
    assert.equal(out.sent, false);
    assert.equal(out.draft.changeKey, 'CK-A');
    assert.deepEqual(out.draft.to, [{ name: 'Alice', address: 'alice@example.com' }]);
    assert.equal(out.draft.recipientCount, 1);
    assert.equal(out.draft.attachments[0].name, 'plan.pdf');
    assert.equal(fake.writes().length, 0);
  });
});

test('send-draft with the previewed changeKey sends exactly once', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stdout } = await runCli(['send-draft', 'DRAFT1', '--expect-change-key', 'CK-A'], { env });
    assert.equal(code, 0);
    const out = JSON.parse(stdout);
    assert.equal(out.sent, true);
    assert.equal(out.mode, 'apply');
    assert.deepEqual(fake.writes().map((w) => [w.method, w.path]), [['POST', '/messages/DRAFT1/send']]);
    assert.equal(fake.boxes.me.messages[0].ParentFolderId, 'F-SENT');
  });
});

test('send-draft without --expect-change-key or --dry-run exits 64 before any request', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stderr } = await runCli(['send-draft', 'DRAFT1'], { env });
    assert.equal(code, 64);
    assert.match(stderr, /--expect-change-key/);
    assert.equal(fake.requests.length, 0);
  });
});

test('a draft edited after its preview is refused (exit 4) and not sent', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const preview = JSON.parse((await runCli(['send-draft', 'DRAFT1', '--dry-run'], { env })).stdout);
    fake.boxes.me.messages[0].ChangeKey = 'CK-B'; // someone edits it in OWA
    const { code, stderr } = await runCli(
      ['send-draft', 'DRAFT1', '--expect-change-key', preview.draft.changeKey],
      { env },
    );
    assert.equal(code, 4);
    assert.match(stderr, /E_PRECONDITION/);
    assert.match(stderr, /changed since it was previewed/);
    assert.equal(sends(fake).length, 0);

    // A dry run with a key verifies it without sending.
    const check = await runCli(['send-draft', 'DRAFT1', '--dry-run', '--expect-change-key', 'CK-A'], { env });
    assert.equal(check.code, 4);
  });
});

test('received messages, drafts outside Drafts, and drafts without recipients are refused', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const cases = [
      ['INBOX1', 'CK-I', /Not a draft/],
      ['ASIDE1', 'CK-O', /not in the Drafts folder/],
      ['EMPTY1', 'CK-E', /no recipients/],
    ];
    for (const [id, key, reason] of cases) {
      const { code, stderr } = await runCli(['send-draft', id, '--expect-change-key', key], { env });
      assert.equal(code, 4, id);
      assert.match(stderr, reason);
    }
    assert.equal(sends(fake).length, 0);
  });
});

test('--max-recipients refuses a draft addressed to more people', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const { code, stderr } = await runCli(
      ['send-draft', 'WIDE1', '--expect-change-key', 'CK-W', '--max-recipients', '3'],
      { env },
    );
    assert.equal(code, 4);
    assert.match(stderr, /has 4 recipients; --max-recipients is 3/);
    assert.equal(sends(fake).length, 0);
  });
});

test('a missing draft, or one from another mailbox, exits 4 with the right hint', async () => {
  const shared = mailbox();
  shared.messages = [{ ...mailbox().messages[0], Id: 'SHAREDDRAFT1' }];
  await withFake({ me: mailbox(), [SHARED]: shared }, async (fake, env) => {
    const missing = await runCli(['send-draft', 'NOPE1', '--expect-change-key', 'x'], { env });
    assert.equal(missing.code, 4);
    assert.match(missing.stderr, /Draft not found in your mailbox/);
    assert.match(missing.stderr, /Check Sent Items/);

    const foreign = await runCli(['send-draft', 'SHAREDDRAFT1', '--expect-change-key', 'CK-A'], { env });
    assert.equal(foreign.code, 4);
    assert.match(foreign.stderr, /different mailbox/);
    assert.equal(sends(fake).length, 0);
  });
});

test('send-draft --mailbox sends from the shared mailbox Drafts', async () => {
  const shared = mailbox();
  shared.messages = [{ ...mailbox().messages[0], Id: 'SHAREDDRAFT1' }];
  await withFake({ me: mailbox(), [SHARED]: shared }, async (fake, env) => {
    const { code, stdout } = await runCli(
      ['--mailbox', SHARED, 'send-draft', 'SHAREDDRAFT1', '--expect-change-key', 'CK-A'],
      { env },
    );
    assert.equal(code, 0);
    assert.equal(JSON.parse(stdout).mailbox, SHARED);
    assert.ok(fake.requests.every((r) => r.mailbox === SHARED && r.anchor === `SMTP:${SHARED}`));
    assert.equal(fake.boxes[SHARED].messages[0].ParentFolderId, 'F-SENT');
  });
});

test('a 503 on send is not retried; a 429 is', async () => {
  const unavailable = { 'POST /messages/DRAFT1/send': { status: 503, body: { error: { code: 'ServerBusy', message: 'busy' } } } };
  await withFake(
    { me: mailbox() },
    async (fake, env) => {
      const { code } = await runCli(['send-draft', 'DRAFT1', '--expect-change-key', 'CK-A'], { env });
      assert.equal(code, 3);
      assert.equal(sends(fake).length, 1, 'a 503 send may have been processed; never repeat it');
    },
    { faults: unavailable },
  );

  const throttled = {
    'POST /messages/DRAFT1/send': { status: 429, headers: { 'Retry-After': '0' }, body: { error: { code: 'TooManyRequests', message: 'slow down' } } },
  };
  await withFake(
    { me: mailbox() },
    async (fake, env) => {
      const { code } = await runCli(['send-draft', 'DRAFT1', '--expect-change-key', 'CK-A'], { env });
      assert.equal(code, 3);
      assert.equal(sends(fake).length, 3, '429 means not processed, so retrying is safe');
    },
    { faults: throttled },
  );
});

test('a dropped connection reports outcome unknown and never resends', async () => {
  await withFake(
    { me: mailbox() },
    async (fake, env) => {
      const { code, stdout, stderr } = await runCli(['send-draft', 'DRAFT1', '--expect-change-key', 'CK-A'], { env });
      assert.equal(code, 3);
      assert.match(stderr, /E_NETWORK/);
      assert.match(stderr, /Check Sent Items/);
      const out = JSON.parse(stdout);
      assert.equal(out.ok, false);
      assert.equal(out.sent, null);
      assert.equal(out.outcome, 'unknown');
      assert.equal(out.draftStillInDrafts, true);
      assert.equal(sends(fake).length, 1);
    },
    { faults: { 'POST /messages/DRAFT1/send': 'drop' } },
  );
});

test('OUTLOOK_NO_SEND refuses send-draft (exit 4) but still allows the preview', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const noSend = { ...env, OUTLOOK_NO_SEND: '1' };
    const preview = await runCli(['send-draft', 'DRAFT1', '--dry-run'], { env: noSend });
    assert.equal(preview.code, 0);
    const { code, stderr } = await runCli(['send-draft', 'DRAFT1', '--expect-change-key', 'CK-A'], { env: noSend });
    assert.equal(code, 4);
    assert.match(stderr, /OUTLOOK_NO_SEND/);
    assert.equal(sends(fake).length, 0);
  });
});

test('OUTLOOK_NO_SEND also blocks plain send but not organising', async () => {
  await withFake({ me: mailbox() }, async (fake, env) => {
    const noSend = { ...env, OUTLOOK_NO_SEND: '1' };
    const message = JSON.stringify({ Subject: 'x', ToRecipients: [to('a@example.com')] });
    const send = await runCli(['send', message], { env: noSend });
    assert.equal(send.code, 4);
    const organise = await runCli(['mark-read', 'INBOX1'], { env: noSend });
    assert.equal(organise.code, 0);
  });
});

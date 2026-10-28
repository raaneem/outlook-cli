// Checks and preview for `send-draft`.
//
// Sending cannot be undone, so the draft is re-read right before sending
// and must still be exactly what was approved: an unsent draft (never a
// received message), sitting in Drafts, with at least one recipient, and
// carrying the ChangeKey the --dry-run preview printed. ChangeKey is
// Exchange's version stamp for an item; it changes on every edit,
// including one made in OWA after the preview was approved.

import { AppError, E } from './errors.mjs';

export const DRAFT_SELECT = [
  'Id',
  'ChangeKey',
  'IsDraft',
  'ParentFolderId',
  'Subject',
  'From',
  'ToRecipients',
  'CcRecipients',
  'BccRecipients',
  'HasAttachments',
  'BodyPreview',
  'LastModifiedDateTime',
  'WebLink',
].join(',');

function people(list) {
  return (list ?? []).map((r) => ({
    name: r?.EmailAddress?.Name ?? null,
    address: r?.EmailAddress?.Address ?? null,
  }));
}

/** What the user approves: everything that decides where the mail goes. */
export function draftPreview(draft, attachments = []) {
  const to = people(draft.ToRecipients);
  const cc = people(draft.CcRecipients);
  const bcc = people(draft.BccRecipients);
  return {
    id: draft.Id,
    changeKey: draft.ChangeKey ?? null,
    subject: draft.Subject ?? '',
    from: draft.From ? people([draft.From])[0] : null,
    to,
    cc,
    bcc,
    recipientCount: to.length + cc.length + bcc.length,
    attachments: attachments.map((a) => ({
      name: a.Name ?? null,
      size: a.Size ?? null,
      contentType: a.ContentType ?? null,
      isInline: a.IsInline === true,
    })),
    bodyPreview: draft.BodyPreview ?? '',
    lastModified: draft.LastModifiedDateTime ?? null,
    webLink: draft.WebLink ?? null,
  };
}

function refuse(message, hint) {
  return new AppError({ code: E.PRECONDITION, message, hint });
}

/**
 * Throw E_PRECONDITION unless `draft` may be sent as previewed. An
 * undefined `expectChangeKey` skips the version check (a plain dry run).
 */
export function assertSendable(draft, { draftsFolderId, expectChangeKey, maxRecipients = null, where }) {
  if (draft.IsDraft !== true) {
    throw refuse(
      `Not a draft: ${draft.Id}`,
      'send-draft only sends unsent drafts; received and already-sent messages are refused.',
    );
  }
  if (!draftsFolderId || draft.ParentFolderId !== draftsFolderId) {
    throw refuse(
      `Draft is not in the Drafts folder of ${where}.`,
      'Only drafts in Drafts are sent. Move it back in Outlook, or send it from there.',
    );
  }
  const { recipientCount } = draftPreview(draft);
  if (recipientCount === 0) {
    throw refuse('Draft has no recipients.', 'Add recipients in Outlook before sending.');
  }
  if (maxRecipients !== null && recipientCount > maxRecipients) {
    throw refuse(
      `Draft has ${recipientCount} recipients; --max-recipients is ${maxRecipients}.`,
      'Check To/Cc/Bcc in the --dry-run preview. Raise --max-recipients only if that many is intended.',
    );
  }
  if (expectChangeKey !== undefined && draft.ChangeKey !== expectChangeKey) {
    throw refuse(
      'Draft changed since it was previewed (ChangeKey mismatch); not sending.',
      'Run `send-draft <id> --dry-run` again, show the new preview to the user, and pass its changeKey.',
    );
  }
}

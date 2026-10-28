// Shared / delegated mailbox targeting (`--mailbox`).
//
// Outlook REST addresses another mailbox as /users/{smtp} instead of /me,
// using the same Bearer token: the OWA token already carries the *.Shared
// scopes, and Exchange checks the signed-in user's own permissions on that
// mailbox (Full Access, Send As) on every call. No new consent is involved.
//
// Two wire details matter:
//   - Exchange routes requests by `x-anchormailbox`. The captured value
//     anchors the signed-in user's mailbox, so it must be replaced for a
//     shared mailbox, not supplemented.
//   - Message ids are bound to the mailbox they live in. An id used against
//     the wrong mailbox fails with ErrorInvalidMailboxItemId rather than
//     touching anything (LEARNINGS.md section 16).

import { AppError, E } from './errors.mjs';
import { DEFAULT_RESOURCE, resourceBase } from './resources.mjs';

const ADDRESS_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/**
 * Validate a `--mailbox` value. Returns the lower-cased address, or null
 * when the flag was not given (act on the signed-in user's own mailbox).
 */
export function parseMailbox(value) {
  if (value === undefined || value === null) return null;
  const address = String(value).trim().toLowerCase();
  if (!ADDRESS_RE.test(address)) {
    throw new AppError({
      code: E.ARGS,
      message: `Invalid --mailbox: ${value}`,
      hint: 'Pass the SMTP address of the shared mailbox, e.g. --mailbox shared@example.com.',
    });
  }
  return address;
}

/**
 * API base for a shared mailbox: the Outlook base with /me swapped for
 * /users/{smtp}. Derived from resourceBase() so OUTLOOK_API_BASE redirects
 * shared-mailbox calls in tests exactly like /me calls.
 */
export function mailboxBase(mailbox) {
  const root = resourceBase(DEFAULT_RESOURCE).replace(/\/me\/?$/, '');
  return `${root}/users/${encodeURIComponent(mailbox)}`;
}

/** Captured headers re-anchored to `mailbox`. */
export function anchorHeaders(headers, mailbox) {
  const out = {};
  for (const [key, value] of Object.entries(headers)) {
    // Drop every spelling of the header: fetch would send both and join
    // them into a single comma-separated value Exchange cannot route.
    if (key.toLowerCase() !== 'x-anchormailbox') out[key] = value;
  }
  out['x-anchormailbox'] = `SMTP:${mailbox}`;
  return out;
}

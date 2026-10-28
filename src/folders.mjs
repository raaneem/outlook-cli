// Destination folders for `move` / `archive`, plus child-folder listing.
//
// A destination resolves, in order, as:
//   - a well-known alias (Inbox, Archive, Deleted, Junk, Sent, Drafts,
//     Outbox), via GET /mailfolders/{WellKnownName}. Never by display name:
//     names are localised, so a mailbox set to another language has no
//     folder called "Deleted Items";
//   - a folder Id, as printed by `outlook folders`;
//   - a display-name path like "Inbox/Projects/2026", walked one level at a
//     time and matched case-insensitively. A missing or ambiguous segment
//     is an error, never a guess.
//
// `api` is injected (the CLI passes runApi) so auth, --mailbox routing and
// error mapping stay in one place.

import { AppError, E } from './errors.mjs';
import { isMissingItem } from './actions.mjs';
import { wellKnownFolder, looksLikeFolderId } from './odata.mjs';

const PAGE_SIZE = 250;
// 5000 folders under one parent is far past any real mailbox; the cap only
// stops a misbehaving server from paging forever.
const MAX_PAGES = 20;
const FOLDER_SELECT = '$select=Id,DisplayName';

function describe(folder, wellKnownName = null) {
  return { id: folder.Id, displayName: folder.DisplayName ?? null, wellKnownName };
}

function notFound(spec, where, detail) {
  return new AppError({
    code: E.ARGS,
    message: `Folder not found in ${where}: ${spec}`,
    hint: `${detail} Run \`outlook folders\` (with the same --mailbox) to see what exists.`,
  });
}

/** Every child folder under `parentId` (the top level when null), all pages. */
export async function listChildFolders(api, parentId) {
  const base = parentId
    ? `/mailfolders/${encodeURIComponent(parentId)}/childfolders`
    : '/mailfolders';
  const all = [];
  for (let page = 0; page < MAX_PAGES; page++) {
    const body = await api(`${base}?$top=${PAGE_SIZE}&$skip=${page * PAGE_SIZE}&${FOLDER_SELECT}`);
    const value = body?.value ?? [];
    all.push(...value);
    if (value.length < PAGE_SIZE) break;
  }
  return all;
}

async function getWellKnown(api, name, spec, where) {
  try {
    return await api(`/mailfolders/${name}?${FOLDER_SELECT}`);
  } catch (err) {
    if (!isMissingItem(err)) throw err;
    throw notFound(spec, where, `This mailbox has no ${name} folder.`);
  }
}

async function walkPath(api, spec, where) {
  const segments = spec.split('/').map((s) => s.trim()).filter(Boolean);
  if (segments.length === 0) throw notFound(spec, where, 'The folder path is empty.');
  let current = null;
  for (const [index, segment] of segments.entries()) {
    // "Inbox/Projects": a well-known first segment anchors the path.
    const known = index === 0 && segments.length > 1 ? wellKnownFolder(segment) : null;
    if (known) {
      current = await getWellKnown(api, known, spec, where);
      continue;
    }
    const children = await listChildFolders(api, current?.Id ?? null);
    const matches = children.filter(
      (f) => (f.DisplayName ?? '').toLowerCase() === segment.toLowerCase(),
    );
    const parent = current ? `under "${current.DisplayName}"` : 'at the top level';
    if (matches.length === 0) {
      throw notFound(
        spec,
        where,
        `No folder named "${segment}" ${parent}. For a subfolder give its path, e.g. "Inbox/${segment}".`,
      );
    }
    if (matches.length > 1) {
      throw new AppError({
        code: E.ARGS,
        message: `Folder path is ambiguous in ${where}: ${spec}`,
        hint: `${matches.length} folders are named "${segment}" ${parent}: ${matches
          .map((f) => f.Id)
          .join(', ')}. Pass one of those Ids instead.`,
      });
    }
    current = matches[0];
  }
  return current;
}

/**
 * Resolve a user-supplied folder to `{ id, displayName, wellKnownName }`.
 * Unknown folders throw E_ARGS before anything is changed.
 */
export async function resolveDestination(api, spec, { mailbox = null } = {}) {
  const where = mailbox ?? 'your mailbox';
  const raw = String(spec ?? '').trim();
  if (!raw) {
    throw new AppError({ code: E.ARGS, message: 'No destination folder given.', hint: 'Pass --to <folder>.' });
  }

  const known = wellKnownFolder(raw);
  if (known) return describe(await getWellKnown(api, known, raw, where), known);

  if (looksLikeFolderId(raw)) {
    try {
      return describe(await api(`/mailfolders/${encodeURIComponent(raw)}?${FOLDER_SELECT}`));
    } catch (err) {
      if (!isMissingItem(err)) throw err;
      // A long display-name path can pass for an Id; only paths get a
      // second chance.
      if (!raw.includes('/')) throw notFound(raw, where, 'No folder has that Id.');
    }
  }
  return describe(await walkPath(api, raw, where));
}

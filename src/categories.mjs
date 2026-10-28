// Category changes for `categorize`.
//
// Exchange accepts any string as a category and stores unknown names as
// uncoloured categories, so a typo or a casing slip ("For Information" vs
// "For information") silently forks the categories a shared mailbox's team
// filters on. Names for --add / --set are therefore matched
// case-insensitively against the mailbox's master category list and written
// in its exact spelling; unknown names are refused unless --allow-unlisted.
// --remove matches the message's own categories instead, so stray names
// that were never in the master list can still be cleaned up.

import { AppError, E } from './errors.mjs';

const MAX_NAME_LENGTH = 255;

const lower = (s) => s.toLowerCase();

/** Drop case-insensitive duplicates, keeping the first spelling. */
function dedupe(names) {
  const seen = new Set();
  return names.filter((name) => {
    const key = lower(name);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function tidy(names, flag) {
  return names.map((raw) => {
    const name = String(raw).trim();
    if (!name || name.length > MAX_NAME_LENGTH) {
      throw new AppError({
        code: E.ARGS,
        message: `Invalid ${flag} category: ${JSON.stringify(String(raw).slice(0, 80))}`,
        hint: `Category names are 1-${MAX_NAME_LENGTH} characters.`,
      });
    }
    return name;
  });
}

/**
 * Validate the requested change. Exactly one style per call: --add/--remove
 * (combinable), --set (replace the whole list), or --clear.
 */
export function parseCategoryChange({ add = [], remove = [], set, clear = false } = {}) {
  const change = {
    add: dedupe(tidy(add, '--add')),
    remove: dedupe(tidy(remove, '--remove')),
    set: set ? dedupe(tidy(set, '--set')) : null,
    clear: !!clear,
  };
  const incremental = change.add.length > 0 || change.remove.length > 0;
  if ((change.set && change.clear) || ((change.set || change.clear) && incremental)) {
    throw new AppError({
      code: E.ARGS,
      message: 'Use --add/--remove, --set, or --clear -- not a mix.',
      hint: '--set replaces every category, --clear removes them all, --add/--remove adjust the existing list.',
    });
  }
  if (!incremental && !change.set && !change.clear) {
    throw new AppError({
      code: E.ARGS,
      message: 'Nothing to change.',
      hint: 'Pass --add <name>, --remove <name>, --set <name>, or --clear.',
    });
  }
  return change;
}

/** The mailbox's master categories, `[{ DisplayName, Color, ... }]`. */
export async function loadMasterCategories(api, where) {
  try {
    const body = await api('/MasterCategories');
    return Array.isArray(body?.value) ? body.value : [];
  } catch (err) {
    if (!(err instanceof AppError) || err.code !== E.HTTP) throw err;
    throw new AppError({
      code: E.HTTP,
      status: err.status,
      serverCode: err.serverCode,
      message: `Could not read the master category list of ${where} (HTTP ${err.status}${
        err.serverCode ? `, ${err.serverCode}` : ''
      }).`,
      hint: 'Pass --allow-unlisted to skip the check; names are then written exactly as typed.',
    });
  }
}

/**
 * Rewrite --add / --set names to the master list's exact spelling, refusing
 * any name not in it. A null `master` means --allow-unlisted: names pass
 * through as typed.
 */
export function resolveCategoryNames(change, master, where) {
  if (!master) return change;
  const byKey = new Map(
    master.filter((c) => typeof c?.DisplayName === 'string').map((c) => [lower(c.DisplayName), c.DisplayName]),
  );
  const canonical = (name) => {
    const hit = byKey.get(lower(name));
    if (hit) return hit;
    const known = [...byKey.values()];
    throw new AppError({
      code: E.ARGS,
      message: `Unknown category in ${where}: ${name}`,
      hint: known.length
        ? `Known categories: ${known.join(', ')}. Pass --allow-unlisted to use a name that is not in the list.`
        : 'This mailbox has no master categories. Pass --allow-unlisted to use a name anyway.',
    });
  };
  return {
    ...change,
    add: dedupe(change.add.map(canonical)),
    set: change.set ? dedupe(change.set.map(canonical)) : null,
  };
}

/** The category list after applying `change` to `current`. */
export function applyCategoryChange(current, change) {
  if (change.clear) return [];
  if (change.set) return [...change.set];
  const removing = new Set(change.remove.map(lower));
  const next = current.filter((c) => !removing.has(lower(c)));
  for (const name of change.add) {
    const at = next.findIndex((c) => lower(c) === lower(name));
    if (at === -1) next.push(name);
    else next[at] = name; // replaces a casing variant with the master spelling
  }
  return next;
}

/** Same categories, ignoring order. Spelling counts, so casing fixes are writes. */
export function sameCategories(a, b) {
  if (a.length !== b.length) return false;
  const left = [...a].sort();
  const right = [...b].sort();
  return left.every((name, i) => name === right[i]);
}

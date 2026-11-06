import { isBindableKey, isCommandChord, parseChord, platformChord, type ShortcutChord } from './chords';

/**
 * What a host sets for one shortcut id: chords that replace the defaults, `null` to disable it,
 * or a per-platform form where an absent platform keeps its default.
 * @experimental
 */
export type ShortcutBinding =
  | ShortcutChord
  | readonly ShortcutChord[]
  | null
  | {
      mac?: ShortcutChord | readonly ShortcutChord[] | null;
      other?: ShortcutChord | readonly ShortcutChord[] | null;
    };

/**
 * A host binding the editor changed or refused. `displaced-default` is informational: the host
 * chord `chord` of `id` took over the default of `displaced`.
 * @experimental
 */
export interface ShortcutIssue {
  code: 'unknown-id' | 'invalid-chord' | 'reserved-chord' | 'duplicate-chord' | 'displaced-default';
  id: string;
  chord?: string;
  /** The other shortcut id involved, for `duplicate-chord` and `displaced-default`. */
  other?: string;
}

export interface ShortcutDefault {
  id: string;
  /** Default chords on the current platform. */
  chords: readonly string[];
}

export interface ResolveShortcutsOptions {
  mac: boolean;
  /** Whether a host may bind `id` although it has no default. */
  isBindable(id: string): boolean;
  /** Whether the surface keeps `chord` for itself, so `id` cannot have it. */
  isReserved?(chord: string, id: string): boolean;
}

export interface ResolvedShortcuts {
  /** Effective canonical chords by id; an id without chords is absent or empty. */
  chords: ReadonlyMap<string, readonly string[]>;
  /** The id each effective canonical chord runs. */
  owners: ReadonlyMap<string, string>;
  issues: readonly ShortcutIssue[];
}

function hostChords(binding: ShortcutBinding, mac: boolean): readonly string[] | null | undefined {
  if (binding === null || typeof binding === 'string' || Array.isArray(binding)) {
    return typeof binding === 'string' ? [binding] : (binding as readonly string[] | null);
  }
  const value = (binding as Exclude<ShortcutBinding, string | readonly string[] | null>)[mac ? 'mac' : 'other'];
  return typeof value === 'string' ? [value] : (value as readonly string[] | null | undefined);
}

/**
 * Applies a host's map to the defaults. Each rule is deterministic: invalid and reserved chords
 * are dropped, a chord two host entries claim is dropped from both, and a host chord takes a
 * default chord from whichever id held it. An entry left with no chords keeps its default.
 */
export function resolveShortcuts(
  defaults: readonly ShortcutDefault[],
  map: Readonly<Record<string, ShortcutBinding | undefined>> | undefined,
  options: ResolveShortcutsOptions
): ResolvedShortcuts {
  const { mac, isBindable, isReserved } = options;
  const issues: ShortcutIssue[] = [];
  const known = new Map(defaults.map((entry) => [entry.id, entry]));
  const requested = new Map<string, string[]>();
  const disabled = new Set<string>();

  for (const [id, binding] of Object.entries(map ?? {})) {
    if (binding === undefined) continue;
    if (!known.has(id) && !isBindable(id)) {
      issues.push({ code: 'unknown-id', id });
      continue;
    }
    const chords = hostChords(binding, mac);
    if (chords === undefined) continue;
    if (chords === null || chords.length === 0) {
      disabled.add(id);
      continue;
    }
    const accepted: string[] = [];
    for (const chord of chords) {
      const canonical = typeof chord === 'string' ? platformChord(chord, mac) : null;
      if (!canonical || !isBindableKey(parseChord(canonical).key) || !isCommandChord(canonical)) {
        issues.push({ code: 'invalid-chord', id, chord: String(chord) });
      } else if (isReserved?.(canonical, id)) {
        issues.push({ code: 'reserved-chord', id, chord });
      } else if (!accepted.includes(canonical)) {
        accepted.push(canonical);
      }
    }
    if (accepted.length > 0) requested.set(id, accepted);
  }

  const claims = new Map<string, string[]>();
  for (const [id, chords] of requested) {
    for (const chord of chords) claims.set(chord, [...(claims.get(chord) ?? []), id]);
  }
  for (const [chord, ids] of claims) {
    if (ids.length < 2) continue;
    ids.forEach((id, index) => {
      issues.push({ code: 'duplicate-chord', id, chord, other: ids[(index + 1) % ids.length] });
      const left = requested.get(id)!.filter((held) => held !== chord);
      if (left.length > 0) requested.set(id, left);
      else requested.delete(id);
    });
    claims.delete(chord);
  }

  const chords = new Map<string, readonly string[]>();
  const owners = new Map<string, string>();
  for (const [id, held] of requested) {
    chords.set(id, Object.freeze(held));
    for (const chord of held) owners.set(chord, id);
  }
  for (const { id, chords: defaultChords } of defaults) {
    if (requested.has(id)) continue;
    if (disabled.has(id)) {
      chords.set(id, Object.freeze([]));
      continue;
    }
    const held: string[] = [];
    for (const chord of defaultChords) {
      const canonical = platformChord(chord, mac)!;
      const taker = owners.get(canonical);
      if (taker === undefined) {
        held.push(canonical);
      } else if (taker !== id) {
        issues.push({ code: 'displaced-default', id: taker, chord: canonical, other: id });
      }
    }
    chords.set(id, Object.freeze(held));
    for (const chord of held) owners.set(chord, id);
  }
  for (const id of disabled) if (!chords.has(id)) chords.set(id, Object.freeze([]));
  return { chords, owners, issues: Object.freeze(issues) };
}

import type { PluginGrant } from '../host-contracts/plugins';

const READ_ONLY: PluginGrant<string> = Object.freeze({ commands: Object.freeze([]) });

/**
 * The effective grant: unknown fields dropped, write-only permissions ignored without
 * `document: 'write'`, commands deduplicated. Frozen and structurally comparable.
 */
export function normalizeGrant<CommandId extends string>(
  grant: PluginGrant<CommandId> | null | undefined
): PluginGrant<CommandId> {
  if (!grant || typeof grant !== 'object') return READ_ONLY as PluginGrant<CommandId>;
  const commands = Array.isArray(grant.commands)
    ? [...new Set(grant.commands.filter((id): id is CommandId => typeof id === 'string'))].sort()
    : [];
  if (grant.document !== 'write') {
    return Object.freeze({ commands: Object.freeze(commands) }) as PluginGrant<CommandId>;
  }
  return Object.freeze({
    commands: Object.freeze(commands),
    document: 'write',
    ...(grant.editBatches === true ? { editBatches: true } : {}),
    ...(grant.untrackedHistory === true ? { untrackedHistory: true } : {}),
  }) as PluginGrant<CommandId>;
}

export function sameGrant(a: PluginGrant<string>, b: PluginGrant<string>): boolean {
  const left = a.commands ?? [];
  const right = b.commands ?? [];
  return (
    a.document === b.document &&
    a.editBatches === b.editBatches &&
    a.untrackedHistory === b.untrackedHistory &&
    left.length === right.length &&
    left.every((id, index) => id === right[index])
  );
}

export function grantsCommand(grant: PluginGrant<string>, id: string): boolean {
  return grant.commands?.includes(id) ?? false;
}

export function grantsWrite(grant: PluginGrant<string>): boolean {
  return grant.document === 'write';
}

/** Whether a batch with this history mode may be applied. */
export function grantsEditBatch(
  grant: PluginGrant<string>,
  history: 'separate' | 'none' | undefined
): boolean {
  return (
    grant.document === 'write' &&
    grant.editBatches === true &&
    (history !== 'none' || grant.untrackedHistory === true)
  );
}

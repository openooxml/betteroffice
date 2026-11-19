/** @internal Whether a checkpoint is due when a session is given no `checkpointPolicy`. */
export function defaultCheckpointPolicy(stats: { applied: number; bytes: number }): boolean {
  return stats.applied >= 128 || stats.bytes >= 1024 * 1024;
}

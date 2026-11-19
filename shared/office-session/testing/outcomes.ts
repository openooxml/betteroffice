import type { CommandOutcome } from '../commands/types';

/** @internal An outcome as a duplicate or a refusal reports it: without the effect recovery replays. */
export function withoutEffect(outcome: CommandOutcome): CommandOutcome {
  const copy = { ...outcome };
  delete copy.effect;
  return copy;
}

import type { PresentationCommandAck, PresentationCommandOpenResult, PresentationCommandSaveResult } from '@betteroffice/pptx/worker';

const MAX_RECOVERY_UPDATE_BYTES = 64 * 1024 * 1024;

type RecoveryBaseline = { kind: 'source'; throughSeq: 0; engine: string }
  | { kind: 'update'; update: Uint8Array; throughSeq: number };

export const peerlessRecoveryFactory = {
  create(source: Blob | Uint8Array, journalByteLimit?: number) {
    const sourceSize = source instanceof Uint8Array ? source.byteLength : source.size;
    let saved: { bytes?: Uint8Array; throughSeq: number; version: string } = { throughSeq: 0, version: '' };
    let baseline: RecoveryBaseline | undefined;
    let journal: { throughSeq: number; version: string; updates: Uint8Array[] }[] = [];
    let journalBytes = 0;
    let retainedEntries = 0;
    let rebaseFloor = 0;
    return {
      get saved() { return saved; },
      get baseline() { return baseline; },
      get journal() { return journal; },
      open(update: ArrayBuffer | undefined, throughSeq: number, version: string,
        kind: PresentationCommandOpenResult['recoveryBaseline'] = 'update', engine?: string) {
        saved = { ...saved, throughSeq, version };
        baseline = kind === 'source' ? !update && throughSeq === 0 && sourceSize > 0 && engine
          ? { kind: 'source', throughSeq: 0, engine } : undefined
          : update ? { kind: 'update', update: new Uint8Array(update), throughSeq } : undefined;
      },
      append(throughSeq: number, version: string, updates: NonNullable<PresentationCommandAck['updates']>) {
        if (throughSeq <= (baseline?.throughSeq ?? saved.throughSeq) || !updates.length) return;
        const retained = updates.map((update) => new Uint8Array(update.bytes).slice());
        journal.push({ throughSeq, version, updates: retained });
        retainedEntries += 1;
        journalBytes += retained.reduce((sum, update) => sum + update.byteLength, 0);
      },
      validate() {
        if (!baseline || journal.length !== retainedEntries) throw new Error('Recovery baseline is unavailable');
        let throughSeq = baseline.throughSeq;
        for (const entry of journal) {
          if (entry.throughSeq < throughSeq || entry.throughSeq <= baseline.throughSeq) throw new Error('Recovery baseline is unavailable');
          throughSeq = entry.throughSeq;
        }
        return baseline;
      },
      needsRebase() {
        return !baseline || journalBytes > Math.max(rebaseFloor, journalByteLimit ?? Math.max(16 * 1024 * 1024,
          (baseline.kind === 'source' ? sourceSize : baseline.update.byteLength) / 2));
      },
      checkpoint(result: PresentationCommandSaveResult) {
        if (result.throughSeq < saved.throughSeq) return;
        saved = { bytes: new Uint8Array(result.bytes), throughSeq: result.throughSeq, version: result.version };
        if (!result.recoveryUpdate || result.throughSeq < (baseline?.throughSeq ?? 0)) return;
        if (baseline?.kind === 'source' && result.recoveryUpdate.byteLength > MAX_RECOVERY_UPDATE_BYTES) {
          rebaseFloor = Math.max(rebaseFloor, journalBytes * 2);
          return;
        }
        baseline = { kind: 'update', update: new Uint8Array(result.recoveryUpdate), throughSeq: result.throughSeq };
        journal = journal.filter((entry) => entry.throughSeq > result.throughSeq);
        journalBytes = journal.reduce((sum, entry) => sum + entry.updates.reduce((sum, update) => sum + update.byteLength, 0), 0);
        retainedEntries = journal.length;
      },
      hasUncheckpointedUpdates() { return journal.some((entry) => entry.throughSeq > saved.throughSeq); },
    };
  },
};

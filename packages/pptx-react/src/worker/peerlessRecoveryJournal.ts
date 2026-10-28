import type { PresentationCommandAck, PresentationCommandSaveResult } from '@betteroffice/pptx/worker';

export const peerlessRecoveryFactory = {
  create(bytes: Uint8Array, journalByteLimit?: number) {
    let saved = { bytes, throughSeq: 0, version: '' };
    let baseline: { update: Uint8Array; throughSeq: number } | undefined;
    let journal: { throughSeq: number; version: string; updates: Uint8Array[] }[] = [];
    let journalBytes = 0;
    return {
      get saved() { return saved; },
      get baseline() { return baseline; },
      get journal() { return journal; },
      open(update: ArrayBuffer | undefined, throughSeq: number, version: string) {
        saved = { ...saved, throughSeq, version };
        baseline = update ? { update: new Uint8Array(update), throughSeq } : undefined;
      },
      append(throughSeq: number, version: string, updates: NonNullable<PresentationCommandAck['updates']>) {
        if (throughSeq <= (baseline?.throughSeq ?? saved.throughSeq) || !updates.length) return;
        const retained = updates.map((update) => new Uint8Array(update.bytes).slice());
        journal.push({ throughSeq, version, updates: retained });
        journalBytes += retained.reduce((sum, update) => sum + update.byteLength, 0);
      },
      needsRebase() {
        return !baseline || journalBytes > (journalByteLimit ?? Math.max(16 * 1024 * 1024, (baseline?.update.byteLength ?? 0) / 2));
      },
      checkpoint(result: PresentationCommandSaveResult) {
        if (result.throughSeq < saved.throughSeq) return;
        saved = { bytes: new Uint8Array(result.bytes), throughSeq: result.throughSeq, version: result.version };
        if (!result.recoveryUpdate || result.throughSeq < (baseline?.throughSeq ?? 0)) return;
        baseline = { update: new Uint8Array(result.recoveryUpdate), throughSeq: result.throughSeq };
        journal = journal.filter((entry) => entry.throughSeq > result.throughSeq);
        journalBytes = journal.reduce((sum, entry) => sum + entry.updates.reduce((sum, update) => sum + update.byteLength, 0), 0);
      },
      hasUncheckpointedUpdates() { return journal.some((entry) => entry.throughSeq > saved.throughSeq); },
    };
  },
};

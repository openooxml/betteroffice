import type {
  Checkpoint, CommandAck, CommandEnvelope, FrameEnvelope,
} from '../../../../shared/office-session/commands/types';
import type {
  PresentationCommandAck, PresentationCommandEditState,
  PresentationCommandEnvelope, PresentationCommandFrame, PresentationCommandIdentity,
  PresentationCommandIntent, PresentationCommandOpenResult, PresentationCommandSelected,
} from './commandProtocol';

type VersionAssociation = { version: string; clock: number };

export type PresentationSharedEnvelope = CommandEnvelope<PresentationCommandIntent['payload']>;

export type PresentationSharedAck = CommandAck<
  PresentationCommandSelected | null, PresentationCommandEditState, PresentationCommandFrame
> & {
  selection: PresentationCommandSelected | null;
  editState: PresentationCommandEditState;
  pptx: Pick<PresentationCommandAck, 'incarnation' | 'version' | 'outcomes' | 'dirtySlideIds' | 'activeSlideId' | 'historyThroughSeq' | 'slides' | 'size' | 'creation' | 'editResult' | 'fontId' | 'updates' | 'mutationResult' | 'change'> & {
    retainedBasis?: VersionAssociation;
  };
};

export type PresentationSharedMethods = {
  edit(envelope: PresentationSharedEnvelope): PresentationSharedAck;
  checkpoint(): Checkpoint;
};

export type PresentationSharedOpenResult = PresentationCommandOpenResult & {
  shared: PresentationSharedAck;
};

export function createPresentationCommandProjection() {
  let clock = 0;
  const versions = new Map<number, string>();

  function association(version: string): VersionAssociation | undefined {
    for (const [value, token] of versions) {
      if (token === version) return { version, clock: value };
    }
    return undefined;
  }

  return {
    get version() { return clock; },
    initialize(version: string) { versions.set(clock, version); },
    commit(version: string) { versions.set(++clock, version); },
    retain(tokens: Iterable<string>) {
      const retained = new Set(tokens);
      for (const [value, token] of versions) {
        if (value !== clock && !retained.has(token)) versions.delete(value);
      }
    },
    envelope(command: PresentationSharedEnvelope, identity: PresentationCommandIdentity): PresentationCommandEnvelope {
      let version = versions.get(command.basis.version);
      if (version === undefined) {
        version = 'unretained';
        while (association(version)) version += '\0';
      }
      return { ...identity, seq: command.seq,
        basis: { epoch: command.basis.epoch, version },
        ...(command.afterSeq === undefined ? {} : { afterSeq: command.afterSeq }),
        kind: command.kind, payload: command.payload } as PresentationCommandEnvelope;
    },
    ack(value: PresentationCommandAck, retainedBasis?: string): PresentationSharedAck {
      const frame: FrameEnvelope<PresentationCommandFrame> | undefined = value.frame ? {
        coversThroughSeq: value.frame.coversThroughSeq, revision: value.frame.epoch, payload: value.frame,
      } : undefined;
      return { session: value.session, incarnation: 0, epoch: value.epoch,
        throughSeq: value.throughSeq, version: clock, selection: value.selection, editState: value.editState,
        outcomes: value.outcomes.map(({ committed, ...outcome }) =>
          committed ? outcome : { ...outcome, committed: false as const }),
        ...(frame ? { frame } : {}),
        pptx: { incarnation: value.incarnation, version: value.version,
          outcomes: value.outcomes, dirtySlideIds: value.dirtySlideIds, activeSlideId: value.activeSlideId,
          historyThroughSeq: value.historyThroughSeq,
          ...(value.creation ? { creation: value.creation } : {}),
          ...(value.change ? { change: value.change } : {}),
          ...("mutationResult" in value ? { mutationResult: value.mutationResult } : {}),
          ...(value.updates ? { updates: value.updates } : {}),
          ...(value.editResult ? { editResult: value.editResult } : {}),
          ...(value.fontId !== undefined ? { fontId: value.fontId } : {}),
          ...(value.slides ? { slides: value.slides, size: value.size } : {}),
          ...(retainedBasis === undefined ? {} : { retainedBasis: association(retainedBasis) }) } };
    },
    dispose() { versions.clear(); },
  };
}

export function createPresentationSharedAdapter(initial: PresentationSharedAck) {
  const versions = new Map<string, number>();
  const basisVersions = new Set<string>();
  const frames = new Map<number, string>();
  let throughSeq = -1;

  function observe(ack: PresentationSharedAck): void {
    if (ack.throughSeq < throughSeq) return;
    throughSeq = ack.throughSeq;
    versions.set(ack.pptx.version, ack.version);
    basisVersions.add(ack.pptx.version);
    while (basisVersions.size > 256) basisVersions.delete(basisVersions.values().next().value!);
    if (ack.frame) frames.set(ack.frame.revision, ack.pptx.version);
    while (frames.size > 257) frames.delete(frames.keys().next().value!);
    const retained = new Set([...basisVersions, ...frames.values(), ack.pptx.version]);
    if (ack.pptx.retainedBasis) {
      retained.add(ack.pptx.retainedBasis.version);
      versions.set(ack.pptx.retainedBasis.version, ack.pptx.retainedBasis.clock);
    }
    for (const token of versions.keys()) {
      if (!retained.has(token)) versions.delete(token);
    }
  }

  observe(initial);
  return {
    envelope(command: PresentationCommandEnvelope): PresentationSharedEnvelope {
      return { session: command.session, incarnation: initial.incarnation, seq: command.seq,
        basis: { epoch: command.basis.epoch, version: versions.get(command.basis.version) ?? -1 },
        ...(command.afterSeq === undefined ? {} : { afterSeq: command.afterSeq }),
        kind: command.kind, payload: command.payload };
    },
    ack(value: PresentationSharedAck): PresentationCommandAck {
      observe(value);
      return { session: value.session, incarnation: value.pptx.incarnation, epoch: value.epoch,
        throughSeq: value.throughSeq, version: value.pptx.version, selection: value.selection,
        editState: value.editState, outcomes: value.pptx.outcomes, dirtySlideIds: value.pptx.dirtySlideIds,
        activeSlideId: value.pptx.activeSlideId,
        historyThroughSeq: value.pptx.historyThroughSeq,
        ...(value.pptx.creation ? { creation: value.pptx.creation } : {}),
        ...(value.pptx.change ? { change: value.pptx.change } : {}),
        ...("mutationResult" in value.pptx ? { mutationResult: value.pptx.mutationResult } : {}),
        ...(value.pptx.updates ? { updates: value.pptx.updates } : {}),
        ...(value.pptx.editResult ? { editResult: value.pptx.editResult } : {}),
        ...(value.pptx.fontId !== undefined ? { fontId: value.pptx.fontId } : {}),
        ...(value.pptx.slides ? { slides: value.pptx.slides, size: value.pptx.size } : {}),
        ...(value.frame ? { frame: value.frame.payload } : value.pptx.activeSlideId === null ? { frame: null } : {}) };
    },
  };
}

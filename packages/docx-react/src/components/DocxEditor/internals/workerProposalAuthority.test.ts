import { expect, mock, test } from 'bun:test';
import {
  proposalSetIdentity,
  type DocxContentControlsResult,
  type DocxExportResult,
  type DocxLayoutMap,
  type DocxPagedStructuredContent,
  type DocxProposalRequest,
  type DocxProposalResult,
  type DocxProposalSnapshot,
  type ResidentDocumentRead,
  type ResidentProposalReply,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { createProposalRegistry, type DocxProposalSession } from '@betteroffice/docx/yrs/proposals';
import type { WorkerOpenedDocument } from '../hooks/useDisplayList';
import {
  beginWorkerProposalHandover,
  failWorkerProposalAuthority,
  handedOverRequest,
  registerWorkerProposalAuthority,
  workerProposalAuthority,
} from './workerProposalAuthority';
import {
  deferWorkerOpenReplica,
  ensureWorkerOpenReplica,
  requestWorkerOpenReplica,
} from './workerOpenReplica';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function reply(version = 'worker-1', changedStories: string[] = []): ResidentProposalReply {
  const snapshot: DocxProposalSnapshot = { version, previewVersion: 0, proposals: [] };
  return {
    mirror: { version, proposals: { previewVersion: 0, entries: [] } },
    result: { ok: true, snapshot },
    changedStories,
    geometry: { version, previewVersion: 0, proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [] },
    updates: [],
    stateVector: new Uint8Array(),
  };
}

function harness(laidOut = () => Promise.resolve()) {
  let mainVersion = 'main-1';
  let mirror: ResidentProposalReply['mirror'] | null = null;
  let registry: ResidentProposalReply['mirror']['proposals'] = { previewVersion: 0, entries: [] };
  const proposalChange = mock(() => {});
  const session = {
    version: () => mirror?.version ?? mainVersion,
    getProposals: (): DocxProposalSnapshot => ({
      version: mirror?.version ?? mainVersion,
      previewVersion: (mirror?.proposals ?? registry).previewVersion,
      proposals: (mirror?.proposals ?? registry).entries.map((entry) => entry.record),
    }),
    mirrorWorkerDocument(next: typeof mirror) {
      const previous = mirror?.proposals ?? registry;
      if (!next && mirror) registry = mirror.proposals;
      mirror = next;
      if (JSON.stringify(previous) !== JSON.stringify(mirror?.proposals ?? registry)) proposalChange();
    },
  } as unknown as YrsSession;
  const events: string[] = [];
  const worker = {
    proposal: mock(async (op: Parameters<WorkerOpenedDocument['proposal']>[0]) => {
      events.push(op.kind);
      return reply();
    }),
    documentRead: mock(async (read: { kind: string }) => {
      events.push(read.kind);
      return { version: 'worker-1', value: { ok: true, version: 'worker-1', view: 'accepted', paragraphs: [] } };
    }),
    handOver: mock(async () => {
      events.push('handOver');
      return { state: Uint8Array.of(1), version: 'worker-2', proposals: reply().mirror.proposals };
    }),
  };
  const relayout = mock(() => {});
  const contentChanged = mock(() => {});
  const projectionChanged = mock((_stories: readonly string[]) => {});
  let current = true;
  const authority = registerWorkerProposalAuthority(
    session, worker as unknown as WorkerOpenedDocument,
    {
      relayout, current: () => current, laidOut, contentChanged, projectionChanged,
      adopted: () => {}, handedOver: () => {},
    }
  );
  deferWorkerOpenReplica(session, () => new Promise(() => {}), () => {}, () => {});
  return {
    session, worker, events, authority, proposalChange, relayout, contentChanged, projectionChanged,
    replace: () => { current = false; },
    mainVersion: (version: string) => { mainVersion = version; },
  };
}

const request: DocxProposalRequest = { expectVersion: 'worker-1', proposals: [] };
const unusedMain = async () => { throw new Error('unexpected main call'); };

test('worker mutation marks reach the peer before mirror listeners and queued saves', async () => {
  const h = harness();
  await h.authority.initialize();
  h.worker.proposal.mockResolvedValueOnce({ ...reply('worker-2', ['body']), projectionStories: ['hf:rId7'] });
  h.projectionChanged.mockImplementation((stories) => {
    expect(h.session.version()).toBe('worker-1~');
    expect(stories).toEqual(['hf:rId7']);
  });
  await h.authority.propose(request, unusedMain);
  expect(h.projectionChanged).toHaveBeenCalledWith(['hf:rId7']);
  h.projectionChanged.mockImplementation(() => {});
  await h.authority.save(async () => {
    expect(h.projectionChanged).toHaveBeenCalledWith(['hf:rId7']);
    return new ArrayBuffer(1);
  });
});

test('save runs after every proposal call already queued without opening the replica', async () => {
  const h = harness();
  await h.authority.initialize();
  const held = deferred<ResidentProposalReply>();
  h.worker.proposal.mockImplementationOnce(async () => {
    h.events.push('propose');
    return held.promise;
  });
  const proposal = h.authority.propose(request, unusedMain);
  const save = mock(async () => {
    h.events.push('save');
    return new ArrayBuffer(1);
  });
  const saving = h.authority.save(save);
  await Promise.resolve();
  expect(save).not.toHaveBeenCalled();
  held.resolve(reply());
  await proposal;
  expect(await saving).toBeInstanceOf(ArrayBuffer);
  expect(h.events).toEqual(['snapshot', 'propose', 'save', 'snapshot']);
});

test('a failed authority rejects queued saves without running them', async () => {
  const h = harness();
  const error = new Error('Worker stopped');
  failWorkerProposalAuthority(h.session, error);
  const save = mock(async () => new ArrayBuffer(1));
  expect(await h.authority.save(save).catch((failure) => failure)).toBe(error);
  expect(save).not.toHaveBeenCalled();
});

test('save remains queued after the peer has taken over proposal calls', async () => {
  const h = harness();
  await h.authority.initialize();
  const handover = await beginWorkerProposalHandover(h.session);
  handover!.complete();
  const bytes = new ArrayBuffer(1);
  expect(await h.authority.save(async () => {
    h.events.push('save');
    return bytes;
  })).toBe(bytes);
  expect(h.events).toEqual(['snapshot', 'handOver', 'save']);
});

function navigationReply(version = 'worker-1', position = 42): ResidentProposalReply {
  const snapshot = reply(version);
  const paragraph = { kind: 'session' as const, sessionId: 'session', story: 'body', paraId: 'p1' };
  snapshot.mirror.proposals.entries = [{
    key: 'jump', suggest: { author: 'Host', date: '2026-09-30T00:00:00Z' },
    record: { id: 'jump', state: 'proposed', paragraph, revisionIds: [], changed: false },
  }];
  snapshot.geometry.proposals = proposalSetIdentity({
    version, previewVersion: 0, proposals: snapshot.mirror.proposals.entries.map(({ record }) => record),
  });
  snapshot.geometry.targets.jump = { ok: true, ranges: [], paragraph: position };
  snapshot.geometry.navigationTargets = {
    jump: { loc: { story: 'body', paraId: 'p1', offset: 0 }, position },
  };
  return snapshot;
}

test('mirrored proposal navigation uses no worker reads and unknown targets use one', async () => {
  const h = harness();
  const snapshot = navigationReply();
  h.worker.proposal.mockResolvedValueOnce(snapshot);
  await h.authority.initialize();
  const main = mock(() => 'unsupported' as const);
  expect(await h.authority.navigationTarget('body', 'p1', main)).toEqual({
    version: 'worker-1', target: snapshot.geometry.navigationTargets!.jump,
  });
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  h.worker.documentRead.mockResolvedValueOnce({ version: 'worker-1', value: 'missing-target' } as never);
  expect(await h.authority.navigationTarget('body', 'other', main)).toEqual({
    version: 'worker-1', target: 'missing-target',
  });
  expect(h.worker.documentRead).toHaveBeenCalledTimes(1);
  expect(h.worker.documentRead.mock.calls[0]![0]).toEqual<{
    kind: 'navigationTarget'; story: string; paraId: string;
  }>({
    kind: 'navigationTarget', story: 'body', paraId: 'other',
  });
  expect(main).not.toHaveBeenCalled();
});

test('mirrored navigation waits for preceding mutations and reads their new geometry', async () => {
  const h = harness();
  h.worker.proposal.mockResolvedValueOnce(navigationReply());
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return pending.promise; });
  const mutation = h.authority.propose(request, unusedMain);
  await posted.promise;
  expect(h.session.version()).toBe('worker-1~');
  const main = mock(() => 'unsupported' as const);
  const navigation = h.authority.navigationTarget('body', 'p1', main);
  const next = navigationReply('worker-2', 99);
  pending.resolve(next);
  await mutation;
  expect(await navigation).toEqual({ version: 'worker-2', target: next.geometry.navigationTargets!.jump });
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(main).not.toHaveBeenCalled();
});

test('navigation ignores stale mirrors and retains exact mirrored target failures', async () => {
  for (const mismatch of ['version', 'previewVersion', 'proposals', 'navigationTargets'] as const) {
    const h = harness();
    const snapshot = navigationReply();
    if (mismatch === 'version') snapshot.geometry.version = 'older';
    if (mismatch === 'previewVersion') snapshot.geometry.previewVersion += 1;
    if (mismatch === 'proposals') snapshot.geometry.proposals = 'other';
    if (mismatch === 'navigationTargets') snapshot.geometry.navigationTargets = undefined;
    h.worker.proposal.mockResolvedValueOnce(snapshot);
    h.worker.documentRead.mockResolvedValueOnce({ version: 'worker-1', value: 'ambiguous-target' } as never);
    expect(await h.authority.navigationTarget('body', 'p1', () => 'unsupported')).toEqual({
      version: 'worker-1', target: 'ambiguous-target',
    });
    expect(h.worker.documentRead).toHaveBeenCalledTimes(1);
  }
  for (const target of ['missing-target', 'ambiguous-target', 'unsupported'] as const) {
    const h = harness();
    const snapshot = navigationReply();
    snapshot.geometry.navigationTargets!.jump = target;
    h.worker.proposal.mockResolvedValueOnce(snapshot);
    expect(await h.authority.navigationTarget('body', 'p1', () => 'unsupported')).toEqual({
      version: 'worker-1', target,
    });
    expect(h.worker.documentRead).not.toHaveBeenCalled();
  }
});

test('navigation before hand-over uses the mirror and navigation after waits for the main replica', async () => {
  const h = harness();
  const snapshot = navigationReply();
  h.worker.proposal.mockImplementationOnce(async (op) => { h.events.push(op.kind); return snapshot; });
  await h.authority.initialize();
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  const before = h.authority.navigationTarget('body', 'p1', () => 'unsupported');
  const ready = requestWorkerOpenReplica(h.session)!;
  const target = { loc: { story: 'body', paraId: 'p1', offset: 0 }, position: 42 };
  const main = mock(() => target);
  const after = h.authority.navigationTarget('body', 'p1', main);
  expect(await before).toEqual({ version: 'worker-1', target: snapshot.geometry.navigationTargets!.jump });
  await ready;
  expect(await after).toEqual({ version: 'main-2', target });
  expect(main).toHaveBeenCalledTimes(1);
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(h.events).toEqual(['snapshot', 'handOver']);
});

test('initialization runs once and serializes reads after an in-flight proposal', async () => {
  const h = harness();
  const first = h.authority.initialize();
  expect(h.authority.initialize()).toBe(first);
  await first;
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async (op) => {
    h.events.push(op.kind);
    posted.resolve();
    return pending.promise;
  });
  const propose = h.authority.propose(request, unusedMain);
  await posted.promise;
  const read = h.authority.paragraphIdentities(unusedMain);
  const snapshot = h.authority.getProposals(unusedMain);
  expect(h.events).toEqual(['snapshot', 'propose']);
  pending.resolve(reply('worker-2'));
  await propose;
  await read;
  expect((await snapshot).version).toBe('worker-2');
  expect(h.events).toEqual(['snapshot', 'propose', 'paragraphIdentities']);
});

test('routed reads wait for layout and the initialization snapshot before posting', async () => {
  const laidOut = deferred<void>();
  const h = harness(() => laidOut.promise);
  const snapshot = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async (op) => {
    h.events.push(op.kind);
    posted.resolve();
    return snapshot.promise;
  });
  const read = h.authority.readParagraphs({
    story: 'body', paraIds: ['p1'], view: 'accepted',
  }, unusedMain);
  await new Promise((done) => setTimeout(done, 0));
  expect(h.worker.proposal).not.toHaveBeenCalled();
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(h.events).toEqual([]);
  expect(h.authority.initialized).toBe(false);

  laidOut.resolve();
  await posted.promise;
  expect(h.events).toEqual(['snapshot']);
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  snapshot.resolve(reply());
  expect(await read).toEqual({ ok: true, version: 'worker-1', view: 'accepted', paragraphs: [] });
  expect(h.events).toEqual(['snapshot', 'readParagraphs']);
  expect(h.worker.proposal).toHaveBeenCalledTimes(1);
  expect(h.authority.initialized).toBe(true);
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('an in-flight document change masks the version without notifying proposal listeners', async () => {
  const h = harness();
  expect(h.authority.geometry()).toBeNull();
  const notify = mock(() => {});
  h.authority.subscribe(notify);
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return pending.promise; });
  const propose = h.authority.propose(request, unusedMain);
  await posted.promise;
  expect(h.session.version()).toBe('worker-1~');
  expect(h.proposalChange).not.toHaveBeenCalled();
  expect(notify).toHaveBeenCalledTimes(1);
  const applied = reply('worker-2', ['body']);
  applied.mirror.proposals.previewVersion = 1;
  pending.resolve(applied);
  expect(await propose).toEqual(applied.result!);
  expect(h.session.version()).toBe('worker-2');
  expect(h.authority.geometry()).toBe(applied.geometry);
  expect(h.proposalChange).toHaveBeenCalledTimes(1);
  expect(notify).toHaveBeenCalledTimes(2);
  expect(h.authority.holdsWorkerState()).toBe(true);
  expect(h.relayout).toHaveBeenCalledTimes(1);
});

test('hydration releases reads waiting for the first layout without posting initialization', async () => {
  const laidOut = deferred<void>();
  const waiting = deferred<void>();
  const h = harness(() => { waiting.resolve(); return laidOut.promise; });
  const main = mock(async () => ({ sessionId: 'session', packageSha256: null, paragraphs: [] }));
  const read = h.authority.paragraphIdentities(main);
  await waiting.promise;
  expect(h.worker.proposal).not.toHaveBeenCalled();
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {});

  await requestWorkerOpenReplica(h.session);
  expect(await read).toEqual({ sessionId: 'session', packageSha256: null, paragraphs: [] });
  expect(h.session.version()).toBe('main-2');
  expect(main).toHaveBeenCalledTimes(1);
  expect(h.worker.proposal).not.toHaveBeenCalled();
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(h.worker.handOver).toHaveBeenCalledTimes(1);
  expect(h.authority.initialized).toBe(false);
  expect(workerProposalAuthority(h.session)).toBeNull();
});

test('a session failure rejects reads waiting for the first layout', async () => {
  const waiting = deferred<void>();
  const laidOut = deferred<void>();
  const h = harness(() => { waiting.resolve(); return laidOut.promise; });
  const read = h.authority.paragraphIdentities(unusedMain);
  const rejected = read.then(() => null, (error: unknown) => error);
  await waiting.promise;
  const failure = new Error('session failed');
  failWorkerProposalAuthority(h.session, failure);
  expect(await rejected).toBe(failure);
  expect(h.worker.proposal).not.toHaveBeenCalled();
  expect(h.worker.documentRead).not.toHaveBeenCalled();
});

test('registry changes hold worker state without relayout and setStates never masks the version', async () => {
  const h = harness();
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return pending.promise; });
  const call = h.authority.setStates({ expectVersion: 'worker-1', expectPreviewVersion: 0, changes: [] }, unusedMain);
  await posted.promise;
  expect(h.session.version()).toBe('worker-1');
  const changed = reply();
  changed.mirror.proposals.previewVersion = 1;
  pending.resolve(changed);
  await call;
  expect(h.authority.holdsWorkerState()).toBe(true);
  expect(h.relayout).not.toHaveBeenCalled();
});

test('a failed initialized authority rejects running and queued calls and hand-over with the same failure', async () => {
  const h = harness();
  await h.authority.initialize();
  h.worker.proposal.mockResolvedValue(reply('worker-2', ['body']));
  await h.authority.propose(request, unusedMain);
  expect(h.authority.holdsWorkerState()).toBe(true);
  const pending = deferred<never>();
  const posted = deferred<void>();
  h.worker.documentRead.mockImplementation(async () => {
    posted.resolve();
    return pending.promise;
  });
  const failure = new Error('worker lost');
  const read = h.authority.paragraphIdentities(unusedMain);
  const rejectedRead = read.then(() => null, (error: unknown) => error);
  await posted.promise;
  const snapshot = h.authority.getProposals(unusedMain);
  const rejectedSnapshot = snapshot.then(() => null, (error: unknown) => error);
  const handover = beginWorkerProposalHandover(h.session)!;
  const rejectedHandover = handover.then(() => null, (error: unknown) => error);
  failWorkerProposalAuthority(h.session, failure);
  expect(await rejectedRead).toBe(failure);
  expect(await rejectedSnapshot).toBe(failure);
  expect(await rejectedHandover).toBe(failure);
  failWorkerProposalAuthority(h.session, new Error('later failure'));
  await expect(h.authority.initialize()).rejects.toBe(failure);
  await expect(h.authority.getProposals(unusedMain)).rejects.toBe(failure);
  await expect(beginWorkerProposalHandover(h.session)!).rejects.toBe(failure);
  expect(h.worker.handOver).not.toHaveBeenCalled();
  expect(h.worker.proposal).toHaveBeenCalledTimes(2);
});

test('a completed hand-over releases exclusive worker state and keeps routing reads to main', async () => {
  const h = harness();
  await h.authority.initialize();
  h.worker.proposal.mockResolvedValue(reply('worker-2', ['body']));
  await h.authority.propose(request, unusedMain);
  const handover = await beginWorkerProposalHandover(h.session)!;
  expect(h.authority.holdsWorkerState()).toBe(true);
  deferWorkerOpenReplica(h.session, async () => () => {
    h.mainVersion('main-2');
    handover.complete();
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  await requestWorkerOpenReplica(h.session);
  expect(h.authority.holdsWorkerState()).toBe(false);
  expect(workerProposalAuthority(h.session)).toBeNull();
  const main = mock(async () => h.session.getProposals());
  expect((await h.authority.getProposals(main)).version).toBe('main-2');
  handover.complete();
  expect(h.authority.holdsWorkerState()).toBe(false);
  expect(main).toHaveBeenCalledTimes(1);
  expect(h.worker.proposal).toHaveBeenCalledTimes(2);
  expect(h.worker.handOver).toHaveBeenCalledTimes(1);
});

test('comment deletion queues with proposals and reads and stores changed worker state', async () => {
  const h = harness();
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementationOnce(async (op) => {
    h.events.push(op.kind);
    posted.resolve();
    return pending.promise;
  });
  const main = mock(() => {});
  const deletion = h.authority.removeComment('7', main);
  await posted.promise;
  expect(h.worker.proposal.mock.calls.at(-1)![0]).toEqual({ kind: 'removeComment', id: '7' });
  expect(h.session.version()).toBe('worker-1~');
  expect(h.authority.holdsWorkerState()).toBe(true);
  expect(h.authority.holdsCommittedWorkerState()).toBe(false);
  const propose = h.authority.propose(request, unusedMain);
  const read = h.authority.readParagraphs({ view: 'accepted' }, unusedMain);
  expect(h.events).toEqual(['snapshot', 'removeComment']);

  const changed = reply('worker-2', ['body']);
  delete changed.result;
  pending.resolve(changed);
  expect(await deletion).toBeUndefined();
  expect(h.session.version()).toBe('worker-2');
  expect(h.authority.geometry()).toBe(changed.geometry);
  expect(h.authority.holdsCommittedWorkerState()).toBe(true);
  expect(h.relayout).toHaveBeenCalledTimes(1);
  expect(h.contentChanged).toHaveBeenCalledTimes(1);
  expect(main).not.toHaveBeenCalled();
  await propose;
  await read;
  expect(h.events).toEqual(['snapshot', 'removeComment', 'propose', 'readParagraphs']);
});

test('unchanged comment deletion does not hold worker state or relayout', async () => {
  const h = harness();
  await h.authority.initialize();
  const unchanged = reply();
  delete unchanged.result;
  h.worker.proposal.mockResolvedValueOnce(unchanged);
  await h.authority.removeComment('missing', unusedMain);
  expect(h.authority.holdsWorkerState()).toBe(false);
  expect(h.authority.geometry()).toBe(unchanged.geometry);
  expect(h.relayout).not.toHaveBeenCalled();
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('comment deletion queued after hand-over waits for and runs the main continuation', async () => {
  const h = harness();
  await h.authority.initialize();
  const handover = await beginWorkerProposalHandover(h.session)!;
  const replica = deferred<() => void>();
  deferWorkerOpenReplica(h.session, () => replica.promise, () => {
    throw new Error('unexpected fallback');
  }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const main = mock(async () => {});
  const deletion = h.authority.removeComment('7', main);
  expect(main).not.toHaveBeenCalled();
  replica.resolve(() => { h.mainVersion('main-2'); handover.complete(); });
  await ready;
  expect(await deletion).toBeUndefined();
  expect(main).toHaveBeenCalledTimes(1);
  expect(h.events).toEqual(['snapshot', 'handOver']);
  expect(h.worker.proposal).toHaveBeenCalledTimes(1);
  expect(h.relayout).not.toHaveBeenCalled();
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('propose and withdraw relayout only when stories change', async () => {
  const h = harness();
  await h.authority.propose(request, unusedMain);
  expect(h.relayout).not.toHaveBeenCalled();
  expect(h.authority.holdsWorkerState()).toBe(false);
  h.worker.proposal.mockResolvedValue(reply('worker-2', ['body']));
  await h.authority.withdraw({ expectVersion: 'worker-1', ids: [] }, unusedMain);
  expect(h.relayout).toHaveBeenCalledTimes(1);
});

for (const kind of ['propose', 'withdraw', 'setStates'] as const) {
  test(`${kind} notifies content changes once only when stories change`, async () => {
    const h = harness();
    await h.authority.initialize();
    const mutate = () => {
      switch (kind) {
        case 'propose':
          return h.authority.propose(request, unusedMain);
        case 'withdraw':
          return h.authority.withdraw({ expectVersion: 'worker-1', ids: [] }, unusedMain);
        case 'setStates':
          return h.authority.setStates({
            expectVersion: 'worker-1', expectPreviewVersion: 0, changes: [],
          }, unusedMain);
      }
    };
    expect(h.contentChanged).not.toHaveBeenCalled();
    await mutate();
    expect(h.contentChanged).not.toHaveBeenCalled();

    const changed = reply('worker-2', ['body', 'hf:header']);
    h.worker.proposal.mockResolvedValue(changed);
    expect(await mutate()).toEqual(changed.result!);
    expect(h.contentChanged).toHaveBeenCalledTimes(1);

    const unchanged = reply('worker-2');
    h.worker.proposal.mockResolvedValue(unchanged);
    expect(await mutate()).toEqual(unchanged.result!);
    expect(h.contentChanged).toHaveBeenCalledTimes(1);

    const refused = reply('worker-2');
    refused.result = {
      ok: false,
      version: 'worker-2',
      failure: { code: 'stale-version', message: 'The document changed' },
    };
    h.worker.proposal.mockResolvedValue(refused);
    expect(await mutate()).toEqual(refused.result);
    expect(h.contentChanged).toHaveBeenCalledTimes(1);

    h.worker.proposal.mockRejectedValue(new Error('worker unavailable'));
    await expect(mutate()).rejects.toThrow('worker unavailable');
    expect(h.contentChanged).toHaveBeenCalledTimes(1);
  });
}

test('an initialization snapshot never notifies content changes', async () => {
  const h = harness();
  h.worker.proposal.mockResolvedValue(reply('worker-1', ['body']));
  await h.authority.initialize();
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('main continuations complete in FIFO order after hand-over begins', async () => {
  const h = harness();
  await h.authority.initialize();
  const handingOver = beginWorkerProposalHandover(h.session)!;
  const replica = deferred<() => void>();
  deferWorkerOpenReplica(h.session, () => replica.promise, () => {
    throw new Error('unexpected fallback');
  }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const pending = deferred<DocxProposalResult>();
  const started = deferred<void>();
  const order: string[] = [];
  const main = mock(async () => {
    started.resolve();
    return pending.promise;
  });
  const mutation = h.authority.propose(request, main).then((result) => {
    order.push('mutation');
    return result;
  });
  const readMain = mock(async () => h.session.getProposals());
  const read = h.authority.getProposals(readMain).then((result) => {
    order.push('read');
    return result;
  });
  const handover = await handingOver;
  expect(main).not.toHaveBeenCalled();
  expect(readMain).not.toHaveBeenCalled();
  replica.resolve(() => { h.mainVersion('main-2'); handover.complete(); });
  await ready;
  await started.promise;
  await new Promise((done) => setTimeout(done, 0));
  expect(order).toEqual([]);
  expect(main).toHaveBeenCalledTimes(1);
  expect(readMain).not.toHaveBeenCalled();

  const result = reply('main-2').result!;
  pending.resolve(result);
  expect(await mutation).toEqual(result);
  expect((await read).version).toBe('main-2');
  expect(order).toEqual(['mutation', 'read']);
  expect(readMain).toHaveBeenCalledTimes(1);
  expect(h.worker.proposal).toHaveBeenCalledTimes(1);
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('hand-over waits for the running worker call and routes queued calls to main with a version rewrite', async () => {
  const h = harness();
  await h.authority.initialize();
  const notify = mock(() => {});
  h.authority.subscribe(notify);
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return pending.promise; });
  const inFlight = h.authority.propose(request, unusedMain);
  await posted.promise;
  const main = mock(async (input: DocxProposalRequest) => {
    expect(handedOverRequest(h.session, input).expectVersion).toBe('main-2');
    return reply('main-2').result!;
  });
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const queued = h.authority.propose({ ...request, expectVersion: 'worker-2' }, main);
  const later = h.authority.getProposals(async () => h.session.getProposals());
  expect(h.worker.handOver).not.toHaveBeenCalled();
  pending.resolve(reply('worker-2', ['body']));
  await inFlight;
  await ready;
  await queued;
  expect((await later).version).toBe('main-2');
  expect(h.worker.proposal).toHaveBeenCalledTimes(2);
  expect(h.worker.handOver).toHaveBeenCalledTimes(1);
  expect(h.authority.geometry()).toBeNull();
  expect(notify).toHaveBeenCalledTimes(2);
  expect(workerProposalAuthority(h.session)).toBeNull();
  const staleMain = mock(async (input: DocxProposalRequest) => {
    expect(handedOverRequest(h.session, input).expectVersion).toBe('worker-2');
    return reply('main-3').result!;
  });
  h.mainVersion('main-3');
  await h.authority.propose({ ...request, expectVersion: 'worker-2' }, staleMain);
});

test('hydration callbacks read after queued main withdrawals finish', async () => {
  const h = harness();
  await h.authority.initialize();
  const applied = reply('worker-2', ['body']);
  applied.mirror.proposals.entries.push({
    record: {
      id: 'held', state: 'proposed', changed: true, revisionIds: [],
      paragraph: { kind: 'session', sessionId: 'session', story: 'body', paraId: 'first' },
    },
    key: 'held', suggest: { author: 'Host', date: '2026-09-29T00:00:00Z' },
  });
  applied.result = {
    ok: true,
    snapshot: {
      version: 'worker-2', previewVersion: 0,
      proposals: applied.mirror.proposals.entries.map((entry) => entry.record),
    },
  };
  applied.geometry.proposals = proposalSetIdentity(applied.result.snapshot);
  h.worker.handOver.mockResolvedValue({
    state: Uint8Array.of(1), version: 'worker-2', proposals: applied.mirror.proposals,
  });
  const proposed = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return proposed.promise; });
  const proposal = h.authority.propose(request, unusedMain);
  await posted.promise;
  const withdrawn = deferred<void>();
  const started = deferred<void>();
  const order: string[] = [];
  let read!: Promise<DocxProposalSnapshot>;
  const readMain = mock(async () => { order.push('read'); return h.session.getProposals(); });
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {
    const authority = workerProposalAuthority(h.session);
    expect(authority).toBe(h.authority);
    read = authority ? authority.getProposals(readMain) : readMain();
  });
  const ready = requestWorkerOpenReplica(h.session)!;
  const withdrawal = h.authority.withdraw({ expectVersion: 'worker-2', ids: ['held'] }, async () => {
    started.resolve();
    await withdrawn.promise;
    h.mainVersion('main-3');
    h.session.mirrorWorkerDocument(reply('main-3').mirror);
    h.session.mirrorWorkerDocument(null);
    order.push('withdraw');
    return reply('main-3').result!;
  });
  proposed.resolve(applied);
  await proposal;
  await ready;
  await started.promise;
  expect(h.session.getProposals().proposals.map((proposal) => proposal.id)).toEqual(['held']);
  expect(readMain).not.toHaveBeenCalled();
  expect(workerProposalAuthority(h.session)).toBe(h.authority);
  withdrawn.resolve();
  await withdrawal;
  expect(await read).toEqual({ version: 'main-3', previewVersion: 0, proposals: [] });
  expect(order).toEqual(['withdraw', 'read']);
  expect(workerProposalAuthority(h.session)).toBeNull();
});

test('routing stops when the replica starts without worker state or finishes', () => {
  const h = harness();
  expect(workerProposalAuthority(h.session)).toBe(h.authority);
  requestWorkerOpenReplica(h.session);
  expect(workerProposalAuthority(h.session)).toBeNull();
  const finished = harness();
  ensureWorkerOpenReplica(finished.session);
  expect(workerProposalAuthority(finished.session)).toBeNull();
});

test('worker state keeps routing after hydration starts', async () => {
  const h = harness();
  await h.authority.initialize();
  h.worker.proposal.mockResolvedValue(reply('worker-2', ['body']));
  await h.authority.propose(request, unusedMain);
  requestWorkerOpenReplica(h.session);
  expect(workerProposalAuthority(h.session)).toBe(h.authority);
});

test('a replaced document rejects the worker result', async () => {
  const h = harness();
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return pending.promise; });
  const call = h.authority.propose(request, unusedMain);
  await posted.promise;
  h.replace();
  pending.resolve(reply('worker-2', ['body']));
  await expect(call).rejects.toThrow('The document changed while applying proposals');
  expect(h.relayout).not.toHaveBeenCalled();
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('failed initialization rejects and is never posted twice', async () => {
  const h = harness();
  h.worker.proposal.mockRejectedValue(new Error('worker unavailable'));
  await expect(h.authority.initialize()).rejects.toThrow('worker unavailable');
  await expect(h.authority.getProposals(unusedMain)).rejects.toThrow('worker unavailable');
  expect(h.worker.proposal).toHaveBeenCalledTimes(1);
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('anchor and navigation reads retain the worker version and input order', async () => {
  const h = harness();
  const anchors = [
    { kind: 'session', sessionId: 'session', story: 'body', paraId: 'second' },
    { kind: 'session', sessionId: 'session', story: 'body', paraId: 'first' },
  ] as const;
  const results = anchors.map((anchor) => ({ status: 'found' as const, anchor }));
  h.worker.documentRead.mockImplementation(async (read) => {
    h.events.push(read.kind);
    if (read.kind === 'resolveParagraphAnchors') {
      return { version: 'worker-3', value: { results } } as never;
    }
    return { version: 'worker-3', value: 'missing-target' } as never;
  });
  expect(await h.authority.resolveParagraphAnchors(anchors, unusedMain)).toEqual({
    version: 'worker-3', results,
  });
  expect(h.worker.documentRead.mock.calls[0]![0]).toEqual<{
    kind: 'resolveParagraphAnchors';
    anchors: typeof anchors;
  }>({ kind: 'resolveParagraphAnchors', anchors });
  expect(await h.authority.navigationTarget('body', 'missing', () => 'unsupported')).toEqual({
    version: 'worker-3', target: 'missing-target',
  });
  expect(h.events).toEqual(['snapshot', 'resolveParagraphAnchors', 'navigationTarget']);
});

test('search reads retain the worker version and queue behind other document reads', async () => {
  const h = harness();
  const waiting = deferred<void>();
  const posted = deferred<void>();
  const carry = { story: 'body', encoded: Uint8Array.of(1) };
  const value = { matches: [{
    story: 'body', paraId: 'p1', start: 2, displayFrom: 3, displayTo: 7,
  }], carried: 0 };
  h.worker.documentRead.mockImplementation(async (read) => {
    h.events.push(read.kind);
    if (read.kind === 'readParagraphs') {
      posted.resolve();
      await waiting.promise;
    }
    return { version: 'worker-3', value } as never;
  });
  const first = h.authority.readParagraphs({ view: 'accepted' }, unusedMain);
  await posted.promise;
  const search = h.authority.searchText('term', true, carry, () => { throw new Error('unexpected main call'); });
  const navigation = h.authority.navigationTarget('body', 'missing', () => 'unsupported');
  expect(h.events).toEqual(['snapshot', 'readParagraphs']);
  waiting.resolve();
  await first;
  expect(await search).toEqual({ version: 'worker-3', value });
  await navigation;
  expect(h.worker.documentRead.mock.calls[1]![0]).toEqual<{
    kind: 'searchText'; query: string; caseSensitive: boolean; carry: typeof carry;
  }>({ kind: 'searchText', query: 'term', caseSensitive: true, carry });
  expect(h.events).toEqual(['snapshot', 'readParagraphs', 'searchText', 'navigationTarget']);
});

test('sticky anchor batches keep their place and return their version even when it differs from the requested version', async () => {
  const h = harness();
  await h.authority.initialize();
  const locs = [
    { story: 'body', paraId: 'p1', offset: 2 },
    { story: 'body', paraId: 'p2', offset: 0 },
  ];
  const anchor = { story: 'body', encoded: Uint8Array.of(2) };
  const value = [anchor, null];
  const versions: string[] = [];
  h.worker.documentRead.mockImplementation(async (read) => {
    h.events.push(read.kind);
    versions.push(h.session.version());
    return { version: h.session.version(), value } as never;
  });
  h.worker.proposal.mockImplementation(async (op) => {
    h.events.push(op.kind);
    return reply('worker-2', ['body']);
  });
  const main = mock(() => { throw new Error('unexpected main call'); });
  const before = h.authority.stickyAnchors(locs, 'worker-1', main);
  const mutation = h.authority.propose(request, unusedMain);
  const after = h.authority.stickyAnchors(locs, 'worker-1', main);
  expect(await before).toEqual({ version: 'worker-1', value });
  await mutation;
  expect(await after).toEqual({ version: 'worker-2', value });
  expect(h.worker.documentRead.mock.calls.map(([read]) => read)).toEqual<Array<{
    kind: string; locs: typeof locs; version: string;
  }>>([
    { kind: 'stickyAnchors', locs, version: 'worker-1' },
    { kind: 'stickyAnchors', locs, version: 'worker-1' },
  ]);
  expect(versions).toEqual(['worker-1', 'worker-2']);
  expect(h.events).toEqual(['snapshot', 'stickyAnchors', 'propose', 'stickyAnchors']);
  expect(main).not.toHaveBeenCalled();
});

test('sticky anchors queued after hand-over use the current-state main fallback', async () => {
  const h = harness();
  await h.authority.initialize();
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const locs = [
    { story: 'body', paraId: 'p1', offset: 2 },
    { story: 'body', paraId: 'p2', offset: 0 },
  ];
  const anchor = { story: 'body', encoded: Uint8Array.of(2) };
  const value = [anchor, null];
  const encode = mock(() => value);
  const first = h.authority.stickyAnchors(locs, 'worker-1', encode);
  const second = h.authority.stickyAnchors(locs, 'worker-1', encode);
  await ready;
  expect(await first).toEqual({ version: 'main-2', value });
  expect(await second).toEqual({ version: 'main-2', value });
  expect(encode).toHaveBeenCalledTimes(2);
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(h.events).toEqual(['snapshot', 'handOver']);
});


test('a rejected worker call restores the visible version and leaves the queue usable', async () => {
  const h = harness();
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async () => { posted.resolve(); return pending.promise; });
  const call = h.authority.withdraw({ expectVersion: 'worker-1', ids: [] }, unusedMain);
  await posted.promise;
  expect(h.session.version()).toBe('worker-1~');
  pending.reject(new Error('worker unavailable'));
  await expect(call).rejects.toThrow('worker unavailable');
  expect((await h.authority.getProposals(unusedMain)).version).toBe('worker-1');
  expect(h.proposalChange).not.toHaveBeenCalled();
  expect(h.contentChanged).not.toHaveBeenCalled();
});

test('a call made before the hand-over began runs in the worker ahead of it', async () => {
  const h = harness();
  await h.authority.initialize();
  const pending = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementation(async (op) => {
    h.events.push(op.kind);
    posted.resolve();
    return pending.promise;
  });
  const inFlight = h.authority.propose(request, unusedMain);
  await posted.promise;
  h.worker.proposal.mockImplementation(async (op) => {
    h.events.push(op.kind);
    return reply('worker-3', ['body']);
  });
  const queued = h.authority.withdraw({ expectVersion: 'worker-2', ids: [] }, unusedMain);
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => {
      h.mainVersion('main-3');
      handover.complete();
    };
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const after = h.authority.getProposals(async () => h.session.getProposals());
  pending.resolve(reply('worker-2', ['body']));
  await inFlight;
  expect((await queued).ok).toBe(true);
  await ready;
  expect((await after).version).toBe('main-3');
  expect(h.events).toEqual(['snapshot', 'propose', 'withdraw', 'handOver']);
});

test('paged exports send the current request and parse the worker result', async () => {
  const h = harness();
  const options = { revisionView: 'markup', expectLayoutVersion: 'layout-1' } as const;
  const currentRequest = JSON.stringify({ renderEnv: {} });
  const result = {
    ok: true, version: 'worker-1',
    content: { structured: {}, layout: { documentVersion: 'worker-1', layoutVersion: 'layout-1', pages: [] } },
  } as unknown as DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>;
  h.worker.documentRead.mockImplementation(async (read) => {
    h.events.push(read.kind);
    return { version: 'worker-1', value: JSON.stringify(result) } as never;
  });
  expect(await h.authority.exportStructuredWithPages(options, async () => currentRequest, unusedMain)).toEqual(result);
  expect(h.worker.documentRead).toHaveBeenCalledWith({ kind: 'exportStructuredWithPages', options, currentRequest });
  expect(h.events).toEqual(['snapshot', 'exportStructuredWithPages']);
});

test('content-control reads route options to the worker and use main after hand-over', async () => {
  const h = harness();
  const options = { stories: ['body'], maxControls: 1 } as const;
  const query = { kind: 'tag', tag: 'field' } as const;
  const result: DocxContentControlsResult = {
    ok: true, version: 'worker-1',
    content: { schemaVersion: 1, anchorScope: 'session', includedStories: ['body'], controls: [], complete: true, diagnostics: [] },
  };
  h.worker.documentRead.mockResolvedValue({ version: 'worker-1', value: result } as never);
  expect(await h.authority.listContentControls(options, unusedMain)).toEqual(result);
  expect(await h.authority.findContentControls(query, options, unusedMain)).toEqual(result);
  expect(await h.authority.listContentControls(undefined, unusedMain)).toEqual(result);
  expect(await h.authority.findContentControls(query, undefined, unusedMain)).toEqual(result);
  expect(h.worker.documentRead.mock.calls.map(([read]) => read)).toEqual<ResidentDocumentRead[]>([
    { kind: 'listContentControls', options },
    { kind: 'findContentControls', query, options },
    { kind: 'listContentControls', options: {} },
    { kind: 'findContentControls', query, options: {} },
  ]);
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const mainResult = { ...result, version: 'main-2' };
  const main = mock(async () => mainResult);
  const listed = h.authority.listContentControls(options, main);
  const found = h.authority.findContentControls(query, options, main);
  await ready;
  expect(await listed).toEqual(mainResult);
  expect(await found).toEqual(mainResult);
  expect(main).toHaveBeenCalledTimes(2);
  expect(h.worker.documentRead).toHaveBeenCalledTimes(4);
});

test('paged exports queued after hand-over run the main continuation', async () => {
  const h = harness();
  await h.authority.initialize();
  deferWorkerOpenReplica(h.session, async () => {
    const handover = await beginWorkerProposalHandover(h.session)!;
    return () => { h.mainVersion('main-2'); handover.complete(); };
  }, () => { throw new Error('unexpected fallback'); }, () => {});
  const ready = requestWorkerOpenReplica(h.session)!;
  const result = {
    ok: false, version: 'main-2',
    failure: { code: 'layout-unavailable', target: null, message: 'No layout is ready.' },
  } as const;
  const main = mock(async () => result);
  const exportResult = h.authority.exportStructuredWithPages({ revisionView: 'markup' }, async () => '{}', main);
  await ready;
  expect(await exportResult).toEqual(result);
  expect(main).toHaveBeenCalledTimes(1);
  expect(h.worker.documentRead).not.toHaveBeenCalled();
  expect(h.events).toEqual(['snapshot', 'handOver']);
});

test('paged exports read their layout request after the calls queued ahead of them', async () => {
  const h = harness();
  await h.authority.initialize();
  h.worker.documentRead.mockImplementation(async (read) => {
    h.events.push(read.kind);
    return { version: 'worker-1', value: JSON.stringify({ ok: true }) } as never;
  });
  const ahead = h.authority.getProposals(unusedMain);
  const read = h.authority.exportStructuredWithPages({ revisionView: 'markup' }, async () => {
    h.events.push('request');
    return null;
  }, unusedMain);
  await ahead;
  expect(await read).toBeNull();
  expect(h.events.indexOf('request')).toBeGreaterThan(h.events.indexOf('snapshot'));
  expect(h.worker.documentRead).not.toHaveBeenCalled();
});

function editorRoundHarness(laidOut = async () => {}) {
  const local = { version: 'worker-1', previewVersion: 0, proposals: [] };
  const session = {
    version: () => 'worker-1', encodeStateVector: () => new Uint8Array(),
    getProposals: () => local,
    createWorkerProposalRegistry: (state: ResidentProposalReply['mirror']['proposals']) => {
      const registry = createProposalRegistry(session);
      registry.mirror({ version: session.version(), proposals: state });
      registry.mirror(null);
      return registry;
    },
    storiesChangedSince: () => ({ revision: 0, stories: [] }),
  } as unknown as YrsSession & DocxProposalSession;
  const events: string[] = [];
  const worker = {
    proposal: mock(async (op: Parameters<WorkerOpenedDocument['proposal']>[0]): Promise<ResidentProposalReply> => {
      events.push(op.kind);
      return { ...reply(), peerDiff: new Uint8Array() };
    }),
    documentRead: mock(async () => { throw new Error('unexpected document read'); }),
    handOver: mock(async () => { throw new Error('unexpected handover'); }),
    integrateProposalUpdate: mock((_update: Uint8Array, _stories: readonly string[]): readonly string[] => []),
  };
  const peerUpdated = mock((_stories: readonly string[]) => {});
  const authority = registerWorkerProposalAuthority(session, worker as unknown as WorkerOpenedDocument, {
    editorPeer: true, current: () => true, laidOut, relayout: () => {}, adopted: () => {},
    contentChanged: () => {}, peerUpdated,
  });
  return { session, worker, authority, peerUpdated, events };
}

test('integrated proposal updates notify peerUpdated only for changed stories', async () => {
  const h = editorRoundHarness();
  await h.authority.initialize();
  const states = { expectVersion: 'worker-1', expectPreviewVersion: 0, changes: [] };
  expect(await h.authority.setStates(states, unusedMain)).toMatchObject({ ok: true });
  expect(h.worker.integrateProposalUpdate).toHaveBeenCalledTimes(1);
  expect(h.peerUpdated).not.toHaveBeenCalled();
  h.worker.integrateProposalUpdate.mockReturnValueOnce(['body', 'hf:rId7']);
  h.worker.proposal.mockResolvedValueOnce({ ...reply('worker-1', ['body']), peerDiff: new Uint8Array() });
  expect(await h.authority.setStates(states, unusedMain)).toMatchObject({ ok: true });
  expect(h.peerUpdated).toHaveBeenCalledTimes(1);
  expect(h.peerUpdated).toHaveBeenCalledWith(['body', 'hf:rId7']);
});

test('editor save stays behind a round admitted before the first layout', async () => {
  const layout = deferred<void>();
  const waiting = deferred<void>();
  const h = editorRoundHarness(() => { waiting.resolve(); return layout.promise; });
  const round = h.authority.propose(request, unusedMain);
  const save = mock(async () => { h.events.push('save'); return 'saved'; });
  const saving = h.authority.save(save);
  await waiting.promise;
  expect(h.worker.proposal).not.toHaveBeenCalled();
  expect(save).not.toHaveBeenCalled();
  layout.resolve();
  expect(await round).toMatchObject({ ok: true });
  expect(await saving).toBe('saved');
  expect(h.events).toEqual(['snapshot', 'propose', 'save', 'snapshot']);
});

test('editor save stays behind a round admitted during initialization', async () => {
  const h = editorRoundHarness();
  const initializing = deferred<ResidentProposalReply>();
  const posted = deferred<void>();
  h.worker.proposal.mockImplementationOnce(async () => { h.events.push('snapshot'); posted.resolve(); return initializing.promise; });
  const initialization = h.authority.initialize();
  await posted.promise;
  const round = h.authority.propose(request, unusedMain);
  const saving = h.authority.save(async () => { h.events.push('save'); return 'saved'; });
  initializing.resolve(reply());
  await initialization;
  expect(await round).toMatchObject({ ok: true });
  expect(await saving).toBe('saved');
  expect(h.events).toEqual(['snapshot', 'propose', 'save', 'snapshot']);
});

test('a queued editor round reuses the failed initialization rejection', async () => {
  const h = editorRoundHarness();
  const failure = new Error('snapshot failed');
  h.worker.proposal.mockRejectedValueOnce(failure);
  await expect(h.authority.initialize()).rejects.toBe(failure);
  await expect(h.authority.propose(request, unusedMain)).rejects.toBe(failure);
  expect(h.worker.proposal).toHaveBeenCalledTimes(1);
});

test('an unchanged empty editor decision keeps the worker rebuildable', async () => {
  const h = editorRoundHarness();
  expect(await h.authority.setStates({ expectVersion: 'worker-1', expectPreviewVersion: 0, changes: [] }, unusedMain))
    .toMatchObject({ ok: true });
  expect(h.authority.holdsCommittedWorkerState()).toBe(false);
  h.authority.restart();
  expect(h.authority.initialized).toBe(false);
});

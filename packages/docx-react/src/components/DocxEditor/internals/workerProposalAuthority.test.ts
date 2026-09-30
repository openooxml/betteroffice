import { expect, mock, test } from 'bun:test';
import type {
  DocxProposalRequest,
  DocxProposalResult,
  DocxProposalSnapshot,
  ResidentProposalReply,
  YrsSession,
} from '@betteroffice/docx/yrs';
import type { WorkerOpenedDocument } from '../hooks/useDisplayList';
import {
  beginWorkerProposalHandover,
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
  return {
    mirror: { version, proposals: { previewVersion: 0, entries: [] } },
    result: { ok: true, snapshot: { version, previewVersion: 0, proposals: [] } },
    changedStories,
    geometry: { version, previewVersion: 0, proposals: '[]', targets: {}, hidden: [] },
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
      return { version: 'worker-1', value: { version: 'worker-1', paragraphs: [] } };
    }),
    handOver: mock(async () => {
      events.push('handOver');
      return { state: Uint8Array.of(1), version: 'worker-2', proposals: reply().mirror.proposals };
    }),
  };
  const relayout = mock(() => {});
  const contentChanged = mock(() => {});
  let current = true;
  const authority = registerWorkerProposalAuthority(
    session, worker as unknown as WorkerOpenedDocument,
    {
      relayout, current: () => current, laidOut, contentChanged,
      adopted: () => {}, handedOver: () => {},
    }
  );
  deferWorkerOpenReplica(session, () => new Promise(() => {}), () => {}, () => {});
  return {
    session, worker, events, authority, proposalChange, relayout, contentChanged,
    replace: () => { current = false; },
    mainVersion: (version: string) => { mainVersion = version; },
  };
}

const request: DocxProposalRequest = { expectVersion: 'worker-1', proposals: [] };
const unusedMain = async () => { throw new Error('unexpected main call'); };

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
  expect(await read).toEqual({ version: 'worker-1', paragraphs: [] });
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

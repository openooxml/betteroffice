import { expect, test } from 'bun:test';
import { createPresenter } from './presenter';
import type {
  CommandTicket, EchoEntry, EchoHooks, EchoOverrunSample, FrameEnvelope, PresentationStatus, Presenter,
} from './types';

function ticket(id: number, initialSeq?: number) {
  let seq = initialSeq;
  let acknowledge!: () => void;
  const value: CommandTicket = {
    id,
    get seq() { return seq; },
    ack: new Promise((resolve) => { acknowledge = () => resolve({ seq, status: 'applied' }); }),
    presented: Promise.resolve('presented'),
  };
  return { value, acknowledge, send(value: number) { seq = value; } };
}

function frame(coversThroughSeq: number, revision = coversThroughSeq, baseRevision?: number): FrameEnvelope<string> {
  return { coversThroughSeq, revision, baseRevision, payload: `frame:${revision}` };
}

function painters() {
  const paints: Array<readonly EchoEntry<string>[]> = [];
  const drops: Array<{ entries: readonly EchoEntry<string>[]; reason: string }> = [];
  const presents: Array<{
    frame: FrameEnvelope<string>;
    retired: readonly EchoEntry<string>[];
    remaining: readonly EchoEntry<string>[];
  }> = [];
  const echo: EchoHooks<string, string> = {
    paint: (pending) => { paints.push([...pending]); },
    drop: (entries, reason) => { drops.push({ entries: [...entries], reason }); },
    present: (frame, retired, remaining) => {
      presents.push({ frame, retired: [...retired], remaining: [...remaining] });
    },
  };
  return { echo, paints, drops, presents };
}

function bitmapFrame(coversThroughSeq: number, revision: number, baseRevision?: number) {
  const closes = [0, 0];
  const payload = closes.map((_, index) => ({ close() { closes[index]! += 1; } }) as ImageBitmap);
  const candidate: FrameEnvelope<ImageBitmap[]> = { coversThroughSeq, revision, baseRevision, payload };
  return { candidate, closes };
}

for (const transactionalPresentation of [false, true]) {
  for (const sameCoverage of [false, true]) {
    for (const action of ['reset', 'dispose'] as const) {
      test(`preserves painter ownership after ${transactionalPresentation ? 'transactional' : 'cosmetic'} ${sameCoverage ? 'same' : 'advancing'} coverage overrun ${action}`, () => {
        const first = bitmapFrame(0, 10);
        const next = bitmapFrame(sameCoverage ? 0 : 1, 11);
        let clock = 0;
        let overruns = 0;
        let installed: FrameEnvelope<ImageBitmap[]> | undefined;
        let presenter!: Presenter<string, ImageBitmap[]>;
        presenter = createPresenter<string, ImageBitmap[]>({
          transactionalPresentation, sameCoverageFrames: true,
          now: () => clock, requestFullFrame() {},
          discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
          echo: {
            paint() {}, drop() {},
            present(candidate) {
              installed?.payload.forEach((bitmap) => bitmap.close());
              installed = candidate;
              if (candidate.revision === 11) clock += 3;
            },
          },
          onEchoOverrun(sample) {
            expect(sample.hook).toBe('present');
            overruns += 1;
            presenter[action]();
          },
        });
        presenter.present(first.candidate);
        expect(presenter.present(next.candidate)).toBe(transactionalPresentation || action === 'dispose' ? 'stale' : 'presented');
        expect(overruns).toBe(1);
        expect(installed).toBe(next.candidate);
        expect(first.closes).toEqual([1, 1]);
        expect(next.closes).toEqual([0, 0]);
        installed!.payload.forEach((bitmap) => bitmap.close());
        expect(next.closes).toEqual([1, 1]);
      });
    }
  }
}

for (const rejection of ['older coverage', 'older revision', 'same coverage disabled', 'base mismatch', 'disposed'] as const) {
  test(`closes every bitmap once on ${rejection} and leaves accepted frames owned by the painter`, () => {
    const accepted = bitmapFrame(1, 10);
    const rejected = bitmapFrame(rejection === 'older coverage' ? 0 : rejection === 'base mismatch' ? 2 : 1,
      rejection === 'older revision' ? 9 : 11, rejection === 'base mismatch' ? 99 : undefined);
    let requests = 0;
    const presenter = createPresenter<string, ImageBitmap[]>({
      sameCoverageFrames: rejection !== 'same coverage disabled', now: () => 0,
      requestFullFrame() { requests += 1; },
      discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
    });
    expect(presenter.present(accepted.candidate)).toBe('presented');
    if (rejection === 'disposed') presenter.dispose();
    expect(presenter.present(rejected.candidate)).toBe(rejection === 'base mismatch' ? 'base-mismatch' : 'stale');
    expect(rejected.closes).toEqual([1, 1]);
    expect(accepted.closes).toEqual([0, 0]);
    expect(requests).toBe(rejection === 'base mismatch' ? 1 : 0);
  });
}

for (const rejection of ['returns false', 'throws', 'reset', 'dispose'] as const) {
  const succeeds = rejection === 'reset' || rejection === 'dispose';
  test(`${succeeds ? 'preserves painter ownership' : 'closes every bitmap'} when a transactional install ${rejection}`, () => {
    const attempted = bitmapFrame(1, 11);
    let presenter!: Presenter<string, ImageBitmap[]>;
    presenter = createPresenter<string, ImageBitmap[]>({
      transactionalPresentation: true, now: () => 0, requestFullFrame() {},
      discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
      echo: {
        paint() {}, drop() {},
        present() {
          if (rejection === 'throws') throw new Error('install failed');
          if (rejection === 'returns false') return false;
          if (rejection === 'reset') presenter.reset();
          else presenter.dispose();
        },
      },
    });
    expect(presenter.present(attempted.candidate)).toBe('stale');
    expect(attempted.closes).toEqual(succeeds ? [0, 0] : [1, 1]);
    if (succeeds) {
      attempted.candidate.payload.forEach((bitmap) => bitmap.close());
      expect(attempted.closes).toEqual([1, 1]);
    }
  });
}

test('closes superseded nested frames without closing the successful transactional frame', () => {
  const accepted = bitmapFrame(2, 12);
  const rejected = bitmapFrame(1, 11);
  let presenter!: Presenter<string, ImageBitmap[]>;
  presenter = createPresenter<string, ImageBitmap[]>({
    transactionalPresentation: true, now: () => 0, requestFullFrame() {},
    discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
    echo: {
      paint() {}, drop() {},
      present() { expect(presenter.present(rejected.candidate)).toBe('stale'); },
    },
  });
  expect(presenter.present(accepted.candidate)).toBe('presented');
  expect(rejected.closes).toEqual([1, 1]);
  expect(accepted.closes).toEqual([0, 0]);
});

test('closes a rejected delta even when requesting a full frame throws', () => {
  const rejected = bitmapFrame(1, 11, 99);
  const presenter = createPresenter<string, ImageBitmap[]>({
    requestFullFrame() { throw new Error('request failed'); },
    discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
  });
  expect(() => presenter.present(rejected.candidate)).toThrow('request failed');
  expect(rejected.closes).toEqual([1, 1]);
});

test('closes bitmaps after a failed cosmetic paint while preserving its coverage semantics', () => {
  for (const sameCoverage of [false, true]) {
    const accepted = bitmapFrame(0, 10);
    const rejected = bitmapFrame(sameCoverage ? 0 : 1, 11);
    const presenter = createPresenter<string, ImageBitmap[]>({
      sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
      discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
      echo: {
        paint() {}, drop() {},
        present(candidate) { if (candidate.revision === 11) throw new Error('paint failed'); },
      },
    });
    presenter.present(accepted.candidate);
    expect(presenter.present(rejected.candidate)).toBe('presented');
    expect(presenter.presentedThroughSeq).toBe(rejected.candidate.coversThroughSeq);
    expect(rejected.closes).toEqual([1, 1]);
    expect(accepted.closes).toEqual([0, 0]);
  }
});

test('preserves opt-in same-coverage frames and pending echo without discarding their bitmaps', () => {
  const first = bitmapFrame(0, 10);
  const redraw = bitmapFrame(0, 11, 10);
  const installs: number[] = [];
  const presenter = createPresenter<string, ImageBitmap[]>({
    sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
    discardFrame(candidate) { candidate.payload.forEach((bitmap) => bitmap.close()); },
    echo: { paint() {}, drop() {}, present(candidate) { installs.push(candidate.revision); } },
  });
  const entry = { ticket: ticket(1, 1).value, value: 'pending' };
  presenter.present(first.candidate);
  presenter.add(entry);
  expect(presenter.present(redraw.candidate)).toBe('presented');
  expect(installs).toEqual([10, 11]);
  expect(presenter.revision).toBe(11);
  expect(presenter.pending).toEqual([entry]);
  expect(redraw.closes).toEqual([0, 0]);
});

test('isolates a throwing discard hook from rejection and full-frame requests', () => {
  let discards = 0;
  let requests = 0;
  const presenter = createPresenter<string, string>({
    requestFullFrame() { requests += 1; },
    discardFrame() { discards += 1; throw new Error('discard failed'); },
  });
  expect(presenter.present(frame(1, 11, 99))).toBe('base-mismatch');
  presenter.dispose();
  expect(presenter.present(frame(2, 12))).toBe('stale');
  expect(discards).toBe(2);
  expect(requests).toBe(1);
});

for (const rejection of ['returns false', 'throws'] as const) {
  test(`keeps transactional coverage, revision, echo, and waiters unchanged when the hook ${rejection}`, async () => {
    const hooks = painters();
    let reject = false;
    let requests = 0;
    const observed: Array<{ coverage: number; revision: number | undefined; pending: readonly EchoEntry<string>[] }> = [];
    let presenter!: Presenter<string, string>;
    presenter = createPresenter<string, string>({
      transactionalPresentation: true, now: () => 0,
      requestFullFrame() { requests += 1; },
      echo: { ...hooks.echo, present(candidate, retired, remaining) {
        hooks.echo.present(candidate, retired, remaining);
        observed.push({ coverage: presenter.presentedThroughSeq, revision: presenter.revision, pending: presenter.pending });
        if (!reject) return;
        if (rejection === 'throws') throw new Error('install failed');
        return false;
      } },
    });
    presenter.present(frame(0, 10));
    const first = { ticket: ticket(1, 1).value, value: 'first' };
    const later = { ticket: ticket(2, 2).value, value: 'later' };
    presenter.add(first);
    presenter.add(later);
    let status: PresentationStatus | undefined;
    const waiting = presenter.whenPresented(1);
    void waiting.then((value) => { status = value; });
    reject = true;
    expect(presenter.present(frame(1, 11, 10))).toBe('stale');
    expect(observed[1]).toEqual({ coverage: 0, revision: 10, pending: [first, later] });
    expect(hooks.presents[1]).toEqual({ frame: frame(1, 11, 10), retired: [first], remaining: [later] });
    expect(presenter.presentedThroughSeq).toBe(0);
    expect(presenter.revision).toBe(10);
    expect(presenter.pending).toEqual([first, later]);
    expect(presenter.whenPresented(1)).toBe(waiting);
    await Promise.resolve();
    expect(status).toBeUndefined();
    expect(hooks.drops).toEqual([]);
    expect(hooks.paints).toEqual([[first], [first, later]]);
    expect(requests).toBe(0);
    reject = false;
    expect(presenter.present(frame(1, 12, 10))).toBe('presented');
    expect(await waiting).toBe('presented');
    expect(await presenter.whenPresented(1)).toBe('presented');
    expect(presenter.pending).toEqual([later]);
    expect(presenter.revision).toBe(12);
  });
}

test('preserves reset during a transactional hook or nested install and accepts new coverage', async () => {
  for (const nested of [false, true]) {
    const dispositions: string[] = [];
    const resetStates: Array<{ coverage: number; revision: number | undefined }> = [];
    let presenter!: Presenter<string, string>;
    presenter = createPresenter<string, string>({
      transactionalPresentation: true, now: () => 0, requestFullFrame() {},
      echo: {
        paint() {}, drop() {},
        present(candidate) {
          if (candidate.revision === 12) presenter.reset();
          if (candidate.revision !== 11) return;
          if (nested) dispositions.push(presenter.present(frame(2, 12)));
          else presenter.reset();
          resetStates.push({ coverage: presenter.presentedThroughSeq, revision: presenter.revision });
          if (nested) dispositions.push(presenter.present(frame(0, 0)));
        },
      },
    });
    expect(presenter.present(frame(0, 10))).toBe('presented');
    const entry = { ticket: ticket(1, 1).value, value: 'a' };
    presenter.add(entry);
    let status: PresentationStatus | undefined;
    const waiting = presenter.whenPresented(1).then((value) => { status = value; return value; });
    expect(presenter.present(frame(1, 11, 10))).toBe('stale');
    await Promise.resolve();
    expect(dispositions).toEqual(nested ? ['stale', 'presented'] : []);
    expect(resetStates).toEqual([{ coverage: -1, revision: undefined }]);
    expect(presenter.presentedThroughSeq).toBe(nested ? 0 : -1);
    expect(presenter.revision).toBe(nested ? 0 : undefined);
    expect(presenter.pending).toEqual([entry]);
    expect(status).toBeUndefined();
    if (!nested) expect(presenter.present(frame(0, 0))).toBe('presented');
    expect(presenter.present(frame(1, 1, 0))).toBe('presented');
    expect(await waiting).toBe('presented');
    expect(presenter.pending).toEqual([]);
  }
});

test('preserves a superseding transactional install through failures in two outer hooks', async () => {
  const nested: Array<{ coverage: number; disposition: string }> = [];
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    transactionalPresentation: true, now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.coversThroughSeq === 1) {
          nested.push({ coverage: 2, disposition: presenter.present(frame(2, 12)) });
          throw new Error('outer install failed');
        }
        if (candidate.coversThroughSeq === 2) {
          nested.push({ coverage: 3, disposition: presenter.present(frame(3, 13)) });
          return false;
        }
      },
    },
  });
  presenter.present(frame(0, 10));
  const entries = [1, 2, 3, 4].map((seq) => ({ ticket: ticket(seq, seq).value, value: String(seq) }));
  entries.forEach((entry) => presenter.add(entry));
  const waiting = [1, 2, 3].map((seq) => presenter.whenPresented(seq));
  expect(presenter.present(frame(1, 11, 10))).toBe('stale');
  expect(nested).toEqual([{ coverage: 3, disposition: 'presented' }, { coverage: 2, disposition: 'stale' }]);
  expect(presenter.presentedThroughSeq).toBe(3);
  expect(presenter.revision).toBe(13);
  expect(presenter.pending).toEqual([entries[3]]);
  expect(await Promise.all(waiting)).toEqual(['presented', 'presented', 'presented']);
  expect(await presenter.whenPresented(1)).toBe('presented');
});

test('rejects nested transactional frames that do not supersede the in-progress install', () => {
  const installed: number[] = [];
  const nested: string[] = [];
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    transactionalPresentation: true, sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        installed.push(candidate.revision);
        if (candidate.revision !== 13) return;
        nested.push(presenter.present(frame(1, 14)));
        nested.push(presenter.present(frame(2, 12)));
        nested.push(presenter.present(frame(2, 13)));
        return false;
      },
    },
  });
  presenter.present(frame(0, 10));
  expect(presenter.present(frame(2, 13))).toBe('stale');
  expect(nested).toEqual(['stale', 'stale', 'stale']);
  expect(installed).toEqual([10, 13]);
  expect(presenter.presentedThroughSeq).toBe(0);
  expect(presenter.revision).toBe(10);
});

test('commits a superseding same-coverage transactional revision without lowering it on outer success', () => {
  let nested: string | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    transactionalPresentation: true, sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.revision === 11) nested = presenter.present(frame(1, 13, 10));
      },
    },
  });
  presenter.present(frame(1, 10));
  expect(presenter.present(frame(1, 11, 10))).toBe('presented');
  expect(nested).toBe('presented');
  expect(presenter.revision).toBe(13);
  expect(presenter.presentedThroughSeq).toBe(1);
  expect(presenter.present(frame(1, 12))).toBe('stale');
});

test('keeps transactional same-coverage revision unchanged after rejection and accepts a later install', () => {
  for (const throws of [false, true]) {
    const hooks = painters();
    let reject = false;
    const presenter = createPresenter<string, string>({
      transactionalPresentation: true, sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
      echo: { ...hooks.echo, present(candidate, retired, remaining) {
        hooks.echo.present(candidate, retired, remaining);
        if (!reject) return;
        if (throws) throw new Error('redraw failed');
        return false;
      } },
    });
    presenter.present(frame(0, 10));
    const entry = { ticket: ticket(1, 1).value, value: 'pending' };
    presenter.add(entry);
    reject = true;
    expect(presenter.present(frame(0, 11, 10))).toBe('stale');
    expect(presenter.revision).toBe(10);
    expect(presenter.presentedThroughSeq).toBe(0);
    expect(presenter.pending).toEqual([entry]);
    expect(hooks.presents[1]).toEqual({ frame: frame(0, 11, 10), retired: [], remaining: [entry] });
    reject = false;
    expect(presenter.present(frame(0, 12, 10))).toBe('presented');
    expect(presenter.revision).toBe(12);
    expect(presenter.pending).toEqual([entry]);
    expect(hooks.drops).toEqual([]);
  }
});

test('retires only staged transactional echo and preserves entries added during the hook', () => {
  const hooks = painters();
  const first = { ticket: ticket(1, 1).value, value: 'first' };
  const later = { ticket: ticket(2, 2).value, value: 'later' };
  const added = { ticket: ticket(3, 1).value, value: 'added' };
  let coverageInHook: number | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    transactionalPresentation: true, now: () => 0, requestFullFrame() {},
    echo: { ...hooks.echo, present(candidate, retired, remaining) {
      hooks.echo.present(candidate, retired, remaining);
      if (candidate.coversThroughSeq !== 1) return;
      coverageInHook = presenter.presentedThroughSeq;
      presenter.add(added);
    } },
  });
  presenter.present(frame(0));
  presenter.add(first);
  presenter.add(later);
  expect(presenter.present(frame(1))).toBe('presented');
  expect(coverageInHook).toBe(0);
  expect(hooks.presents[1]).toEqual({ frame: frame(1), retired: [first], remaining: [later] });
  expect(presenter.pending).toEqual([later, added]);
});

test('advances coverage and drops retired waiters on a throwing hook by default', async () => {
  const hooks = painters();
  const presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: { ...hooks.echo, present() { throw new Error('install failed'); } },
  });
  presenter.add({ ticket: ticket(1, 1).value, value: 'first' });
  const waiting = presenter.whenPresented(1);
  expect(presenter.present(frame(1, 10))).toBe('presented');
  expect(presenter.presentedThroughSeq).toBe(1);
  expect(presenter.revision).toBe(10);
  expect(presenter.pending).toEqual([]);
  expect(await waiting).toBe('dropped');
});

test('ignores a false presentation result by default and presents covered waiters', async () => {
  const hooks = painters();
  const presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: { ...hooks.echo, present: () => false },
  });
  presenter.add({ ticket: ticket(1, 1).value, value: 'first' });
  const waiting = presenter.whenPresented(1);
  expect(presenter.present(frame(1))).toBe('presented');
  expect(presenter.presentedThroughSeq).toBe(1);
  expect(presenter.pending).toEqual([]);
  expect(await waiting).toBe('presented');
});

test('evicts the oldest echo at capacity without dropping its coverage waiter', async () => {
  const hooks = painters();
  const presenter = createPresenter({
    echo: hooks.echo, maxEchoEntries: 2, echoOverflow: 'evict-oldest', now: () => 0, requestFullFrame() {},
  });
  const entries = [1, 2, 3].map((seq) => ({ ticket: ticket(seq, seq).value, value: String(seq) }));
  presenter.add(entries[0]!);
  presenter.add(entries[1]!);
  let status: PresentationStatus | undefined;
  const waiting = presenter.whenPresented(1).then((value) => { status = value; return value; });
  expect(presenter.add(entries[2]!)).toBe(true);
  expect(presenter.pending).toEqual([entries[1], entries[2]]);
  expect(hooks.paints).toEqual([[entries[0]], [entries[0], entries[1]], [entries[1], entries[2]]]);
  expect(hooks.drops).toEqual([]);
  await Promise.resolve();
  expect(status).toBeUndefined();
  presenter.present(frame(3));
  expect(await waiting).toBe('presented');
  expect(await presenter.whenPresented(1)).toBe('presented');
  expect(hooks.presents[0]!.retired).toEqual([entries[1], entries[2]]);
});

test('keeps the replacement echo at capacity by silently evicting the oldest entry', async () => {
  const hooks = painters();
  const presenter = createPresenter({
    echo: hooks.echo, maxEchoEntries: 2, echoOverflow: 'evict-oldest', now: () => 0, requestFullFrame() {},
  });
  const first = { ticket: ticket(1, 1).value, value: 'first' };
  const later = { ticket: ticket(2, 2).value, value: 'later' };
  const replacement = { ticket: ticket(3, 3).value, value: 'replacement' };
  presenter.add(first);
  presenter.add(later);
  const waiting = presenter.whenPresented(1);
  presenter.replace([replacement.ticket], replacement);
  expect(presenter.pending).toEqual([later, replacement]);
  expect(hooks.paints).toEqual([[first], [first, later], [later, replacement]]);
  expect(hooks.drops).toEqual([]);
  presenter.present(frame(3));
  expect(await waiting).toBe('presented');
});

test('declines echo additions and replacements at zero capacity even with eviction enabled', () => {
  const hooks = painters();
  const presenter = createPresenter({
    echo: hooks.echo, maxEchoEntries: 0, echoOverflow: 'evict-oldest', now: () => 0, requestFullFrame() {},
  });
  const entry = { ticket: ticket(1, 1).value, value: 'first' };
  expect(presenter.add(entry)).toBe(false);
  presenter.replace([entry.ticket], entry);
  expect(presenter.pending).toEqual([]);
  expect(hooks.paints).toEqual([]);
  expect(hooks.drops).toEqual([]);
});

test('declines echo overflow by default and retains the existing pending entries', () => {
  const hooks = painters();
  const presenter = createPresenter({ echo: hooks.echo, maxEchoEntries: 1, now: () => 0, requestFullFrame() {} });
  const first = { ticket: ticket(1, 1).value, value: 'first' };
  const later = { ticket: ticket(2, 2).value, value: 'later' };
  expect(presenter.add(first)).toBe(true);
  expect(presenter.add(later)).toBe(false);
  presenter.replace([later.ticket], later);
  expect(presenter.pending).toEqual([first]);
  expect(hooks.paints).toEqual([[first]]);
  expect(hooks.drops).toEqual([]);
});

test('installs newer same-coverage frames only when opted in without retiring echo or settling waiters', async () => {
  for (const sameCoverageFrames of [undefined, false, true]) {
    const hooks = painters();
    const presenter = createPresenter({ echo: hooks.echo, sameCoverageFrames, now: () => 0, requestFullFrame() {} });
    presenter.present(frame(0, 10));
    const entry = { ticket: ticket(1, 1).value, value: 'pending' };
    presenter.add(entry);
    let status: PresentationStatus | undefined;
    const waiting = presenter.whenPresented(1).then((value) => { status = value; return value; });
    expect(presenter.present(frame(0, 11, 10))).toBe(sameCoverageFrames ? 'presented' : 'stale');
    expect(presenter.revision).toBe(sameCoverageFrames ? 11 : 10);
    expect(presenter.presentedThroughSeq).toBe(0);
    expect(presenter.pending).toEqual([entry]);
    expect(hooks.presents).toHaveLength(sameCoverageFrames ? 2 : 1);
    if (sameCoverageFrames) expect(hooks.presents[1]).toEqual({ frame: frame(0, 11, 10), retired: [], remaining: [entry] });
    expect(presenter.present(frame(0, 10))).toBe('stale');
    expect(presenter.present(frame(0, presenter.revision!))).toBe('stale');
    await Promise.resolve();
    expect(status).toBeUndefined();
    presenter.present(frame(1, 12));
    expect(await waiting).toBe('presented');
    expect(status).toBe('presented');
    expect(presenter.present(frame(0, 13))).toBe('stale');
    expect(hooks.drops).toEqual([]);
  }
});

test('checks same-coverage bases and contains throwing redraws without changing presentation outcomes', async () => {
  const hooks = painters();
  let requests = 0;
  let throws = false;
  const presenter = createPresenter<string, string>({
    sameCoverageFrames: true, now: () => 0,
    requestFullFrame: () => { requests += 1; },
    echo: { ...hooks.echo, present(candidate, retired, remaining) {
      hooks.echo.present(candidate, retired, remaining);
      if (throws) throw new Error('redraw');
    } },
  });
  presenter.drop([ticket(3, 0).value], 'refused');
  presenter.present(frame(0, 10));
  expect(await presenter.whenPresented(0)).toBe('dropped');
  const covered = { ticket: ticket(1, 0).value, value: 'late echo' };
  const future = { ticket: ticket(2, 1).value, value: 'future' };
  presenter.add(covered);
  presenter.add(future);
  let settled = false;
  const waiting = presenter.whenPresented(1).then((status) => { settled = true; return status; });
  expect(presenter.present(frame(0, 11, 9))).toBe('base-mismatch');
  expect(requests).toBe(1);
  expect(presenter.revision).toBe(10);
  expect(presenter.present(frame(0, 11, 10))).toBe('presented');
  expect(await presenter.whenPresented(0)).toBe('dropped');
  throws = true;
  expect(presenter.present(frame(0, 12, 11))).toBe('presented');
  expect(hooks.presents[1]).toEqual({ frame: frame(0, 11, 10), retired: [], remaining: [covered, future] });
  expect(hooks.presents[2]).toEqual({ frame: frame(0, 12, 11), retired: [], remaining: [covered, future] });
  expect(presenter.pending).toEqual([covered, future]);
  await Promise.resolve();
  expect(settled).toBe(false);
  expect(await presenter.whenPresented(0)).toBe('dropped');
  expect(presenter.revision).toBe(12);
  expect(hooks.drops).toEqual([]);
  throws = false;
  presenter.present(frame(1, 13, 12));
  expect(await waiting).toBe('presented');
});

test('keeps a reentrant revision 13 install while presenting revision 11 and rejects delayed revision 12', async () => {
  for (const coverage of [1, 2]) {
    const revisions: number[] = [];
    let revisionInHook: number | undefined;
    let nestedDisposition: string | undefined;
    let presenter!: Presenter<string, string>;
    presenter = createPresenter<string, string>({
      sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
      echo: {
        paint() {}, drop() {},
        present(candidate) {
          revisions.push(candidate.revision);
          if (candidate.revision === 11) {
            revisionInHook = presenter.revision;
            nestedDisposition = presenter.present(frame(coverage, 13));
          }
        },
      },
    });
    presenter.present(frame(1, 10));
    const waiting = presenter.whenPresented(coverage);
    expect(presenter.present(frame(coverage, 11, 10))).toBe('presented');
    expect(revisionInHook).toBe(11);
    expect(nestedDisposition).toBe('presented');
    expect(presenter.revision).toBe(13);
    expect(presenter.presentedThroughSeq).toBe(coverage);
    expect(presenter.present(frame(coverage, 12))).toBe('stale');
    expect(revisions).toEqual([10, 11, 13]);
    expect(await waiting).toBe('presented');
  }
});

test('keeps a reset performed reentrantly by a presentation hook', () => {
  for (const coverage of [1, 2]) {
    let presenter!: Presenter<string, string>;
    presenter = createPresenter<string, string>({
      sameCoverageFrames: true, now: () => 0, requestFullFrame() {},
      echo: {
        paint() {}, drop() {},
        present(candidate) { if (candidate.revision === 11) presenter.reset(); },
      },
    });
    presenter.present(frame(1, 10));
    presenter.present(frame(coverage, 11, 10));
    expect(presenter.revision).toBeUndefined();
    expect(presenter.presentedThroughSeq).toBe(-1);
    expect(presenter.present(frame(0, 1))).toBe('presented');
  }
});

test('settles reentrant coverage waiters as dropped when the presentation hook throws', async () => {
  let reentrant: Promise<PresentationStatus> | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.coversThroughSeq !== 1) return;
        reentrant = presenter.whenPresented(1);
        throw new Error('present');
      },
    },
  });
  presenter.present(frame(0));
  presenter.add({ ticket: ticket(1, 1).value, value: 'first' });
  const waiting = presenter.whenPresented(1);
  expect(presenter.present(frame(1))).toBe('presented');
  expect(reentrant).toBe(waiting);
  expect(await waiting).toBe('dropped');
  expect(await reentrant).toBe('dropped');
  expect(await presenter.whenPresented(1)).toBe('dropped');
  expect(presenter.presentedThroughSeq).toBe(1);
});

test('preserves nested successful coverage when the outer presentation hook throws', async () => {
  let nestedDisposition: string | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.coversThroughSeq !== 1) return;
        nestedDisposition = presenter.present(frame(2));
        throw new Error('present');
      },
    },
  });
  presenter.present(frame(0));
  for (let seq = 1; seq <= 3; seq += 1) {
    presenter.add({ ticket: ticket(seq, seq).value, value: String(seq) });
  }
  const first = presenter.whenPresented(1);
  const second = presenter.whenPresented(2);
  let futureStatus: PresentationStatus | undefined;
  const future = presenter.whenPresented(3).then((status) => { futureStatus = status; return status; });
  expect(presenter.present(frame(1))).toBe('presented');
  expect(nestedDisposition).toBe('presented');
  expect(await first).toBe('presented');
  expect(await second).toBe('presented');
  expect(await presenter.whenPresented(1)).toBe('presented');
  expect(await presenter.whenPresented(2)).toBe('presented');
  expect(futureStatus).toBeUndefined();
  expect(presenter.presentedThroughSeq).toBe(2);
  expect(presenter.revision).toBe(2);
  expect(presenter.pending.map((entry) => entry.value)).toEqual(['3']);
  presenter.present(frame(3));
  expect(await future).toBe('presented');
});

test('drops later coverage lookups when both nested presentation hooks throw without prior waiters', async () => {
  let nestedDisposition: string | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.coversThroughSeq === 0) return;
        if (candidate.coversThroughSeq === 1) nestedDisposition = presenter.present(frame(2));
        throw new Error('present');
      },
    },
  });
  presenter.present(frame(0));
  expect(presenter.present(frame(1))).toBe('presented');
  expect(nestedDisposition).toBe('presented');
  expect(await presenter.whenPresented(1)).toBe('dropped');
  expect(await presenter.whenPresented(2)).toBe('dropped');
});

test('drops an existing outer coverage waiter when both nested presentation hooks throw', async () => {
  let nestedDisposition: string | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.coversThroughSeq === 0) return;
        if (candidate.coversThroughSeq === 1) nestedDisposition = presenter.present(frame(2));
        throw new Error('present');
      },
    },
  });
  presenter.present(frame(0));
  presenter.add({ ticket: ticket(1, 1).value, value: 'first' });
  const waiting = presenter.whenPresented(1);
  expect(presenter.present(frame(1))).toBe('presented');
  expect(nestedDisposition).toBe('presented');
  expect(await waiting).toBe('dropped');
  expect(await presenter.whenPresented(1)).toBe('dropped');
  expect(await presenter.whenPresented(2)).toBe('dropped');
});

test('presents outer coverage waiters when the inner presentation hook throws and the outer hook succeeds', async () => {
  let nestedDisposition: string | undefined;
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    now: () => 0, requestFullFrame() {},
    echo: {
      paint() {}, drop() {},
      present(candidate) {
        if (candidate.coversThroughSeq === 1) nestedDisposition = presenter.present(frame(2));
        if (candidate.coversThroughSeq === 2) throw new Error('present');
      },
    },
  });
  presenter.present(frame(0));
  presenter.add({ ticket: ticket(1, 1).value, value: 'first' });
  const waiting = presenter.whenPresented(1);
  expect(presenter.present(frame(1))).toBe('presented');
  expect(nestedDisposition).toBe('presented');
  expect(await waiting).toBe('presented');
  expect(await presenter.whenPresented(1)).toBe('presented');
  expect(await presenter.whenPresented(2)).toBe('dropped');
});

test('counts nested hook timing once within the cumulative event budget', () => {
  let clock = 0;
  const samples: EchoOverrunSample[] = [];
  let presenter!: Presenter<string, string>;
  presenter = createPresenter({
    now: () => clock, requestFullFrame() {},
    onEchoOverrun: (sample) => { samples.push(sample); },
    echo: {
      paint() { clock += 0.75; }, drop() {},
      present() {
        clock += 0.25;
        presenter.add({ ticket: ticket(1).value, value: 'nested' });
        clock += 0.5;
      },
    },
  });
  presenter.event(() => {
    presenter.present(frame(0));
    expect(samples).toEqual([]);
    presenter.add({ ticket: ticket(2).value, value: 'later' });
  });
  expect(samples).toEqual([{ hook: 'paint', elapsedMs: 0.75, eventTotalMs: 2.25 }]);
});

test('commits remaining echo before presenting and keeps reentrant additions', () => {
  const first = ticket(1, 1).value;
  const later = ticket(2, 2).value;
  const added = ticket(3).value;
  let observed: string[] = [];
  let presenter!: Presenter<string, string>;
  presenter = createPresenter<string, string>({
    requestFullFrame() {}, now: () => 0,
    echo: {
      paint() {}, drop() {},
      present() {
        observed = presenter.pending.map((entry) => entry.value);
        presenter.add({ ticket: added, value: 'added' });
      },
    },
  });
  presenter.add({ ticket: first, value: 'first' });
  presenter.add({ ticket: later, value: 'later' });
  presenter.present(frame(1));
  expect(observed).toEqual(['later']);
  expect(presenter.pending.map((entry) => entry.value)).toEqual(['later', 'added']);
});

test('settles ack before presentation and resolves whenPresented only on covering installation', async () => {
  const presenter = createPresenter<string, string>({ requestFullFrame() {} });
  const command = ticket(1, 4);
  let status: PresentationStatus | undefined;
  const presented = presenter.whenPresented(4).then((value) => { status = value; return value; });
  command.acknowledge();
  expect(await command.value.ack).toEqual({ seq: 4, status: 'applied' });
  expect(status).toBeUndefined();
  presenter.present(frame(3));
  await Promise.resolve();
  expect(status).toBeUndefined();
  presenter.present(frame(5));
  expect(await presented).toBe('presented');
  expect(await presenter.whenPresented(4)).toBe('presented');
});

test('drops stale frames and requests a full frame on delta base mismatch', async () => {
  const hooks = painters();
  let requests = 0;
  const presenter = createPresenter({ echo: hooks.echo, requestFullFrame: () => { requests += 1; } });
  expect(presenter.present(frame(2, 10))).toBe('presented');
  expect(presenter.present(frame(1, 11, 999))).toBe('stale');
  expect(presenter.present(frame(2, 12))).toBe('stale');
  let settled = false;
  const pending = presenter.whenPresented(3).then((value) => { settled = true; return value; });
  expect(presenter.present(frame(3, 11, 9))).toBe('base-mismatch');
  await Promise.resolve();
  expect(settled).toBe(false);
  expect(requests).toBe(1);
  expect(presenter.revision).toBe(10);
  expect(presenter.presentedThroughSeq).toBe(2);
  expect(hooks.presents.length).toBe(1);
  expect(presenter.present(frame(3, 12))).toBe('presented');
  expect(await pending).toBe('presented');
  expect(presenter.present(frame(4, 13, 12))).toBe('presented');
  expect(presenter.revision).toBe(13);
});

test('reconciles reversed ack delivery by coverage and retires echo atomically', async () => {
  const hooks = painters();
  const presenter = createPresenter({ echo: hooks.echo, requestFullFrame() {} });
  const late = ticket(1, 5);
  const early = ticket(2, 3);
  const unsent = ticket(3);
  const entries = [
    { ticket: late.value, value: 'late' },
    { ticket: early.value, value: 'early' },
    { ticket: unsent.value, value: 'unsent' },
  ];
  entries.forEach((entry) => presenter.add(entry));
  const earlyPresented = presenter.whenPresented(3);
  presenter.present(frame(4, 20));
  expect(hooks.presents).toEqual([{ frame: frame(4, 20), retired: [entries[1]], remaining: [entries[0], entries[2]] }]);
  expect(presenter.pending).toEqual([entries[0], entries[2]]);
  expect(await earlyPresented).toBe('presented');
  expect(presenter.present(frame(3, 19))).toBe('stale');
  unsent.send(6);
  presenter.present(frame(5, 21, 20));
  expect(hooks.presents[1]!.retired).toEqual([entries[0]]);
  expect(hooks.presents[1]!.remaining).toEqual([entries[2]]);
  expect(hooks.paints.length).toBe(3);
  presenter.present(frame(6, 22, 21));
  expect(presenter.pending).toEqual([]);
});

test('installs open frames, bounds echo, and reports measured cumulative painter overruns', () => {
  let clock = 0;
  const samples: EchoOverrunSample[] = [];
  const costs = { paint: 1.25, drop: 0.5, present: 0.25 };
  const presenter = createPresenter<string, string>({
    echo: {
      paint() { clock += costs.paint; },
      drop() { clock += costs.drop; },
      present() { clock += costs.present; },
    },
    now: () => clock,
    onEchoOverrun: (sample) => { samples.push(sample); },
    requestFullFrame() {},
  });
  expect(presenter.presentedThroughSeq).toBe(-1);
  expect(presenter.revision).toBeUndefined();
  expect(presenter.present(frame(0, 0))).toBe('presented');
  presenter.event(() => {
    presenter.add({ ticket: ticket(1).value, value: 'a' });
    presenter.event(() => presenter.add({ ticket: ticket(2).value, value: 'b' }));
    presenter.drop([presenter.pending[0]!.ticket], 'refused');
    presenter.present(frame(1, 1));
  });
  expect(samples).toEqual([
    { hook: 'paint', elapsedMs: 1.25, eventTotalMs: 2.5 },
    { hook: 'drop', elapsedMs: 0.5, eventTotalMs: 3 },
    { hook: 'present', elapsedMs: 0.25, eventTotalMs: 3.25 },
  ]);
  presenter.add({ ticket: ticket(3).value, value: 'c' });
  expect(samples.length).toBe(3);
  for (let id = 4; id <= 65; id += 1) expect(presenter.add({ ticket: ticket(id).value, value: 'x' })).toBe(true);
  expect(presenter.pending.length).toBe(64);
  expect(presenter.add({ ticket: ticket(66).value, value: 'overflow' })).toBe(false);
  expect(samples.length).toBe(3);
});

test('replaces alias echoes with the merged echo in one paint and honors configured caps', () => {
  const hooks = painters();
  const presenter = createPresenter({ echo: hooks.echo, maxEchoEntries: 2, requestFullFrame() {} });
  const first = ticket(1).value;
  const second = ticket(2).value;
  const third = ticket(3).value;
  presenter.add({ ticket: first, value: 'a' });
  presenter.add({ ticket: second, value: 'b' });
  expect(presenter.add({ ticket: third, value: 'c' })).toBe(false);
  presenter.replace([first, second], { ticket: first, value: 'ab' });
  expect(presenter.pending).toEqual([{ ticket: first, value: 'ab' }]);
  expect(hooks.paints.length).toBe(3);
  expect(hooks.drops).toEqual([]);
  presenter.replace([first]);
  expect(presenter.pending).toEqual([]);
  expect(hooks.paints[3]).toEqual([]);
  const disabled = createPresenter<string, string>({ maxEchoEntries: 0, requestFullFrame() {} });
  expect(disabled.add({ ticket: third, value: 'c' })).toBe(false);
});

test('drops refused and cancelled echo with damaged repaint and settles even without echo', async () => {
  const hooks = painters();
  const presenter = createPresenter({ echo: hooks.echo, requestFullFrame() {} });
  const refused = ticket(1, 1).value;
  const cancelled = ticket(2, 2).value;
  const remaining = { ticket: ticket(3, 3).value, value: 'remaining' };
  const entry = { ticket: refused, value: 'refused' };
  presenter.add(entry);
  presenter.add(remaining);
  const refusedWait = presenter.whenPresented(1);
  const cancelledWait = presenter.whenPresented(2);
  presenter.drop([refused], 'refused');
  presenter.drop([cancelled], 'dependency-refused');
  expect(hooks.drops).toEqual([{ entries: [entry], reason: 'refused' }]);
  expect(presenter.pending).toEqual([remaining]);
  expect(await refusedWait).toBe('dropped');
  expect(await cancelledWait).toBe('dropped');
  presenter.present(frame(3));
  expect(await presenter.whenPresented(1)).toBe('dropped');
  expect(await presenter.whenPresented(2)).toBe('dropped');
  expect(await presenter.whenPresented(3)).toBe('presented');
});

test('uses explicit echo seq for coverage and dropping before the ticket getter changes', async () => {
  const presenter = createPresenter<string, string>({ requestFullFrame() {} });
  const command = ticket(1).value;
  presenter.add({ ticket: command, seq: 4, value: 'echo' });
  const waiting = presenter.whenPresented(4);
  presenter.drop([command], 'cancelled');
  expect(await waiting).toBe('dropped');
  expect(presenter.pending).toEqual([]);
});

test('reset retains echo and waiters but requires a full frame from the new incarnation', async () => {
  const hooks = painters();
  let requests = 0;
  const presenter = createPresenter({ echo: hooks.echo, requestFullFrame: () => { requests += 1; } });
  presenter.present(frame(4, 10));
  const entry = { ticket: ticket(1, 5).value, value: 'pending' };
  presenter.add(entry);
  const pending = presenter.whenPresented(5);
  presenter.reset();
  expect(presenter.presentedThroughSeq).toBe(-1);
  expect(presenter.revision).toBeUndefined();
  expect(presenter.pending).toEqual([entry]);
  expect(presenter.present(frame(5, 11, 10))).toBe('base-mismatch');
  expect(requests).toBe(1);
  presenter.present(frame(0, 0));
  expect(presenter.pending).toEqual([entry]);
  presenter.present(frame(5, 1, 0));
  expect(await pending).toBe('presented');
  expect(presenter.pending).toEqual([]);
});

test('dispose drops pending waiters and echo idempotently and ignores later frames', async () => {
  const hooks = painters();
  const presenter = createPresenter({ echo: hooks.echo, requestFullFrame() {} });
  const entry = { ticket: ticket(1, 1).value, value: 'pending' };
  presenter.add(entry);
  const pending = presenter.whenPresented(1);
  presenter.dispose();
  presenter.dispose();
  expect(await pending).toBe('dropped');
  expect(await presenter.whenPresented(2)).toBe('dropped');
  expect(hooks.drops).toEqual([{ entries: [entry], reason: 'disposed' }]);
  expect(presenter.pending).toEqual([]);
  expect(presenter.add(entry)).toBe(false);
  expect(presenter.present(frame(2))).toBe('stale');
  expect(hooks.presents).toEqual([]);
});

test('shares waiters for the same seq and bounds distinct pending coverage requests', async () => {
  const presenter = createPresenter<string, string>({ requestFullFrame() {} });
  const oldest = presenter.whenPresented(1);
  expect(presenter.whenPresented(1)).toBe(oldest);
  for (let seq = 2; seq <= 1025; seq += 1) presenter.whenPresented(seq);
  expect(await oldest).toBe('dropped');
  expect(await presenter.whenPresented(1)).toBe('dropped');
  const latest = presenter.whenPresented(1025);
  presenter.present(frame(1025));
  expect(await latest).toBe('presented');
});

test('retains only the latest 256 dropped seqs while settling existing waiters exactly', async () => {
  const presenter = createPresenter<string, string>({ requestFullFrame() {} });
  for (let seq = 1; seq <= 257; seq += 1) {
    const waiting = presenter.whenPresented(seq);
    presenter.drop([ticket(seq, seq).value], 'refused');
    expect(await waiting).toBe('dropped');
  }
  presenter.present(frame(257));
  expect(await presenter.whenPresented(1)).toBe('presented');
  expect(await presenter.whenPresented(2)).toBe('dropped');
  expect(await presenter.whenPresented(257)).toBe('dropped');
});

test('starts a fresh painter budget after a hook throws', () => {
  let clock = 0;
  let fail = true;
  const samples: EchoOverrunSample[] = [];
  const presenter = createPresenter<string, string>({
    echo: {
      paint() {
        clock += fail ? 3 : 1;
        if (fail) throw new Error('paint failed');
      },
      drop() {},
      present() {},
    },
    now: () => clock,
    onEchoOverrun: (sample) => { samples.push(sample); },
    requestFullFrame() {},
  });
  expect(presenter.add({ ticket: ticket(1).value, value: 'a' })).toBe(true);
  expect(samples).toEqual([{ hook: 'paint', elapsedMs: 3, eventTotalMs: 3 }]);
  fail = false;
  expect(presenter.add({ ticket: ticket(2).value, value: 'b' })).toBe(true);
  expect(samples).toEqual([{ hook: 'paint', elapsedMs: 3, eventTotalMs: 3 }]);
});

test('contains all throwing hooks, measures them, and drops failed presentations', async () => {
  let clock = 0;
  let fail = true;
  const samples: EchoOverrunSample[] = [];
  const presenter = createPresenter<string, string>({
    echo: {
      paint() { clock += 3; throw new Error('paint'); },
      drop() { clock += 3; throw new Error('drop'); },
      present() { clock += fail ? 3 : 1; if (fail) throw new Error('present'); },
    },
    now: () => clock, requestFullFrame() {},
    onEchoOverrun: (sample) => { samples.push(sample); },
  });
  const refused = ticket(1, 1).value;
  const applied = ticket(2, 2).value;
  presenter.event(() => {
    expect(presenter.add({ ticket: refused, value: 'refused' })).toBe(true);
    expect(presenter.add({ ticket: applied, value: 'applied' })).toBe(true);
    presenter.drop([refused], 'refused');
  });
  const waiting = presenter.whenPresented(2);
  const withoutEcho = presenter.whenPresented(3);
  expect(presenter.present(frame(4))).toBe('presented');
  expect(await waiting).toBe('dropped');
  expect(await withoutEcho).toBe('dropped');
  expect(await presenter.whenPresented(2)).toBe('dropped');
  expect(await presenter.whenPresented(4)).toBe('dropped');
  expect(presenter.pending).toEqual([]);
  expect(presenter.presentedThroughSeq).toBe(4);
  expect(presenter.revision).toBe(4);
  expect(samples).toEqual([
    { hook: 'paint', elapsedMs: 3, eventTotalMs: 3 },
    { hook: 'paint', elapsedMs: 3, eventTotalMs: 6 },
    { hook: 'drop', elapsedMs: 3, eventTotalMs: 9 },
    { hook: 'present', elapsedMs: 3, eventTotalMs: 3 },
  ]);
  fail = false;
  expect(presenter.present(frame(5, 5, 4))).toBe('presented');
  expect(await presenter.whenPresented(5)).toBe('presented');
  expect(samples).toHaveLength(4);
});

import { afterEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as queryModule from '../layout/render/displayListQueries';
import type { DisplayPage } from '../layout/render/displayList';
import { createDisplayListQueries, type DisplayListRect } from '../layout/render/displayListQueries';
import { FRAME_DELTA_VERSION, retainedFramePageById, type RetainedFrame } from '../layout/render/frameDelta';
import { loadRustDisplayListQueryEngine, type RustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import type { YrsLoc, YrsStorySegment } from './index';
import { createResidentDocxCommandHooks } from './residentCommandActions';
import { releaseResidentSelectionReads } from './residentCommandSelection';
import { ResidentSessionFonts } from './residentFontEnvironment';
import type { DocxCommandApplyContext, DocxCommandEnvelope, DocxCommandIntent, ResidentCommandHooks, ResidentCommandHooksFactory, ResidentCommandHostOptions, ResidentCommandInitRequest } from './residentCommandProtocol';

const releasesAfterTest: (() => void)[] = [];

afterEach(() => {
  for (const release of releasesAfterTest.splice(0)) release();
});

function splitRowWorker(direction: 'up' | 'down', queryEngine?: RustDisplayListQueryEngine, boundary?: 'cold' | 'warm', malformedTable: boolean | 'anonymous' = false, options: { caretPage?: number; malformedNeighbour?: boolean; warmCaret?: boolean; trailingEnd?: number; caretHint?: 'missing' | 'misleading'; productionHooks?: boolean } = {}) {
  const focus: YrsLoc = { story: boundary ? 'body' : 'body:t0:r0c0', paraId: boundary ? 'wrapped' : 'cell', offset: 20 };
  const position = 24;
  const caretPage = options.caretPage ?? 1;
  const neighbour = caretPage + (direction === 'up' ? -1 : 1);
  const builds: number[][] = [];
  const releases: number[][] = [];
  const queryFrames: { frameEpoch: number; disposed: boolean; failed: boolean; verticalMoves: number }[] = [];
  const navigationQueryFrames: typeof queryFrames = [];
  const caretHints: (DisplayListRect | null)[] = [];
  const nativeQueries = createDisplayListQueries;
  const retain = (pages: DisplayPage[], frameEpoch: number): RetainedFrame => ({
    protocolVersion: FRAME_DELTA_VERSION, docEpoch: 1, layoutEpoch: 1, frameEpoch,
    pages: pages.map((page) => ({ page, pageIndex: page.pageIndex, pageId: BigInt(page.pageIndex + 1),
      fingerprint: 0n, primitiveIds: new BigUint64Array() })),
    damagedPageIds: new Set(), removedPageIds: new Set(), displayList: { pages },
  });
  let frame = retain(Array.from({ length: 80 }, (_, pageIndex) => ({
    pageIndex, width: 600, height: 800, primitives: [], unbuilt: true,
    positionSpan: [position - 20, position + (caretPage === 1 || options.caretHint ? 1600 : 220)] as [number, number],
  })), 1);
  const text = (value: string, docStart: number, docEnd: number, baselineY = 40): DisplayPage['primitives'][number] => ({
    kind: 'text', text: value, x: 20, baselineY, width: value.length * 10,
    font: '400 16px Arial', color: '#000000', docStart, docEnd,
    paraId: malformedTable === 'anonymous' ? undefined : boundary || docEnd <= position ? focus.paraId : 'next',
    ...(boundary ? {} : {
      table: malformedTable ? { tableId: 't0' } : { tableId: 't0', rowStart: 0, rowEnd: 1, rowCount: 1, columnCount: 1 },
      cell: { row: 0, col: 0, rowSpan: 1, colSpan: 1, cellId: focus.story },
    }),
  });
  const primitives = (pageIndex: number): DisplayPage['primitives'] => {
    const result = pageIndex < caretPage
      ? [text('abcdefghij', 4, options.caretHint ? position - 1 : 14), ...(boundary ? [text('klmnopqrst', 14, position, 80)] : [])]
      : pageIndex === caretPage ? [boundary ? text('uvwx', position, 28) : text('qrst', 20, options.trailingEnd ?? position)]
      : [boundary ? text('yzabcd', 28, 34) : text('next', 25, 29)];
    return options.malformedNeighbour && pageIndex === neighbour
      ? result.map((primitive) => ({ ...primitive, table: { tableId: 't0' } })) : result;
  };
  const changePages = (pages: number[], unbuilt: boolean, frameEpoch: number) => {
    expect(frameEpoch).toBe(frame.frameEpoch);
    frame = retain(frame.displayList.pages.map<DisplayPage>((page) => pages.includes(page.pageIndex)
      ? { ...page, unbuilt, primitives: !queryEngine || unbuilt || page.pageIndex > caretPage + 1 ? [] : primitives(page.pageIndex) }
      : page), frameEpoch + 1);
    return frame;
  };
  if (boundary === 'warm') changePages([0], false, frame.frameEpoch);
  if (options.warmCaret) changePages([caretPage], false, frame.frameEpoch);
  let read: RetainedFrame | undefined;
  const factory: ResidentCommandHooksFactory = (engine, init, environment) => options.productionHooks
    ? createResidentDocxCommandHooks(engine, init, { readSelectionLayout: () => {
      const layout = environment!.readSelectionLayout();
      read = layout.frame;
      return layout;
    } }) : {
      preflight: (command) => {
        read = environment!.readSelectionLayout().frame;
        return { status: 'ready', prepared: { command: command as DocxCommandIntent, selection: null, range: null } };
      },
      apply: () => { throw new Error('Unused apply'); },
      selection: () => null,
      editState: () => { throw new Error('Unused edit state'); },
    };
  const sticky = new Map<number, YrsLoc>();
  const source = readFileSync(new URL('./residentEngineWorker.ts', import.meta.url), 'utf8');
  const parseFailures = source.indexOf('let queryParseFailures = 0;');
  const binding = source.slice(parseFailures < 0 ? source.indexOf('function bindResidentCommands(') : parseFailures,
    source.indexOf('function replacementCommandInit('));
  const dependencies = {
    sessionFonts: null, ResidentSessionFonts, scope: {}, proposals: null, resetResidentObjectHandles() {},
    createProposalRegistry: () => ({ snapshot: () => ({ previewVersion: 0, proposals: [] }) }),
    session: {
      proposalEngine: { version: () => '0' }, selection: () => ({ anchor: focus, head: focus }),
      setSelection: () => {},
      encodeStickyPosition: (at: YrsLoc) => {
        const key = sticky.size + 1;
        sticky.set(key, at);
        return { story: at.story, encoded: Uint8Array.of(key) };
      },
      resolveStickyPosition: (at: { encoded: Uint8Array }) => sticky.get(at.encoded[0]!) ?? null,
      geometryReader: {
        version: () => '0',
        hasStory: (story: string) => story === 'body' || story === focus.story,
        storyIds: () => ['body', focus.story],
        paragraphSpans: (story: string) => story === focus.story
          ? [{ paraId: focus.paraId, length: 20 }, { paraId: 'next', length: 4 }] : [{ paraId: 'body', length: 1 }],
        paragraphIdCount: (story: string, paraId: string) => story === focus.story && [focus.paraId, 'next'].includes(paraId) ? 1 : 0,
        storySegments: (story: string): YrsStorySegment[] => story === 'body' ? [{
          kind: 'embed', embedKind: 'table', payload: { rows: [{ cells: [{ story: focus.story }] }] }, attributes: {},
        }, { kind: 'pilcrow', paraId: 'body', properties: {}, attributes: {} }] : [
          { kind: 'text', text: 'abcdefghijklmnopqrst', attributes: {} },
          { kind: 'pilcrow', paraId: focus.paraId, properties: {}, attributes: {} },
          { kind: 'text', text: 'next', attributes: {} },
          { kind: 'pilcrow', paraId: 'next', properties: {}, attributes: {} },
        ],
        locateParagraph: (story: string, paraId: string) => {
          expect(story).toBe(focus.story);
          if (options.productionHooks && paraId === 'next') return { start: 28, end: 32 };
          expect(paraId).toBe(focus.paraId);
          return { start: 7, end: 7 + (boundary ? 30 : 20) };
        },
      },
      buildDisplayPagesFrame: (pages: number[], frameEpoch: number) => {
        builds.push([...pages]);
        return changePages(pages, false, frameEpoch);
      },
      releaseDisplayPagesFrame: (pages: number[], frameEpoch: number) => {
        releases.push([...pages]);
        return changePages(pages, true, frameEpoch);
      },
    },
    commandHooksFactory: factory,
    collaboration: null,
    init: { epoch: 1, readOnly: false, layoutWithRegions: true, layoutInput: '{}', extras: '{}' } as ResidentCommandInitRequest,
    retainedFrame: frame, retainedLayoutVersion: null, layoutRevision: 0, commandFrameChanges: null,
    commandInit: null, commandHost: null, releaseCommandHost: null, releaseResidentSelectionReads,
    createResidentCommandHost: (options: ResidentCommandHostOptions) => options.hooks,
    displayPositionIndex: () => ({ positionOf: (at: YrsLoc, root: string) => {
      expect(at).toEqual(focus);
      expect(root).toBe('body');
      return position;
    } }),
    createDisplayListQueries: (list: RetainedFrame['displayList']) => {
      const entry = { frameEpoch: frame.frameEpoch, disposed: false, failed: false, verticalMoves: 0 };
      queryFrames.push(entry);
      if (queryEngine) {
        const queries = nativeQueries(list, queryEngine);
        const record = <T>(result: T): T => {
          entry.failed ||= queries.sourceState().status === 'error';
          return result;
        };
        return {
          ...queries,
          caretRect: (at: number) => {
            const rect = record(queries.caretRect(at));
            const hint = options.caretHint === 'missing' && rect?.height === 0 ? null : rect;
            caretHints.push(hint);
            return hint;
          },
          rangeRectsOnPages: (from: number, to: number, first: number, last: number) =>
            record(queries.rangeRectsOnPages?.(from, to, first, last) ?? []),
          verticalMove: (at: number, direction: 'up' | 'down', goalX?: number) => {
            entry.verticalMoves += 1;
            return record(queries.verticalMove(at, direction, goalX));
          },
          dispose: () => { entry.disposed = true; queries.dispose(); },
        };
      }
      return {
        dispose: () => { entry.disposed = true; },
        caretRect: (at: number) => {
          expect(at).toBe(position);
          if (!list.pages[caretPage]!.unbuilt) return { pageIndex: caretPage, x: 20, y: 40, width: 0, height: 10 };
          return list.pages[0]!.unbuilt ? null : { pageIndex: 0, x: 20, y: 40, width: 0, height: 0 };
        },
      };
    },
    loadRustDisplayListQueryEngine: async () => {}, loadedRustDisplayListQueryEngine: () => queryEngine,
    applyFrameDelta: (_frame: RetainedFrame, next: RetainedFrame) => next,
    decodeFrameDelta: (next: RetainedFrame) => next,
    retainedPageByKey: (current: RetainedFrame, pageId: string) => retainedFramePageById(current, BigInt(pageId)),
  };
  const worker = new Function(...Object.keys(dependencies),
    new Bun.Transpiler({ loader: 'ts' }).transformSync(binding) +
    '\nlet hooks = bindResidentCommands(init); return { get hooks() { return hooks; }, rebind: () => { hooks = bindResidentCommands(init); }, release: () => { releaseCommandHost?.(); releaseCommandHost = null; }, changes: commandFrameChanges, install: (frame) => { retainedFrame = frame; } };'
  )(...Object.values(dependencies)) as { hooks: ResidentCommandHooks; rebind(): void; release(): void; changes: { pages: Set<number> }; install(frame: RetainedFrame): void };
  releasesAfterTest.push(worker.release);
  return {
    builds, releases, queryFrames, navigationQueryFrames, caretHints, focus, caretPage, neighbour, position, changes: worker.changes,
    frame: () => read!,
    rebind: () => worker.rebind(),
    probeDisplayList(pageIndex: number) {
      return { pages: frame.displayList.pages.map((page) => page.pageIndex === pageIndex
        ? { ...page, unbuilt: false, primitives: primitives(pageIndex) } : page) };
    },
    async navigate() {
      const queries = options.productionHooks ? spyOn(queryModule, 'createDisplayListQueries').mockImplementation((list) => {
        const facade = dependencies.createDisplayListQueries(list);
        navigationQueryFrames.push(queryFrames.at(-1)!);
        return facade as queryModule.DisplayListQueries;
      }) : undefined;
      try {
        const command = { kind: 'navigate', basis: { epoch: 1, version: 0 },
          payload: { unit: 'line', direction, extend: false, goalX: 20 } } as DocxCommandEnvelope;
        const result = await worker.hooks.preflight(command, { basis: command.basis, currentVersion: 0 } as DocxCommandApplyContext);
        expect(result.status).toBe('ready');
        return result.status === 'ready' ? result.prepared.selection : null;
      } finally {
        queries?.mockRestore();
      }
    },
    evictNeighbour() { worker.install(changePages([neighbour], true, frame.frameEpoch)); },
  };
}

test('rebinding_command_host_disposes_selection_queries', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  const worker = splitRowWorker('down', engine, undefined, false, { productionHooks: true });
  await worker.navigate();
  expect(worker.navigationQueryFrames.length).toBeGreaterThan(0);
  const previous = [...worker.queryFrames];
  expect(previous.some((entry) => !entry.disposed)).toBe(true);
  worker.rebind();
  expect(previous.every((entry) => entry.disposed)).toBe(true);
  await worker.navigate();
  expect(worker.queryFrames.at(-1)!.disposed).toBe(false);
});

test('long_split_row_arrows_stop_after_caret_and_neighbour', async () => {
  for (const direction of ['down', 'up'] as const) {
    const worker = splitRowWorker(direction);
    await worker.navigate();
    expect(worker.builds.flat().length).toBeLessThanOrEqual(3);
    expect(worker.builds).toEqual([[0], [1], [worker.neighbour]]);
    expect(worker.frame().pages[1]!.page.unbuilt).toBe(false);
    expect(worker.frame().pages[worker.neighbour]!.page.unbuilt).toBe(false);
    expect(worker.frame().pages.slice(3).every(({ page }) => page.unbuilt)).toBe(true);
    worker.builds.length = 0;
    worker.evictNeighbour();
    await worker.navigate();
    expect(worker.builds.flat().length).toBeLessThanOrEqual(2);
    expect(worker.builds).toEqual([[worker.neighbour]]);
    worker.builds.length = 0;
    await worker.navigate();
    expect(worker.builds).toEqual([]);
    expect(worker.queryFrames.slice(0, -1).every((entry) => entry.disposed)).toBe(true);
    expect(worker.queryFrames.at(-1)!.disposed).toBe(false);
  }
});

test('split_row_probe_releases_misses', async () => {
  const worker = splitRowWorker('down');
  await worker.navigate();
  expect(worker.releases).toEqual([[0]]);
  expect(worker.frame().pages[0]!.page.unbuilt).toBe(true);
  expect([...worker.changes.pages]).toEqual([1, 2]);
  expect(worker.queryFrames.some((entry) => entry.frameEpoch === 2 && entry.disposed)).toBe(true);
  expect(worker.queryFrames.slice(0, -1).every((entry) => entry.disposed)).toBe(true);
});

test('wrapped_page_boundary_probe_uses_real_forward_geometry', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const boundary of ['cold', 'warm'] as const) {
    for (const direction of ['up', 'down'] as const) {
      const worker = splitRowWorker(direction, engine, boundary);
      await worker.navigate();
      expect(worker.frame().pages[1]!.page.unbuilt).toBe(false);
      expect(worker.frame().pages[worker.neighbour]!.page.unbuilt).toBe(false);
      expect(worker.builds).toEqual([
        ...(boundary === 'cold' ? [[0]] : []), [1],
        ...(boundary === 'cold' || direction === 'down' ? [[worker.neighbour]] : []),
      ]);
      expect(worker.releases).toEqual(boundary === 'cold' ? [[0]] : []);
      expect([...worker.changes.pages]).toEqual(boundary === 'warm' && direction === 'up' ? [1] : [1, worker.neighbour]);
      const queries = createDisplayListQueries(worker.frame().displayList, engine);
      try {
        const caret = queries.caretRect(worker.position);
        expect(caret?.pageIndex).toBe(1);
        expect(caret?.height).toBeGreaterThan(0);
        expect(queries.rangeRects(worker.position, worker.position + 1)[0]?.pageIndex).toBe(1);
        expect(queries.verticalMove(worker.position, direction, 20)?.position).toBe(direction === 'up' ? 14 : 28);
      } finally {
        queries.dispose();
      }
      expect(worker.queryFrames.slice(0, -1).every((entry) => entry.disposed)).toBe(true);
      expect(worker.queryFrames.at(-1)!.disposed).toBe(false);
    }
  }
}, 30_000);

test('split_cell_probe_input_exposes_native_trailing_geometry', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  const worker = splitRowWorker('down', engine);
  const queries = createDisplayListQueries(worker.probeDisplayList(1), engine);
  try {
    expect(queries.rangeRects(worker.position, worker.position + 1)).toEqual([]);
    const trailing = queries.rangeRectsOnPages!(worker.position - 1, worker.position, 1, 1)[0];
    expect(trailing?.pageIndex).toBe(1);
    expect(trailing?.height).toBeGreaterThan(0);
  } finally {
    queries.dispose();
  }
}, 30_000);

test('split_cell_paragraph_end_probe_uses_real_trailing_geometry', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const direction of ['down', 'up'] as const) {
    const worker = splitRowWorker(direction, engine);
    await worker.navigate();
    expect(worker.builds).toEqual([[0], [1], [worker.neighbour]]);
    expect(worker.releases).toEqual([[0]]);
    expect([...worker.changes.pages]).toEqual([1, worker.neighbour]);
    const queries = createDisplayListQueries(worker.frame().displayList, engine);
    try {
      expect(queries.caretRect(worker.position)?.height).toBe(0);
      expect(queries.rangeRects(worker.position, worker.position + 1)).toEqual([]);
      expect(queries.rangeRectsOnPages!(worker.position - 1, worker.position, 1, 1)[0]?.height).toBeGreaterThan(0);
      expect(queries.verticalMove(worker.position, direction, 20)?.position).toBe(direction === 'up' ? 4 : 25);
    } finally {
      queries.dispose();
    }
    worker.builds.length = 0;
    worker.evictNeighbour();
    await worker.navigate();
    expect(worker.builds).toEqual([[worker.neighbour]]);
    expect(worker.frame().pages[1]!.page.unbuilt).toBe(false);
    expect(worker.queryFrames.slice(0, -1).every((entry) => entry.disposed)).toBe(true);
  }
}, 30_000);

test('malformed_table_metadata_navigation_stays_within_caret_and_neighbour', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const direction of ['down', 'up'] as const) {
    const worker = splitRowWorker(direction, engine, undefined, true);
    await worker.navigate();
    expect(worker.builds).toEqual([[0], [1], [worker.neighbour]]);
    expect(worker.releases).toEqual([[0]]);
    expect([...worker.changes.pages]).toEqual([1, worker.neighbour]);
    expect(worker.frame().pages[1]!.page.unbuilt).toBe(false);
    expect(worker.frame().pages[worker.neighbour]!.page.unbuilt).toBe(false);
    expect(worker.frame().pages.slice(3).every(({ page }) => page.unbuilt)).toBe(true);
    worker.builds.length = 0;
    worker.evictNeighbour();
    await worker.navigate();
    expect(worker.builds).toEqual([[worker.neighbour]]);
    expect(worker.frame().pages[1]!.page.unbuilt).toBe(false);
    expect(worker.queryFrames.slice(0, -1).every((entry) => entry.disposed)).toBe(true);
    const anonymous = splitRowWorker(direction, engine, undefined, 'anonymous');
    await anonymous.navigate();
    expect(anonymous.builds).toEqual([[0], [1]]);
    expect(anonymous.releases).toEqual([[0], [1]]);
    expect([...anonymous.changes.pages]).toEqual([]);
    expect(anonymous.frame().pages.every(({ page }) => page.unbuilt)).toBe(true);
  }
}, 30_000);

test('malformed_table_metadata_parse_failure_is_surfaced', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const direction of ['down', 'up'] as const) {
    const capture = spyOn(console, 'error').mockImplementation(() => {});
    try {
      const worker = splitRowWorker(direction, engine, undefined, true);
      const diagnostics = () => capture.mock.calls.map(([prefix, entry]) => {
        expect(prefix).toBe('[docx] display-list query parse failure');
        return entry as { error: string; queryParseFailures: number };
      });
      await worker.navigate();
      const failures = worker.queryFrames.filter((entry) => entry.failed);
      expect(failures.length).toBeGreaterThan(0);
      expect(diagnostics()).toHaveLength(failures.length);
      for (const diagnostic of diagnostics()) {
        expect(diagnostic.error).toMatch(/missing field.*rowStart/);
      }
      expect(diagnostics().map((diagnostic) => diagnostic.queryParseFailures))
        .toEqual(failures.map((_entry, index) => index + 1));
      worker.builds.length = 0;
      await worker.navigate();
      expect(worker.builds).toEqual([]);
      const nextFailures = worker.queryFrames.filter((entry) => entry.failed);
      expect(nextFailures.length).toBe(failures.length + 1);
      expect(diagnostics()).toHaveLength(nextFailures.length);
      expect(diagnostics().at(-1)?.queryParseFailures).toBe(nextFailures.length);
      for (const diagnostic of diagnostics()) {
        expect(diagnostic.error).toMatch(/missing field.*rowStart/);
      }
      await worker.navigate();
      expect(worker.builds).toEqual([]);
      expect(worker.queryFrames.filter((entry) => entry.failed)).toHaveLength(nextFailures.length);
      expect(diagnostics()).toHaveLength(nextFailures.length);
      expect(diagnostics().at(-1)?.queryParseFailures).toBe(nextFailures.length);
    } finally {
      capture.mockRestore();
    }
  }
}, 30_000);

test('later_split_cell_caret_bounds_native_and_malformed_navigation', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const malformedTable of [false, true]) {
    for (const direction of ['down', 'up'] as const) {
      const worker = splitRowWorker(direction, engine, undefined, malformedTable, { caretPage: 6 });
      await worker.navigate();
      expect(worker.builds).toEqual([[worker.caretPage], [worker.neighbour]]);
      expect(worker.releases).toEqual([]);
      expect([...worker.changes.pages]).toEqual([worker.caretPage, worker.neighbour]);
      expect(worker.frame().pages.filter(({ page }) => !page.unbuilt).map(({ pageIndex }) => pageIndex))
        .toEqual([worker.caretPage, worker.neighbour].sort((a, b) => a - b));
      worker.builds.length = 0;
      worker.evictNeighbour();
      await worker.navigate();
      expect(worker.builds).toEqual([[worker.neighbour]]);
      expect(worker.queryFrames.slice(0, -1).every((entry) => entry.disposed)).toBe(true);
      const missing = splitRowWorker(direction, engine, undefined, malformedTable, { caretPage: 6, trailingEnd: 22 });
      await missing.navigate();
      expect(missing.builds).toEqual(malformedTable ? [[missing.caretPage], [missing.caretPage + 1]]
        : [[missing.caretPage], ...Array.from({ length: 80 }, (_, page) => [page]).filter(([page]) => page !== missing.caretPage)]);
      expect(missing.releases).toEqual(missing.builds);
      expect([...missing.changes.pages]).toEqual([]);
      expect(missing.frame().pages.every(({ page }) => page.unbuilt)).toBe(true);
    }
  }
}, 30_000);

test('split_cell_missing_and_misleading_hints_only_bound_malformed_navigation', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const caretHint of ['missing', 'misleading'] as const) {
    for (const malformedTable of [false, true]) {
      for (const direction of ['down', 'up'] as const) {
        const capture = spyOn(console, 'error').mockImplementation(() => {});
        try {
          const worker = splitRowWorker(direction, engine, undefined, malformedTable, {
            caretPage: 6, caretHint, productionHooks: true,
          });
          const selection = await worker.navigate();
          expect(worker.caretHints[0]).toEqual(caretHint === 'missing' ? null : {
            pageIndex: 0, x: 0, y: 0, width: 0, height: 0,
          });
          if (malformedTable) {
            expect(worker.builds).toEqual([[0], [1]]);
            expect(worker.releases).toEqual(worker.builds);
            expect([...worker.changes.pages]).toEqual([]);
            expect(worker.frame().pages.every(({ page }) => page.unbuilt)).toBe(true);
            expect(selection?.focus).toMatchObject(worker.focus);
            expect(capture.mock.calls).toEqual([1, 2].map((queryParseFailures) => [
              '[docx] display-list query parse failure', {
                error: expect.stringMatching(/missing field.*rowStart/), queryParseFailures,
              },
            ]));
          } else {
            expect(worker.builds).toEqual([...Array.from({ length: 7 }, (_, page) => [page]), [worker.neighbour]]);
            expect(worker.releases).toEqual(Array.from({ length: 6 }, (_, page) => [page]));
            expect([...worker.changes.pages]).toEqual([worker.caretPage, worker.neighbour]);
            expect(selection?.focus).toMatchObject({ story: worker.focus.story,
              paraId: direction === 'up' ? worker.focus.paraId : 'next', offset: 0 });
            expect(capture.mock.calls).toEqual([]);
          }
          expect(worker.navigationQueryFrames.reduce((count, entry) => count + entry.verticalMoves, 0)).toBe(1);
        } finally {
          capture.mockRestore();
        }
      }
    }
  }
}, 30_000);

test('malformed_neighbour_vertical_move_parse_failure_is_surfaced_once', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const direction of ['down', 'up'] as const) {
    const verticalMove = engine.verticalMoveJson;
    const capture = spyOn(console, 'error').mockImplementation(() => {});
    try {
      const worker = splitRowWorker(direction, engine, undefined, false, {
        malformedNeighbour: true, warmCaret: true, productionHooks: true,
      });
      await worker.navigate();
      expect(engine.verticalMoveJson).toBe(verticalMove);
      expect(worker.builds).toEqual([[worker.neighbour]]);
      expect(worker.releases).toEqual([]);
      expect(worker.queryFrames.slice(0, -1).every((entry) => !entry.failed)).toBe(true);
      expect(worker.navigationQueryFrames.at(-1)?.failed).toBe(true);
      expect(worker.navigationQueryFrames.at(-1)?.verticalMoves).toBe(1);
      expect(worker.queryFrames.reduce((count, entry) => count + entry.verticalMoves, 0)).toBe(1);
      expect(capture.mock.calls).toEqual([['[docx] display-list query parse failure', {
        error: expect.stringMatching(/missing field.*rowStart/), queryParseFailures: 1,
      }]]);
      worker.builds.length = 0;
      await worker.navigate();
      await worker.navigate();
      expect(worker.builds).toEqual([]);
      expect(worker.navigationQueryFrames.filter((entry) => entry.failed)).toHaveLength(1);
      expect(capture.mock.calls).toHaveLength(1);
      worker.evictNeighbour();
      await worker.navigate();
      expect(worker.builds).toEqual([[worker.neighbour]]);
      expect(worker.navigationQueryFrames.filter((entry) => entry.failed)).toHaveLength(2);
      expect(worker.queryFrames.reduce((count, entry) => count + entry.verticalMoves, 0)).toBe(2);
      expect(capture.mock.calls).toHaveLength(2);
      expect(capture.mock.calls.at(-1)?.[1]).toEqual({
        error: expect.stringMatching(/missing field.*rowStart/), queryParseFailures: 2,
      });
    } finally {
      capture.mockRestore();
    }
  }
}, 30_000);

test('production_navigation_queries_once_per_arrow', async () => {
  const engine = await loadRustDisplayListQueryEngine();
  for (const direction of ['down', 'up'] as const) {
    const verticalMove = engine.verticalMoveJson;
    const worker = splitRowWorker(direction, engine, undefined, false, { productionHooks: true });
    for (let arrow = 1; arrow <= 3; arrow += 1) {
      await worker.navigate();
      expect(engine.verticalMoveJson).toBe(verticalMove);
      expect(worker.queryFrames.reduce((count, entry) => count + entry.verticalMoves, 0)).toBe(arrow);
      expect(worker.navigationQueryFrames.reduce((count, entry) => count + entry.verticalMoves, 0)).toBe(arrow);
    }
  }
}, 30_000);

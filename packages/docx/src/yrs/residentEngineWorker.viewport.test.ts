import { expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { DisplayPage } from '../layout/render/displayList';
import { createDisplayListQueries } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine, type RustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import type { DocxCommandFrame, ResidentCommandHooks, ResidentCommandInitRequest } from './residentCommandProtocol';

function commandFrames(options: { cell?: boolean; range?: boolean; splitEnd?: boolean; queryEngine?: RustDisplayListQueryEngine } = {}) {
  const source = readFileSync(`${import.meta.dir}/residentEngineWorker.ts`, 'utf8');
  const start = source.indexOf('\n', source.indexOf('async frame({ coversThroughSeq, full })')) + 1;
  const end = source.indexOf('\n      },\n      snapshot:', start);
  let revision = 0;
  let delta: { full: boolean; baseFrameEpoch: number; frameEpoch: number; pageCount: number;
    operations: { kind: 'upsert'; pageId: bigint; pageIndex: number; page: DisplayPage }[] };
  let caretPage = options.splitEnd ? 1 : 0;
  let offsetDelta = 0;
  let version = 0;
  let projectedVersion = -1;
  const projections: number[] = [];
  let retainBuilt = false;
  let relaidBuild = false;
  const story = options.cell ? 'body:t0:r0c0' : 'body';
  const text = (value: string, docStart: number, docEnd: number): DisplayPage['primitives'][number] => ({
    kind: 'text', text: value, x: 20, baselineY: 40, width: value.length * 10,
    font: '400 16px Arial', color: '#000000', docStart, docEnd, paraId: docEnd <= 24 ? 'paragraph' : 'next',
    table: { tableId: 't0', rowStart: 0, rowEnd: 1, rowCount: 1, columnCount: 1 }, cell: { row: 0, col: 0, rowSpan: 1, colSpan: 1, cellId: story },
    ...(options.splitEnd && docEnd <= 24 ? { fragmentDocStart: 3, fragmentDocEnd: 25 } : {}),
  });
  const primitives = (pageIndex: number): DisplayPage['primitives'] => options.splitEnd
    ? pageIndex === 0 ? [text('abcdefghij', 4, 14)] : pageIndex === 1 ? [text('qrst', 20, 24)]
      : pageIndex === 2 ? [text('next', 25, 29)] : []
    : options.cell ? [{ ...text('ab', pageIndex * 100 + 1, pageIndex * 100 + 3), paraId: 'paragraph',
      fragmentDocStart: 0, fragmentDocEnd: 2002 }] : [];
  const pages = Array.from({ length: options.splitEnd ? 80 : 20 }, (_, pageIndex) => ({ pageId: BigInt(pageIndex + 1), pageIndex,
    page: { pageIndex, width: 200, height: 100, primitives: options.splitEnd || pageIndex > 2 ? [] : primitives(pageIndex),
      unbuilt: options.splitEnd || pageIndex > 2,
      positionSpan: options.splitEnd ? [4, 44] : [pageIndex * 100, (pageIndex + 1) * 100] } as DisplayPage, primitiveIds: [] }));
  const state = {
    changes: { pages: new Set<number>(), updates: [] as Uint8Array[], viewport: undefined as { start: number; end: number; frameEpoch: number } | undefined },
    retainedFrame: { frameEpoch: 0, pages, displayList: { pages: pages.map(({ page }) => page) } },
    commandDisplayWindow: [0, 3] as [number, number], layoutDirty: false,
    pendingOffscreenPageIds: new Set<string>(), activeOffscreenPageIds: new Set<string>(),
    selectionQueries: undefined as { frameEpoch: number; queries: ReturnType<typeof createDisplayListQueries> } | undefined,
    pendingUpdates: [], glyphCache: {},
  };
  const advance = (expected: number, changed: number[] = []) => {
    delta = { full: expected === 0, baseFrameEpoch: expected, frameEpoch: ++revision, pageCount: pages.length,
      operations: (expected === 0 ? pages.map(({ pageIndex }) => pageIndex) : changed).map((pageIndex) => ({
        kind: 'upsert', pageIndex, pageId: pages[pageIndex]!.pageId, page: { ...pages[pageIndex]!.page },
      })) };
    return Uint8Array.of(revision);
  };
  const built: number[][] = [];
  const released: number[][] = [];
  const windows: number[][] = [];
  const nativeBuilds: number[] = [];
  const focus = () => ({ story, paraId: 'paragraph', offset: (options.splitEnd ? 20 : caretPage * 100 + 1) + offsetDelta });
  const engine = {
    proposalEngine: { version: () => String(version) }, retainedHeadersFootersJson: () => null,
    geometryReader: { locateParagraph: () => ({ start: 7, end: 7 + (options.splitEnd ? 20 : 2000) }) },
    setDisplayWindow: (...window: number[]) => { windows.push(window); },
    setDisplayRetainBuiltPages(retain: boolean) { retainBuilt = retain; }, applyUpdate() {},
    buildDisplayListFrame(_extras: string, expected: number) {
      const changed = !retainBuilt && relaidBuild ? [caretPage] : [];
      relaidBuild = false;
      for (const index of changed) {
        pages[index]!.page.unbuilt = false;
        pages[index]!.page.primitives = primitives(index);
        nativeBuilds.push(index);
      }
      return advance(expected, changed);
    },
    buildDisplayPagesFrame(indices: number[], expected: number) {
      built.push(indices);
      for (const index of indices) { pages[index]!.page.unbuilt = false; pages[index]!.page.primitives = primitives(index); }
      return advance(expected, indices);
    },
    releaseDisplayPagesFrame(indices: number[], expected: number) {
      released.push(indices);
      for (const index of indices) { pages[index]!.page.unbuilt = true; pages[index]!.page.primitives = []; }
      return advance(expected, indices);
    },
    residentCaretSnapshot: () => ({ frameEpoch: revision, caretRect: options.cell || pages[caretPage]!.page.unbuilt ? null
      : { pageId: String(caretPage + 1), pageIndex: caretPage, x: 0, y: 0, height: 20 } }),
  };
  const init: ResidentCommandInitRequest = { session: 'viewport', incarnation: 1, epoch: 1, readOnly: false,
    fonts: [], layoutInput: '{}', layoutWithRegions: true, layoutRevision: 1, extras: '{}', displayWindow: [0, 3] };
  const dependencies = { engine: { ...engine, partGeometry: () => [] }, fonts: { snapshot: null }, init, layoutRevision: 1, revision: 0, relayout() {},
    hooks: { selection: () => ({ anchor: { ...focus(), offset: focus().offset - (options.range ? 1 : 0) }, focus: focus() }),
      editState: () => ({ storedMarks: null }) },
    displayPositionIndex: () => ({ positionOf: (at: { offset: number }) => {
      if (projectedVersion !== version) { projectedVersion = version; projections.push(version); }
      return options.splitEnd ? 24 : at.offset + (options.cell ? 1 : 0);
    } }),
    queriesFor: () => {
      if (!options.queryEngine) return { caretRect: (position: number) => ({ pageIndex: Math.floor(position / 100),
        x: 0, y: 0, width: 0, height: pages[Math.floor(position / 100)]!.page.unbuilt ? 0 : 20 }) };
      if (state.selectionQueries?.frameEpoch !== revision) {
        state.selectionQueries?.queries.dispose();
        state.selectionQueries = { frameEpoch: revision,
          queries: createDisplayListQueries({ pages: pages.map(({ page }) => ({ ...page })) }, options.queryEngine) };
      }
      return state.selectionQueries!.queries;
    },
    decodeFrameDelta: () => delta,
    applyWorkerFrame: () => { state.retainedFrame.frameEpoch = revision; return delta; },
    async finalizeResidentFrame(bytes: Uint8Array) {
      state.retainedFrame.frameEpoch = revision;
      return { bytes, delta, caret: engine.residentCaretSnapshot() };
    },
    projectCommandEchoContext: () => undefined,
    readResidentRangeHighlights: async () => ({ rects: [[]] }),
    retainedPageByKey: (_frame: unknown, id: string) => pages[Number(id) - 1],
    offscreenDpr: 1, offscreenZoom: 1,
    stageCommandPresentation: async (input: { layout: { frame: { frameEpoch: number } }; pages: unknown[] }) => ({
      value: { frameEpoch: input.layout.frame.frameEpoch,
        pages: input.pages.map((page) => ({ page, frameEpoch: input.layout.frame.frameEpoch })) }, transfer: [],
    }),
    buildCommandFrame(input: { coversThroughSeq: number; stagedPresentation: unknown }) {
      return { value: { revision, coversThroughSeq: input.coversThroughSeq, payload: input }, transfer: [] };
    },
  };
  const body = new Bun.Transpiler({ loader: 'ts' }).transformSync(`async function frame({ coversThroughSeq, full }) {${source.slice(start, end)} }`);
  const frame = new Function('state', ...Object.keys(dependencies), `with (state) { ${body}\nreturn frame; }`)(state, ...Object.values(dependencies)) as
    (input: Parameters<ResidentCommandHooks['frame']>[0]) => Promise<{ value: DocxCommandFrame }>;
  return { state, built, released, windows, nativeBuilds, projections, engine,
    async command(pageIndex: number, relayout: boolean, offset = 0) {
      caretPage = pageIndex;
      offsetDelta = offset;
      version += 1;
      state.layoutDirty = relayout;
      relaidBuild = relayout;
      return (await frame({ coversThroughSeq: 1, version: 1, full: false })).value;
    },
    async viewport(start: number, end: number) {
      state.changes.viewport = { start, end, frameEpoch: revision };
      return (await frame({ coversThroughSeq: 0, version: 0, full: false })).value;
    },
  };
}

test('viewport_request_builds_entering_and_releases_distant_pages', async () => {
  const worker = commandFrames();
  const away = await worker.viewport(10, 13);
  expect(worker.built.flat()).toEqual(expect.arrayContaining([10, 11, 12]));
  expect(worker.released.flat()).toEqual(expect.arrayContaining([1, 2]));
  expect(worker.released.flat()).not.toContain(0);
  expect(worker.windows.at(-1)).toEqual([10, 13]);
  expect(away.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toEqual(expect.arrayContaining([0, 10, 11, 12]));
  const back = await worker.viewport(0, 3);
  expect(worker.released.flat()).toEqual(expect.arrayContaining([10, 11, 12]));
  expect(back.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toEqual([0, 1, 2]);
  const before = worker.built.length;
  const repeated = await worker.viewport(0, 3);
  expect(worker.built.slice(before)).toEqual([[]]);
  expect(repeated.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toEqual([0, 1, 2]);
  expect(repeated.coversThroughSeq).toBe(0);
  worker.state.changes.pages.add(18);
  const selection = await worker.viewport(0, 3);
  expect(worker.windows.at(-1)).toEqual([0, 3]);
  expect(selection.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(18);
  expect(worker.released.flat()).not.toContain(0);
});

test('typing_across_page_boundary_off_window_builds_and_stages_model_caret', async () => {
  const worker = commandFrames();
  await worker.viewport(10, 13);
  expect(worker.state.retainedFrame.pages[1]!.page.unbuilt).toBe(true);
  const typed = await worker.command(1, true);
  expect(typed.payload.caret.caretRect).toMatchObject({ pageIndex: 1 });
  expect(worker.nativeBuilds).toContain(1);
  expect(typed.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(1);
  expect(worker.released.at(-1)).not.toContain(1);
});

test('boundary_navigation_off_window_builds_and_stages_model_caret', async () => {
  const worker = commandFrames();
  await worker.viewport(10, 13);
  for (const pageIndex of [19, 0]) {
    expect(worker.state.retainedFrame.pages[pageIndex]!.page.unbuilt).toBe(true);
    const navigated = await worker.command(pageIndex, false);
    expect(navigated.payload.caret.caretRect).toMatchObject({ pageIndex });
    expect(worker.state.retainedFrame.pages[pageIndex]!.page.unbuilt).toBe(false);
    expect(navigated.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(pageIndex);
    expect(worker.released.at(-1)).not.toContain(pageIndex);
  }
});

test('keystrokes_with_unchanged_window_and_caret_issue_no_release_work', async () => {
  const worker = commandFrames();
  await worker.viewport(10, 13);
  const scans = spyOn(worker.state.retainedFrame.pages, 'filter');
  const released = worker.released.length;
  try {
    await worker.command(0, true);
    await worker.command(0, true);
    expect(worker.released).toHaveLength(released);
    expect(scans).not.toHaveBeenCalled();
  } finally { scans.mockRestore(); }
});

test('collapsed_table_cell_scroll_publishes_model_caret_when_native_snapshot_is_null', async () => {
  const worker = commandFrames({ cell: true });
  await worker.viewport(0, 3);
  const selected = await worker.command(4, false);
  expect(worker.engine.residentCaretSnapshot().caretRect).toBeNull();
  expect(selected.payload.caret.caretRect).toMatchObject({ pageId: '5', pageIndex: 4, height: 20 });
  const scrolled = await worker.viewport(15, 18);
  expect(worker.engine.residentCaretSnapshot().caretRect).toBeNull();
  expect(scrolled.payload.caret).toMatchObject({ frameEpoch: scrolled.revision,
    caretRect: { pageId: '5', pageIndex: 4, height: 20 } });
  expect(scrolled.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(4);
  expect(worker.state.retainedFrame.pages[4]!.page.unbuilt).toBe(false);
  expect(worker.released.at(-1)).not.toContain(4);
  const range = commandFrames({ cell: true, range: true });
  expect((await range.viewport(15, 18)).payload.caret.caretRect).toBeNull();
});

test('unchanged_window_typing_performs_no_full_story_projection_rebuild', async () => {
  for (const cell of [false, true]) {
    const worker = commandFrames({ cell });
    await worker.viewport(10, 13);
    const before = worker.projections.length;
    await worker.command(0, true, 1);
    await worker.command(0, true, 2);
    expect(worker.projections.slice(before)).toEqual([]);
  }
});

test('split_cell_paragraph_end_command_frame_bounds_builds_and_releases_misses', async () => {
  const queryEngine = await loadRustDisplayListQueryEngine();
  const worker = commandFrames({ cell: true, splitEnd: true, queryEngine });
  try {
    const scrolled = await worker.viewport(70, 73);
    expect(worker.built).toEqual([[70, 71, 72], [0], [1]]);
    expect(worker.released).toEqual([[0]]);
    expect(scrolled.payload.caret.caretRect).toMatchObject({ pageId: '2', pageIndex: 1 });
    expect(scrolled.payload.caret.caretRect!.height).toBeGreaterThan(0);
    const queries = createDisplayListQueries(worker.state.retainedFrame.displayList, queryEngine);
    try {
      expect(queries.caretRect(24)?.height).toBe(0);
      expect(queries.rangeRectsOnPages!(23, 24, 1, 1)[0]?.height).toBeGreaterThan(0);
    } finally { queries.dispose(); }
    worker.built.length = 0;
    worker.released.length = 0;
    const repeated = await worker.command(1, false);
    expect(worker.built).toEqual([]);
    expect(worker.released).toEqual([]);
    expect(repeated.payload.caret.caretRect).toMatchObject({ pageIndex: 1 });
    expect(worker.state.retainedFrame.pages.filter(({ page }) => !page.unbuilt).map(({ pageIndex }) => pageIndex))
      .toEqual([1, 70, 71, 72]);
  } finally { worker.state.selectionQueries?.queries.dispose(); }
}, 30_000);

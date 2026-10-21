import { expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { DisplayPage } from '../layout/render/displayList';
import { createDisplayListQueries } from '../layout/render/displayListQueries';
import { loadRustDisplayListQueryEngine, type RustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import type { DocxCommandFrame, DocxFrameReservations, ResidentCommandHooks, ResidentCommandInitRequest } from './residentCommandProtocol';
import { readPartGeometry } from './residentParts';
import { readCellSelection } from './residentCommandTableSelection';
import { readResidentCellSelectionRects } from './residentObjectGeometry';

function commandFrames(options: { cell?: boolean; range?: boolean; splitEnd?: boolean; queryEngine?: RustDisplayListQueryEngine; pageCount?: number } = {}) {
  const source = readFileSync(`${import.meta.dir}/residentEngineWorker.ts`, 'utf8');
  const preflight = source.indexOf('async preflight(command, context) {');
  const recordCause = new Function('changes', 'command', new Bun.Transpiler({ loader: 'ts' }).transformSync(
    source.slice(source.indexOf('\n', preflight) + 1, source.indexOf("        if (command.kind === 'selectPoint'", preflight)),
  ));
  const start = source.indexOf('\n', source.indexOf('async frame({ coversThroughSeq, full })')) + 1;
  const end = source.indexOf('\n      },\n      snapshot:', start);
  let revision = 0;
  let delta: { full: boolean; baseFrameEpoch: number; frameEpoch: number; pageCount: number; docEpoch: number; layoutEpoch: number;
    operations: { kind: 'upsert'; pageId: bigint; pageIndex: number; page: DisplayPage }[] };
  const deltas = new Map<number, typeof delta>();
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
  const pages = Array.from({ length: options.pageCount ?? (options.splitEnd ? 80 : 20) }, (_, pageIndex) => ({ pageId: BigInt(pageIndex + 1), pageIndex,
    page: { pageIndex, width: 200, height: 100, header: { kind: 'header', rId: 'rId1', y: 0, height: 10, primitives: [] }, primitives: options.splitEnd || pageIndex > 2 ? [] : primitives(pageIndex),
      unbuilt: options.splitEnd || pageIndex > 2,
      positionSpan: options.splitEnd ? [4, 44] : [pageIndex * 100, (pageIndex + 1) * 100] } as DisplayPage, primitiveIds: [] }));
  const state = {
    changes: { causes: new Set<'selection' | 'other'>(['other']), selectionFrames: [] as Uint8Array[],
      pages: new Set<number>(), updates: [] as Uint8Array[], viewport: undefined as { start: number; end: number; frameEpoch: number } | undefined },
    retainedFrame: { docEpoch: 1, layoutEpoch: 1, frameEpoch: 0, pages, displayList: { pages: pages.map(({ page }) => page) } },
    commandDisplayWindow: [0, 3] as [number, number], layoutDirty: false,
    pendingOffscreenPageIds: new Set<string>(), activeOffscreenPageIds: new Set<string>(),
    selectionQueries: undefined as { frameEpoch: number; queries: ReturnType<typeof createDisplayListQueries> } | undefined,
    pendingUpdates: [], glyphCache: {}, laidOutReservations: undefined as DocxFrameReservations | undefined,
  };
  const advance = (expected: number, changed: number[] = []) => {
    delta = { full: expected === 0, baseFrameEpoch: expected, frameEpoch: ++revision, pageCount: pages.length,
      docEpoch: 1, layoutEpoch: 1,
      operations: (expected === 0 ? pages.map(({ pageIndex }) => pageIndex) : changed).map((pageIndex) => ({
        kind: 'upsert', pageIndex, pageId: pages[pageIndex]!.pageId, page: { ...pages[pageIndex]!.page },
      })) };
    deltas.set(revision, delta);
    return Uint8Array.of(revision);
  };
  const built: number[][] = [];
  const released: number[][] = [];
  const windows: number[][] = [];
  const nativeBuilds: number[] = [];
  const focus = () => ({ story, paraId: 'paragraph', offset: (options.splitEnd ? 20 : caretPage * 100 + 1) + offsetDelta });
  const work = { wholeListBuilds: 0, partGeometryPageVisits: 0 };
  const engine = {
    proposalEngine: { version: () => String(version) }, retainedHeadersFootersJson: () => null, partGeometry: (frame: Parameters<typeof readPartGeometry>[1]) => {
      work.partGeometryPageVisits += frame.pages.length;
      return readPartGeometry({ partStory: (id: string) => `header:${id}` } as Parameters<typeof readPartGeometry>[0], frame);
    },
    geometryReader: { locateParagraph: () => ({ start: 7, end: 7 + (options.splitEnd ? 20 : 2000) }) },
    setDisplayWindow: (...window: number[]) => { windows.push(window); },
    setDisplayRetainBuiltPages(retain: boolean) { retainBuilt = retain; }, applyUpdate() { version += 1; },
    buildDisplayListFrame(_extras: string, expected: number) {
      work.wholeListBuilds += 1;
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
      const missing = indices.filter(index => pages[index]!.page.unbuilt);
      for (const index of missing) { pages[index]!.page.unbuilt = false; pages[index]!.page.primitives = primitives(index); }
      return advance(expected, missing);
    },
    releaseDisplayPagesFrame(indices: number[], expected: number) {
      released.push(indices);
      for (const index of indices) { pages[index]!.page.unbuilt = true; pages[index]!.page.primitives = []; }
      return advance(expected, indices);
    },
    cellSelection: () => null,
    residentCaretSnapshot: () => ({ frameEpoch: revision, caretRect: options.cell || pages[caretPage]!.page.unbuilt ? null
      : { pageId: String(caretPage + 1), pageIndex: caretPage, x: 0, y: 0, height: 20 } }),
  };
  const init: ResidentCommandInitRequest = { session: 'viewport', incarnation: 1, epoch: 1, readOnly: false,
    fonts: [], layoutInput: '{}', layoutWithRegions: true, layoutRevision: 1, extras: '{}', displayWindow: [0, 3] };
  const dependencies = { engine, fonts: { snapshot: null }, init, layoutRevision: 1, revision: 0, relayout() {},
    layoutPreviewVersion: 0, readResidentFrameGeometry: async () => undefined,
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
    decodeFrameDelta: (bytes: Uint8Array) => deltas.get(bytes[0]!)!,
    applyWorkerFrame: () => { state.retainedFrame.frameEpoch = revision; return delta; },
    async finalizeResidentFrame(bytes: Uint8Array) {
      state.retainedFrame.frameEpoch = revision;
      return { bytes, delta, caret: engine.residentCaretSnapshot() };
    },
    readCellSelection, readDocxSelectionFocus: () => null, readResidentCellSelectionRects,
    projectCommandEchoContext: () => undefined,
    commandEchoGlyphs: () => ({ provider: () => undefined, unsent: () => [], sent() {} }),
    readResidentRangeHighlights: async () => ({ rects: [[{ pageIndex: caretPage, x: offsetDelta, y: 0, width: 10, height: 20 }]] }),
    retainedPageByKey: (_frame: unknown, id: string) => pages[Number(id) - 1],
    offscreenDpr: 1, offscreenZoom: 1,
    stageCommandPresentation: async (input: { layout: { frame: { frameEpoch: number } }; pages: unknown[] }) => ({
      value: { frameEpoch: input.layout.frame.frameEpoch,
        pages: input.pages.map((page) => ({ page, frameEpoch: input.layout.frame.frameEpoch })) }, transfer: [],
    }),
    buildCommandFrame(input: { coversThroughSeq: number; stagedPresentation: unknown; metadata: typeof delta }) {
      return { value: { revision, baseRevision: input.metadata.full ? undefined : input.metadata.baseFrameEpoch,
        coversThroughSeq: input.coversThroughSeq, payload: input }, transfer: [] };
    },
  };
  const body = new Bun.Transpiler({ loader: 'ts' }).transformSync(`async function frame({ coversThroughSeq, full }) {${source.slice(start, end)} }`);
  const frame = new Function('state', ...Object.keys(dependencies), `with (state) { ${body}\nreturn frame; }`)(state, ...Object.values(dependencies)) as
    (input: Parameters<ResidentCommandHooks['frame']>[0]) => Promise<{ value: DocxCommandFrame }>;
  return { state, built, released, windows, nativeBuilds, projections, engine, work, get delta() { return delta; },
    queueCommand(kind: string) { recordCause(state.changes, { kind }); },
    preparePage(pageIndex: number) {
      const bytes = engine.buildDisplayPagesFrame([pageIndex], state.retainedFrame.frameEpoch);
      state.retainedFrame.frameEpoch = revision;
      state.changes.selectionFrames.push(bytes);
      state.changes.pages.add(pageIndex);
    },
    async command(pageIndex: number, relayout: boolean, offset = 0, kind = 'navigate') {
      caretPage = pageIndex;
      offsetDelta = offset;
      if (relayout) version += 1;
      recordCause(state.changes, { kind: relayout ? 'insertText' : kind });
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

test('command_frames_name_the_reservation_set_their_layout_laid_out', async () => {
  const worker = commandFrames();
  expect((await worker.command(0, true)).payload).not.toHaveProperty('reservations');
  worker.state.laidOutReservations = { version: '3', key: 'k1' };
  expect((await worker.command(0, true)).payload.reservations).toEqual({ version: '3', key: 'k1' });
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
    expect(worker.built).toEqual([[]]);
    expect(worker.released).toEqual([]);
    expect(repeated.payload.caret.caretRect).toMatchObject({ pageIndex: 1 });
    expect(worker.state.retainedFrame.pages.filter(({ page }) => !page.unbuilt).map(({ pageIndex }) => pageIndex))
      .toEqual([1, 70, 71, 72]);
  } finally { worker.state.selectionQueries?.queries.dispose(); }
}, 30_000);

test('selection_only_frames_reuse_display_pages_and_part_geometry_in_large_documents', async () => {
  for (const range of [false, true]) {
    const worker = commandFrames({ pageCount: 856, range });
    const warm = await worker.viewport(0, 3);
    const retainedPages = [...worker.state.retainedFrame.pages];
    worker.work.wholeListBuilds = 0;
    worker.work.partGeometryPageVisits = 0;
    worker.built.length = 0;
    worker.queueCommand('selectPoint');
    worker.queueCommand('extendSelection');
    const selected = await worker.command(0, false, 1);
    expect(worker.work).toEqual({ wholeListBuilds: 0, partGeometryPageVisits: 0 });
    expect(worker.built).toEqual([[]]);
    expect(worker.delta.operations).toEqual([]);
    expect(selected.payload.parts).toBe(warm.payload.parts);
    worker.state.retainedFrame.pages.forEach((page, index) => expect(page).toBe(retainedPages[index]));
    expect(selected.revision).toBeGreaterThan(warm.revision);
    expect(selected.payload.caret.frameEpoch).toBe(selected.revision);
    expect(selected.payload.selectionRects).toEqual(range ? [{ pageIndex: 0, x: 1, y: 0, width: 10, height: 20 }] : []);
    const edited = await worker.command(0, true, 2);
    expect(worker.work.wholeListBuilds).toBe(1);
    expect(worker.work.partGeometryPageVisits).toBe(856);
    expect(edited.payload.parts).toEqual(warm.payload.parts);
  }
});

test('selection_only_boundary_move_materializes_destination_and_updates_its_parts', async () => {
  const worker = commandFrames({ pageCount: 856 });
  await worker.viewport(0, 3);
  worker.work.wholeListBuilds = 0;
  worker.work.partGeometryPageVisits = 0;
  worker.built.length = 0;
  const selected = await worker.command(3, false);
  expect(worker.work.wholeListBuilds).toBe(0);
  expect(worker.work.partGeometryPageVisits).toBe(1);
  expect(worker.built).toEqual([[], [3]]);
  expect(selected.payload.caret.caretRect).toMatchObject({ pageIndex: 3 });
  expect(selected.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(3);
  expect(selected.payload.parts).toEqual(readPartGeometry(
    { partStory: (id: string) => `header:${id}` } as Parameters<typeof readPartGeometry>[0],
    worker.state.retainedFrame as unknown as Parameters<typeof readPartGeometry>[1],
  ));
});

test('coalesced_nonselection_causes_keep_full_work_even_when_selection_is_last', async () => {
  const worker = commandFrames({ pageCount: 856 });
  await worker.viewport(0, 3);
  worker.work.wholeListBuilds = 0;
  worker.work.partGeometryPageVisits = 0;
  for (const kind of ['insertText', 'applyRemoteUpdate', 'setRenderEnvironment', 'setFontEnvironment', 'futureCommand']) {
    worker.queueCommand(kind);
    await worker.command(0, false, 1);
    expect(worker.work).toEqual({ wholeListBuilds: 1, partGeometryPageVisits: 856 });
    worker.work.wholeListBuilds = 0;
    worker.work.partGeometryPageVisits = 0;
  }
});

test('selection_boundary_move_before_first_viewport_materializes_destination', async () => {
  const worker = commandFrames({ pageCount: 856 });
  await worker.command(0, true);
  worker.work.wholeListBuilds = 0;
  worker.work.partGeometryPageVisits = 0;
  worker.built.length = 0;
  const selected = await worker.command(3, false);
  expect(worker.work).toEqual({ wholeListBuilds: 0, partGeometryPageVisits: 1 });
  expect(worker.built).toEqual([[], [3]]);
  expect(selected.payload.caret.caretRect).toMatchObject({ pageIndex: 3 });
  expect(selected.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(3);
});

test('remote_update_invalidates_geometry_and_the_next_selection_reuses_it', async () => {
  const worker = commandFrames({ pageCount: 856 });
  await worker.viewport(0, 3);
  worker.work.wholeListBuilds = 0;
  worker.work.partGeometryPageVisits = 0;
  worker.state.changes.updates.push(Uint8Array.of(1));
  await worker.command(0, false, 1);
  expect(worker.work).toEqual({ wholeListBuilds: 1, partGeometryPageVisits: 856 });
  worker.work.wholeListBuilds = 0;
  worker.work.partGeometryPageVisits = 0;
  await worker.command(0, false, 2);
  expect(worker.work).toEqual({ wholeListBuilds: 0, partGeometryPageVisits: 0 });
});

test('selection_preflight_page_builds_publish_incrementally_from_the_previous_frame', async () => {
  const worker = commandFrames({ pageCount: 856 });
  const warm = await worker.viewport(0, 3);
  worker.work.wholeListBuilds = 0;
  worker.work.partGeometryPageVisits = 0;
  worker.queueCommand('navigate');
  worker.preparePage(3);
  worker.preparePage(4);
  const selected = await worker.command(4, false);
  expect(worker.work).toEqual({ wholeListBuilds: 0, partGeometryPageVisits: 2 });
  expect(selected.baseRevision).toBe(warm.revision);
  expect(selected.payload.pageFrames).toHaveLength(2);
  expect(selected.payload.caret.caretRect).toMatchObject({ pageIndex: 4 });
  expect(selected.payload.stagedPresentation?.pages.map(({ page }) => page.pageIndex)).toContain(4);
  expect(selected.payload.parts).toEqual(readPartGeometry(
    { partStory: (id: string) => `header:${id}` } as Parameters<typeof readPartGeometry>[0],
    worker.state.retainedFrame as unknown as Parameters<typeof readPartGeometry>[1],
  ));
});

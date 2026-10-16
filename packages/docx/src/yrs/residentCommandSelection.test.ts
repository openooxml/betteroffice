import { afterEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as queryModule from '../layout/render/displayListQueries';
import * as rustQueries from '../layout/render/rustDisplayList';
import type { DisplayPage, TextRunPrimitive } from '../layout/render/displayList';
import { applyFrameDelta, decodeFrameDelta, retainedFramePageById, FRAME_DELTA_VERSION, type RetainedFrame } from '../layout/render/frameDelta';
import { loadRustDisplayListQueryEngine, loadedRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { DisplayPositionIndex } from './displayPositionIndex';
import { displayPositionToYrsLoc } from './inputPositionMap';
import { YrsPositionProjection } from './yrsPositionProjection';
import { readResidentFrameGeometry, resetResidentObjectHandles } from './residentObjectGeometry';
import type { YrsLoc, YrsStoryRange, YrsStorySegment } from './index';
import { mintCommandSelection } from './residentCommandAnchors';
import { ResidentSessionFonts } from './residentFontEnvironment';
import { createProposalRegistry, proposalRevisionPreview } from './proposals';
import { applyDocxCommand, preflightDocxCommand } from './residentCommandActions';
import { releaseResidentSelectionReads, resolveDocxSelectionIntent, type DocxSelectionCommandIntent } from './residentCommandSelection';
import type { DocxCommandEnvelope, DocxCommandSelection, DocxNavigationIntent, ResidentCommandHooks, ResidentCommandHooksFactory, ResidentCommandHostOptions, ResidentCommandInitRequest, ResidentDocxCommandEngine } from './residentCommandProtocol';
import { TIMEOUT, closeCommandEngines, command, context, execute, openCommandEngine as openEngine, selection, texts, type CommandArm } from './residentCommandTestSupport';

const selectionEngines = new Set<ResidentDocxCommandEngine>();
const releasesAfterTest: (() => void)[] = [];

afterEach(() => {
  for (const release of releasesAfterTest.splice(0)) release();
  for (const engine of selectionEngines) releaseResidentSelectionReads(engine);
  selectionEngines.clear();
  closeCommandEngines();
});

async function openCommandEngine(...args: Parameters<typeof openEngine>): Promise<CommandArm> {
  const arm = await openEngine(...args);
  selectionEngines.add(arm.engine);
  return arm;
}

test('autofit_cell_click_preserves_dependent_tab_navigation_and_typing', async () => {
  await loadRustDisplayListQueryEngine();
  const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const arm = await openCommandEngine([], { editState: true, bodyXml:
    '<w:p/><w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>`,
  });
  const readLayout = () => layout(arm, true);
  const initial = readLayout();
  const page = initial.frame.pages[0]!;
  const run = page.page.primitives.find((primitive): primitive is TextRunPrimitive =>
    primitive.kind === 'text' && primitive.cell?.row === 0 && primitive.cell.col === 0
  );
  if (!run) throw new Error('First cell has no rendered text');
  const clip = run.clipGroup?.clip;
  if (clip?.x === undefined || clip.w === undefined) throw new Error('First cell has no clip bounds');
  expect(clip.x + clip.w).toBeLessThan(120);
  const clicked = ordered(arm, command(arm, 'selectPoint', {
    page: { pageId: page.pageId.toString(), pageIndex: 0 }, x: run.x + run.width / 2,
    y: run.baselineY - 5, layoutRevision: initial.layoutRevision, frameEpoch: initial.frame.frameEpoch,
    clickCount: 1, modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  }), null, readLayout);
  expect(clicked.editState?.tableContext).toMatchObject({ rowIndex: 0, columnIndex: 0 });
  if (!clicked.selection) throw new Error('Cell click has no selection');
  const firstCell = clicked.selection.focus.story;
  ordered(arm, command(arm, 'navigate', { unit: 'line', direction: 'backward', extend: false }), clicked.selection ?? null, readLayout);
  const basis = { epoch: 1, version: arm.version };
  const insert = command(arm, 'insertText', { text: 'X' }, { basis });
  const forward = command(arm, 'navigateCell', { direction: 'forward', extend: false, createLastRow: true }, { basis, afterSeq: insert.seq });
  const next = command(arm, 'insertText', { text: 'Y' }, { basis, afterSeq: forward.seq });
  const backward = command(arm, 'navigateCell', { direction: 'backward', extend: false, createLastRow: true }, { basis, afterSeq: next.seq });
  const last = command(arm, 'insertText', { text: 'Z' }, { basis, afterSeq: backward.seq });
  let previous = arm.state.selection;
  for (const intent of [insert, forward, next, backward, last]) {
    const result = execute(arm, intent, previous);
    expect(result.status).not.toBe('refused');
    previous = result.selection ?? null;
  }
  expect(arm.engine.paragraphs(firstCell)[0]?.text).toBe('ZXA');
  expect(arm.engine.paragraphs(firstCell.replace(/:r0c0$/, ':r0c1'))[0]?.text).toBe('YB');
}, TIMEOUT);

function layout(arm: CommandArm, authoritativeShaping = false) {
  arm.engine.layoutDocumentWithRegionsRetained(JSON.stringify({
    bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' },
      ...(authoritativeShaping ? { authoritativeShaping, fontChains: { 'calibri|0|0': [0], 'liberation sans|0|0': [0] } } : {}) }, renderEnv: {},
  }));
  const frame = applyFrameDelta(null, decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0)));
  return { frame, layoutRevision: arm.version };
}

function ordered(arm: CommandArm, intent: DocxCommandEnvelope, predecessor: DocxCommandSelection | null, readSelectionLayout = () => layout(arm)) {
  const ctx = context(arm, intent, predecessor);
  const checked = preflightDocxCommand(arm.engine, arm.state, intent, ctx, { readSelectionLayout });
  if (checked.status === 'refused') return checked;
  const result = applyDocxCommand(arm.engine, arm.state, checked.prepared, ctx);
  if (result.status === 'applied') arm.version++;
  return result;
}

function fakeArm(stories: Record<string, string[]>): CommandArm {
  const id = (story: string, index: number) => `${story}-${index}`;
  const ordinal = (story: string, paraId: string) => stories[story]!.findIndex((_, index) => id(story, index) === paraId);
  const sticky = new Map<number, YrsLoc>();
  const engine = {
    geometryReader: {
      version: () => '0',
      hasStory: (story: string) => story in stories,
      storyIds: () => Object.keys(stories),
      paragraphSpans: (story: string) => stories[story]!.map((text, index) => ({ paraId: id(story, index), length: text.length })),
      paragraphIdCount: (story: string, paraId: string) => ordinal(story, paraId) < 0 ? 0 : 1,
      locateParagraph: (story: string, paraId: string) => {
        const index = ordinal(story, paraId);
        const start = stories[story]!.slice(0, index).reduce((sum, text) => sum + text.length + 1, 0);
        return { start, end: start + stories[story]![index]!.length };
      },
      storySegments: (story: string): YrsStorySegment[] => stories[story]!.flatMap<YrsStorySegment>((text, index) => [
        { kind: 'text' as const, text, attributes: {} },
        { kind: 'pilcrow' as const, paraId: id(story, index), properties: {}, attributes: {} },
      ]),
    },
    paragraphs: () => { throw new Error('Whole-story text read'); },
    paragraphSegments: (story: string, indices: readonly number[]): YrsStorySegment[][] => indices.map<YrsStorySegment[]>((index) => [
      { kind: 'text', text: stories[story]![index]!, attributes: {} },
      { kind: 'pilcrow', paraId: id(story, index), properties: {}, attributes: {} },
    ]),
    selectionText: (range: YrsStoryRange) => {
      const paragraphText = stories[range.story]![ordinal(range.story, range.start.paraId)]!;
      return {
        paraId: range.start.paraId, paragraphText,
        selectedText: paragraphText.slice(range.start.offset, range.end.offset),
        before: paragraphText.slice(0, range.start.offset), after: paragraphText.slice(range.end.offset),
      };
    },
    encodeStickyPosition: (at: YrsLoc) => {
      const key = sticky.size + 1;
      sticky.set(key, at);
      return { story: at.story, encoded: Uint8Array.of(key) };
    },
    resolveStickyPosition: (at: { encoded: Uint8Array }) => sticky.get(at.encoded[0]!) ?? null,
  } as unknown as ResidentDocxCommandEngine;
  selectionEngines.add(engine);
  return {
    engine, native: {} as CommandArm['native'], version: 0, seq: 0,
    state: { readOnly: false, selection: selection('body-0', 0), storedMarks: null, stickyPositions: new Map() },
  };
}

function fakeFrame(pages: DisplayPage[], frameEpoch = 1): RetainedFrame {
  return {
    protocolVersion: FRAME_DELTA_VERSION, docEpoch: 1, layoutEpoch: 1, frameEpoch,
    pages: pages.map((page) => ({ page, pageIndex: page.pageIndex, pageId: BigInt(page.pageIndex + 1), fingerprint: 0n, primitiveIds: new BigUint64Array() })),
    damagedPageIds: new Set(), removedPageIds: new Set(), displayList: { pages },
  };
}

function runNavigation(arm: CommandArm, frame: RetainedFrame, navigation: DocxNavigationIntent) {
  const intent = command(arm, 'navigate', navigation) as DocxSelectionCommandIntent;
  return resolveDocxSelectionIntent({
    engine: arm.engine, state: arm.state, command: intent, context: context(arm, intent),
    readLayout: () => ({ frame, layoutRevision: 0 }),
  });
}

function fakeLine(from: number, to: number, pageIndex = 0): queryModule.DisplayListVisualLine {
  return { from, to, pageIndex, x: 0, y: 0, width: 60, height: 10, baseline: 8 };
}

function fakeQueries(overrides: Partial<queryModule.DisplayListQueries>): queryModule.DisplayListQueries {
  const unused = () => { throw new Error('Unexpected display-list query'); };
  return {
    displayList: { pages: [] }, isReady: () => true, sourceState: () => ({ status: 'ready' }),
    whenReady: async () => {}, dispose: () => {}, prime: unused,
    pageCount: unused, pageSize: unused, pageBounds: unused, contentBounds: unused, columnBounds: unused,
    paragraphRects: unused, visualLines: unused, visualLinesOnPage: unused,
    visualLineExtent: unused, visualLineAtPosition: unused, imageAtPoint: unused, imageByPos: unused,
    hitTestRegions: unused, verticalMove: unused, rangeRects: unused, hfRangeRects: unused,
    noteRangeRects: unused, hfCaretRects: unused, noteCaretRects: unused, hfAnchorRects: unused,
    caretRect: unused, anchorRect: unused, sidebarAnchorRect: unused,
    ...overrides,
  };
}

test('d2a_query_fakes_reject_unconfigured_geometry_reads', async () => {
  const queries: queryModule.DisplayListQueries = fakeQueries({
    verticalMove: (position) => ({ position: position + 1, goalX: 20 }),
  });
  expect(queries.sourceState()).toEqual({ status: 'ready' });
  await expect(queries.whenReady()).resolves.toBeUndefined();
  expect(queries.verticalMove(2, 'down', 20)).toEqual({ position: 3, goalX: 20 });
  expect(() => queries.caretRect(2)).toThrow('Unexpected display-list query');
  expect(() => queries.hitTestRegions(0, 20, 40)).toThrow('Unexpected display-list query');
});

test('d2a_worker_test_loader_bundles_selection_dependencies', async () => {
  const { residentWorkerFactory } = await import('./__fixtures__/residentWorker');
  const startWorker = await residentWorkerFactory();
  const worker = startWorker();
  try {
    expect(worker.requests).toEqual([]);
    expect(worker.sessions).toEqual([]);
  } finally {
    worker.terminate();
  }
}, TIMEOUT);

test('d2a_worker_read_only_initial_frame_does_not_load_query_wasm', async () => {
  const { preloadEditWasm } = await import('../wasm/edit');
  const { preloadOpcWasm } = await import('../wasm/opc');
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  await preloadOpcWasm(new Uint8Array(readFileSync(new URL('../wasm/generated/opc/ooxml_opc_bg.wasm', import.meta.url))));
  const { rezipPartsToArrayBuffer, toBytes } = await import('../docx/rezip/parts');
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>')],
  ]);
  const { residentWorkerFactory } = await import('./__fixtures__/residentWorker');
  const { ResidentEngineWorkerClient } = await import('./residentEngineWorkerClient');
  const startWorker = await residentWorkerFactory();
  const queryLoad = spyOn(rustQueries, 'loadRustDisplayListQueryEngine').mockImplementation(() => {
    throw new Error('Unexpected query wasm load');
  });
  const worker = startWorker(8103);
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.open(new Uint8Array(rezipPartsToArrayBuffer(parts)));
    const init: ResidentCommandInitRequest = {
      session: 'd2a-viewer', incarnation: 1, epoch: 1, readOnly: true,
      fonts: [new Uint8Array(readFileSync(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)))],
      layoutInput: JSON.stringify({
        bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
        measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } }, renderEnv: {},
      }),
      layoutWithRegions: true, layoutRevision: 1, extras: '{}',
    };
    const initial = await client.commandInit(init);
    expect(initial.editState.readOnly).toBe(true);
    expect(initial.frame.coversThroughSeq).toBe(0);
    const frame = applyFrameDelta(null, decodeFrameDelta(initial.frame.payload.frame));
    expect(frame.displayList.pages.flatMap((page) => page.primitives).map((primitive) =>
      primitive.kind === 'text' || primitive.kind === 'glyphRun' ? primitive.text : ''
    ).join('')).toBe('Seed');
    expect(worker.sessions).toHaveLength(1);
    expect(worker.requests).toEqual(['open', 'commandInit']);
    expect(queryLoad).not.toHaveBeenCalled();
  } finally {
    client.destroy();
    queryLoad.mockRestore();
  }
}, TIMEOUT);

function workerSelectionHooks(arm: CommandArm, frame: RetainedFrame, buildPages?: (pages: number[], frameEpoch: number) => RetainedFrame) {
  const index = new DisplayPositionIndex({ ...arm.engine.geometryReader, selectionText: arm.engine.selectionText, resolveComment: arm.engine.resolveComment });
  const factory: ResidentCommandHooksFactory = (_engine, _init, environment) => ({
    preflight: (intent, ctx) => preflightDocxCommand(arm.engine, arm.state, intent, ctx, environment),
    apply: (prepared, ctx) => applyDocxCommand(arm.engine, arm.state, prepared, ctx),
    selection: () => arm.state.selection,
    editState: () => { throw new Error('Unused edit-state projection'); },
  });
  const source = readFileSync(new URL('./residentEngineWorker.ts', import.meta.url), 'utf8');
  const binding = source.slice(source.indexOf('function bindResidentCommands('), source.indexOf('function replacementCommandInit('));
  const dependencies = {
    session: {
      ...arm.engine, proposalEngine: { version: () => arm.engine.geometryReader.version() },
      selection: () => ({ anchor: arm.state.selection!.anchor, head: arm.state.selection!.focus }),
      buildDisplayPagesFrame: buildPages ?? (() => { throw new Error('Unexpected page build'); }),
    },
    commandHooksFactory: factory,
    collaboration: null,
    init: { layoutWithRegions: true, layoutInput: '{}', extras: '{}' } as ResidentCommandInitRequest,
    retainedFrame: frame, retainedLayoutVersion: null, layoutRevision: 0, commandFrameChanges: null,
    commandInit: null, commandHost: null, releaseCommandHost: null, releaseResidentSelectionReads,
    readResidentFrameGeometry, resetResidentObjectHandles, proposals: null, createProposalRegistry, proposalRevisionPreview,
    sessionFonts: null, ResidentSessionFonts, scope: { fonts: undefined },
    createResidentCommandHost: (options: ResidentCommandHostOptions) => options.hooks,
    displayPositionIndex: () => index, createDisplayListQueries: queryModule.createDisplayListQueries,
    loadRustDisplayListQueryEngine: async () => {}, loadedRustDisplayListQueryEngine: () => undefined,
    applyFrameDelta: (_frame: RetainedFrame, next: RetainedFrame) => next,
    decodeFrameDelta: (next: RetainedFrame) => next,
    retainedPageByKey: (current: RetainedFrame, pageId: string) => retainedFramePageById(current, BigInt(pageId)),
  };
  const worker = new Function(...Object.keys(dependencies),
    new Bun.Transpiler({ loader: 'ts' }).transformSync(binding) +
    '\nreturn { hooks: bindResidentCommands(init), release: () => { releaseCommandHost?.(); releaseCommandHost = null; }, install: (frame, revision) => { retainedFrame = frame; layoutRevision = revision; } };'
  )(...Object.values(dependencies)) as { hooks: ResidentCommandHooks; release(): void; install: (frame: RetainedFrame, revision: number) => void };
  releasesAfterTest.push(worker.release);
  return worker;
}

test('d2a_navigation_resolves_from_predecessor_selection', async () => {
  expect(() => resolveDocxSelectionIntent({} as Parameters<typeof resolveDocxSelectionIntent>[0])).not.toThrow('not implemented');
  const arm = await openCommandEngine();
  const id = arm.engine.paragraphs('body')[0]!.paraId;
  const second = arm.engine.paragraphs('body')[1]!.paraId;
  const inserted = execute(arm, command(arm, 'insertText', { text: '!', selection: selection(id, 5) }));
  expect(inserted.status).toBe('applied');
  const predecessor = inserted.selection!;
  execute(arm, command(arm, 'select', { selection: selection(second, 0) }));
  const source = layout(arm);
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockReturnValue(fakeQueries({
    isReady: () => true,
    dispose: () => {},
    verticalMove: (position: number, direction: 'up' | 'down', goalX?: number) => {
      expect(position).toBe(7);
      expect(direction).toBe('down');
      expect(goalX).toBe(20);
      return { position: 17, goalX: 20 };
    },
    caretRect: () => ({ pageIndex: 0, x: 0, y: 0, width: 0, height: 10 }),
    visualLinesOnPage: () => [fakeLine(1, 12)],
  }));
  try {
    const cases: [DocxNavigationIntent, string, number][] = [
      [{ unit: 'character', direction: 'backward', extend: false }, id, 5],
      [{ unit: 'word', direction: 'forward', extend: false }, id, 7],
      [{ unit: 'line', direction: 'backward', extend: false }, id, 0],
      [{ unit: 'line', direction: 'forward', extend: false }, id, 11],
      [{ unit: 'line', direction: 'down', extend: true, goalX: 20 }, second, 3],
      [{ unit: 'document-boundary', direction: 'forward', extend: false }, second, 6],
    ];
    for (const [navigation, paraId, offset] of cases) {
      const intent = command(arm, 'navigate', navigation, { afterSeq: 1, basis: { epoch: 1, version: 0 } });
      const resolved = resolveDocxSelectionIntent({ engine: arm.engine, state: arm.state, command: intent as DocxSelectionCommandIntent, context: context(arm, intent, predecessor), readLayout: () => source });
      expect(resolved.selection.focus).toMatchObject({ story: 'body', paraId, offset });
      if (navigation.extend) expect(resolved.selection.anchor).toEqual(predecessor.anchor);
      const result = ordered(arm, intent, predecessor);
      expect(result.status).toBe('noop');
      expect(result.selection?.focus).toMatchObject({ paraId, offset });
      expect(result.selection?.focus.sticky).toBeInstanceOf(Uint8Array);
    }
  } finally {
    queries.mockRestore();
  }
}, TIMEOUT);

test('d2a_vertical_extension_clamps_at_body_cell_story_boundary', () => {
  const cell = 'body:t0:r0c0';
  const arm = fakeArm({ body: ['Start', 'Before table'], [cell]: ['Cell text'] });
  const segments = arm.engine.geometryReader.storySegments;
  arm.engine.geometryReader.storySegments = (story) => story === 'body' ? [
    ...segments(story),
    { kind: 'embed', embedKind: 'table', attributes: {}, payload: { rows: [{ cells: [{ story: cell }] }] } },
  ] : segments(story);
  const index = new DisplayPositionIndex({
    ...arm.engine.geometryReader, selectionText: arm.engine.selectionText, resolveComment: arm.engine.resolveComment,
  });
  const frame = fakeFrame([{ pageIndex: 0, width: 600, height: 800, primitives: [] }]);
  let target = 0;
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockReturnValue(fakeQueries({
    isReady: () => true, dispose: () => {}, verticalMove: () => ({ position: target, goalX: 25 }),
  }));
  try {
    for (const [story, paraId, destination, direction] of [
      ['body', 'body-1', { story: cell, paraId: `${cell}-0`, offset: 3 }, 'down'],
      [cell, `${cell}-0`, { story: 'body', paraId: 'body-1', offset: 3 }, 'up'],
    ] as const) {
      const selected = mintCommandSelection(arm.engine, arm.state, selection(paraId, 1, 3, story));
      arm.state.selection = selected;
      target = index.positionOf(destination, 'body')!;
      expect(target).toBeGreaterThan(0);
      const resolved = runNavigation(arm, frame, { unit: 'line', direction, extend: true, goalX: 25 });
      expect(resolved.status).toBe('clamped');
      expect(resolved.selection).toEqual(selected);
      expect(resolved.range).toEqual({ story, start: { paraId, offset: 1 }, end: { paraId, offset: 3 } });
    }
  } finally {
    queries.mockRestore();
  }
});

test('d2a_unbuilt_page_hit_preserves_selection_until_geometry_is_available', () => {
  const arm = fakeArm({ body: ['Start', 'Target'] });
  arm.state.selection = selection('body-1', 2, 4);
  const selected = arm.state.selection;
  const page: DisplayPage = { pageIndex: 0, width: 600, height: 800, primitives: [], unbuilt: true };
  const intent = command(arm, 'selectPoint', {
    page: { pageId: '1', pageIndex: 0 }, x: 50, y: 80, layoutRevision: 0, frameEpoch: 1, clickCount: 2,
    modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  }) as DocxSelectionCommandIntent;
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockImplementation((list) => fakeQueries({
    isReady: () => true, dispose: () => {},
    hitTestRegions: () => list.pages[0]!.unbuilt ? null : { region: 'body', pos: 11, target: 'text' },
    visualLinesOnPage: () => [],
  }));
  try {
    const read = (frame: RetainedFrame) => resolveDocxSelectionIntent({
      engine: arm.engine, state: arm.state, command: intent, context: context(arm, intent),
      readLayout: () => ({ frame, layoutRevision: 0 }),
    });
    const pending = read(fakeFrame([page]));
    expect(pending.status).toBe('clamped');
    expect(pending.selection).toEqual(selected);
    expect(queries).not.toHaveBeenCalled();
    const resolved = read(fakeFrame([{ ...page, unbuilt: false }], 2));
    expect(resolved.status).toBe('resolved');
    expect(resolved.selection).toEqual(selection('body-1', 0, 6));
    expect(queries).toHaveBeenCalledTimes(1);
  } finally {
    queries.mockRestore();
  }
});

test('d2a_region_navigation_uses_header_footer_and_note_geometry', () => {
  const arm = fakeArm({ body: ['Body'], 'hf:head': ['Header'], 'hf:foot': ['Footer'], 'fn:7': ['Note'], 'en:8': ['Endnote'] });
  const run = (paraId: string, text: string, groupId?: string): TextRunPrimitive => ({
    kind: 'text', text, paraId, docStart: 1, docEnd: text.length + 1,
    x: 20, baselineY: 40, width: 60, font: '11px sans-serif', color: '#000000', groupId,
  });
  const body = run('body-0', 'Body');
  const header = run('hf:head-0', 'Header');
  const footer = run('hf:foot-0', 'Footer');
  const note = run('fn:7-0', 'Note', 'footnote-7');
  const endnote = run('en:8-0', 'Endnote', 'endnote-8');
  const otherNote = run('fn:9-0', 'Other', 'footnote-9');
  const frame = fakeFrame([{
    pageIndex: 0, width: 600, height: 800, primitives: [body],
    header: { kind: 'header', rId: 'head', y: 0, height: 50, primitives: [header] },
    footer: { kind: 'footer', rId: 'foot', y: 750, height: 50, primitives: [footer] },
    noteAreas: [
      { kind: 'footnote', noteIds: [7, 9], primitives: [note, otherNote] },
      { kind: 'endnote', noteIds: [8], primitives: [endnote] },
    ],
  }]);
  let active = header;
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockImplementation((list) => {
    expect(list.pages[0]!.primitives).toEqual([active]);
    expect(list.pages[0]!.header).toBeUndefined();
    expect(list.pages[0]!.footer).toBeUndefined();
    expect(list.pages[0]!.noteAreas).toBeUndefined();
    return fakeQueries({
      isReady: () => true, dispose: () => {},
      verticalMove: (position: number) => {
        expect(position).toBe(3);
        return { position: 4, goalX: 20 };
      },
      caretRect: () => ({ pageIndex: 0, x: 0, y: 0, width: 0, height: 10 }),
      visualLinesOnPage: () => [fakeLine(1, active.text.length + 1)],
    });
  });
  try {
    for (const [story, primitive] of [['hf:head', header], ['hf:foot', footer], ['fn:7', note], ['en:8', endnote]] as const) {
      active = primitive;
      arm.state.selection = selection(primitive.paraId!, 2, 2, story);
      const moved = runNavigation(arm, frame, { unit: 'line', direction: 'down', extend: true, goalX: 20 });
      expect(moved.selection).toEqual(selection(primitive.paraId!, 2, 3, story));
      const home = runNavigation(arm, frame, { unit: 'line', direction: 'backward', extend: false });
      expect(home.selection).toEqual(selection(primitive.paraId!, 0, 0, story));
      const end = runNavigation(arm, frame, { unit: 'line', direction: 'forward', extend: false });
      expect(end.selection).toEqual(selection(primitive.paraId!, primitive.text.length, primitive.text.length, story));
    }
  } finally {
    queries.mockRestore();
  }
});

test('d2a_character_and_word_navigation_reads_only_the_active_paragraph', () => {
  const arm = fakeArm({ body: ['Start', 'Alpha 😀 beta', 'Last'] });
  const frame = fakeFrame([]);
  const spans = spyOn(arm.engine.geometryReader, 'paragraphSpans');
  const text = spyOn(arm.engine, 'paragraphSegments');
  const accepted = spyOn(arm.engine, 'selectionText');
  try {
    arm.state.selection = selection('body-1', 6);
    const forward = runNavigation(arm, frame, { unit: 'character', direction: 'forward', extend: false });
    expect(forward.selection).toEqual(selection('body-1', 8));
    arm.state.selection = forward.selection;
    const backward = runNavigation(arm, frame, { unit: 'character', direction: 'backward', extend: false });
    expect(backward.selection).toEqual(selection('body-1', 6));
    arm.state.selection = backward.selection;
    const word = runNavigation(arm, frame, { unit: 'word', direction: 'forward', extend: false });
    expect(word.selection).toEqual(selection('body-1', 9));
    expect(spans).toHaveBeenCalledTimes(1);
    expect(text).toHaveBeenCalledTimes(3);
    expect(text.mock.calls.every(([story, indices]) => story === 'body' && indices.length === 1 && indices[0] === 1)).toBe(true);
    expect(accepted).not.toHaveBeenCalled();
  } finally {
    spans.mockRestore();
    text.mockRestore();
    accepted.mockRestore();
  }
});

test('d2a_line_navigation_reuses_caches_and_reads_lines_only_on_the_caret_page', () => {
  const arm = fakeArm({ body: ['Start', 'Active'] });
  arm.state.selection = selection('body-1', 2);
  const frame = fakeFrame(Array.from({ length: 30 }, (_, pageIndex) => ({ pageIndex, width: 600, height: 800, primitives: [] })));
  const segments = spyOn(arm.engine.geometryReader, 'storySegments');
  let disposed = 0;
  let lineReads = 0;
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockReturnValue(fakeQueries({
    isReady: () => true,
    dispose: () => { disposed++; },
    caretRect: () => ({ pageIndex: 12, x: 0, y: 0, width: 0, height: 10 }),
    visualLinesOnPage: (pageIndex: number) => {
      expect(pageIndex).toBe(12);
      lineReads++;
      return [fakeLine(8, 14, 12)];
    },
    visualLineAtPosition: () => { throw new Error('Document-wide line scan'); },
  }));
  try {
    for (const direction of ['backward', 'forward', 'backward', 'forward'] as const) {
      const resolved = runNavigation(arm, frame, { unit: 'line', direction, extend: false });
      expect(resolved.selection).toEqual(selection('body-1', direction === 'backward' ? 0 : 6));
    }
    expect(queries).toHaveBeenCalledTimes(1);
    expect(segments).toHaveBeenCalledTimes(1);
    expect(lineReads).toBe(4);
    expect(disposed).toBe(0);
    runNavigation(arm, { ...frame, frameEpoch: 2 }, { unit: 'line', direction: 'forward', extend: false });
    expect(queries).toHaveBeenCalledTimes(2);
    expect(segments).toHaveBeenCalledTimes(1);
    expect(disposed).toBe(1);
    arm.engine.geometryReader.version = () => '1';
    runNavigation(arm, frame, { unit: 'line', direction: 'forward', extend: false });
    expect(queries).toHaveBeenCalledTimes(3);
    expect(segments).toHaveBeenCalledTimes(2);
    expect(disposed).toBe(2);
  } finally {
    queries.mockRestore();
    segments.mockRestore();
  }
});

test('d2a_click_on_deleted_page_then_typing_lands_at_clamped_caret', async () => {
  expect(() => resolveDocxSelectionIntent({} as Parameters<typeof resolveDocxSelectionIntent>[0])).not.toThrow('not implemented');
  const arm = await openCommandEngine(['Keep', ...Array.from({ length: 100 }, () => 'Deleted paragraph')]);
  const before = layout(arm);
  expect(before.frame.pages.length).toBeGreaterThan(1);
  const page = before.frame.pages.at(-1)!;
  const paragraphs = arm.engine.paragraphs('body');
  const first = paragraphs[0]!.paraId;
  const last = paragraphs.at(-1)!;
  const deletion = command(arm, 'replaceSelection', { text: '', selection: {
    anchor: { story: 'body', paraId: first, offset: 4 }, focus: { story: 'body', paraId: last.paraId, offset: last.text.length },
  } });
  const click = command(arm, 'selectPoint', {
    page: { pageId: page.pageId.toString(), pageIndex: page.pageIndex }, x: 10000, y: 10000,
    layoutRevision: before.layoutRevision, frameEpoch: before.frame.frameEpoch, clickCount: 1,
    modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  }, { afterSeq: deletion.seq, basis: { epoch: 1, version: 0 } });
  const typing = command(arm, 'insertText', { text: '!' }, { afterSeq: click.seq, basis: { epoch: 1, version: 0 } });
  const deleted = execute(arm, deletion);
  expect(deleted.status).toBe('applied');
  const after = layout(arm);
  expect(after.frame.pages.some((entry) => entry.pageId === page.pageId)).toBe(false);
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockReturnValue(fakeQueries({
    isReady: () => true,
    dispose: () => {},
    hitTestRegions: (pageIndex: number) => {
      expect(pageIndex).toBe(0);
      return { region: 'body', pos: 5, target: 'text' };
    },
  }));
  try {
    const resolved = resolveDocxSelectionIntent({ engine: arm.engine, state: arm.state, command: click as DocxSelectionCommandIntent, context: context(arm, click, deleted.selection), readLayout: () => after });
    expect(resolved.status).toBe('clamped');
    expect(resolved.selection).toMatchObject(selection(first, 4));
    const selected = ordered(arm, click, deleted.selection ?? null);
    expect(selected.status).toBe('noop');
    const typed = execute(arm, typing, selected.selection);
    expect(typed.status).toBe('applied');
    expect(texts(arm)).toEqual(['Keep!']);
    expect(typed.selection).toMatchObject(selection(first, 5));
    expect([deleted, selected, typed].every((result) => result.status !== 'refused')).toBe(true);
  } finally {
    queries.mockRestore();
  }
}, TIMEOUT);

test('d2a_tracked_deleted_prefix_keeps_selection_typing_and_navigation_in_raw_offsets', async () => {
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockReturnValue(fakeQueries({
    isReady: () => true, dispose: () => {},
    hitTestRegions: () => ({ region: 'body', pos: 4, target: 'text' }),
  }));
  try {
    for (const clickCount of [2, 3]) {
      const arm = await openCommandEngine([], { bodyXml:
        '<w:p w14:paraId="00000001"><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>XX</w:delText></w:r></w:del><w:r><w:t>abc</w:t></w:r></w:p>',
      });
      const paraId = arm.engine.geometryReader.paragraphSpans('body')[0]!.paraId;
      expect(arm.engine.geometryReader.paragraphSpans('body')[0]!.length).toBe(5);
      const frame = fakeFrame([{ pageIndex: 0, width: 600, height: 800, primitives: [] }]);
      arm.state.selection = selection(paraId, 2);
      const backward = runNavigation(arm, frame, { unit: 'character', direction: 'backward', extend: false });
      expect(backward.selection).toEqual(selection(paraId, 1));
      arm.state.selection = backward.selection;
      for (const offset of [2, 3]) {
        const forward = runNavigation(arm, frame, { unit: 'character', direction: 'forward', extend: false });
        expect(forward.selection).toEqual(selection(paraId, offset));
        arm.state.selection = forward.selection;
      }
      expect(runNavigation(arm, frame, { unit: 'word', direction: 'forward', extend: false }).selection).toEqual(selection(paraId, 5));
      const point = command(arm, 'selectPoint', {
        page: { pageId: '1', pageIndex: 0 }, x: 30, y: 30, layoutRevision: 0, frameEpoch: 1, clickCount,
        modifiers: { shift: false, alt: false, ctrl: false, meta: false },
      });
      const checked = preflightDocxCommand(arm.engine, arm.state, point, context(arm, point), {
        readSelectionLayout: () => ({ frame, layoutRevision: 0 }),
      });
      expect(checked.status).toBe('ready');
      if (checked.status !== 'ready') throw new Error('Point selection refused');
      const selected = applyDocxCommand(arm.engine, arm.state, checked.prepared, context(arm, point));
      expect(selected.selection).toMatchObject(selection(paraId, 0, 5));
      const typing = command(arm, 'insertText', { text: '!' }, { afterSeq: point.seq });
      expect(execute(arm, typing, selected.selection).status).toBe('applied');
      expect(texts(arm)).toEqual(['!']);
    }
  } finally {
    queries.mockRestore();
  }
}, TIMEOUT);

test('d2a_worker_vertical_navigation_builds_unbuilt_neighbours_before_chained_typing', async () => {
  const arm = await openCommandEngine(Array.from({ length: 150 }, (_, index) => `Paragraph ${index}`));
  await loadRustDisplayListQueryEngine();
  const complete = layout(arm).frame;
  expect(complete.pages.length).toBeGreaterThan(1);
  const queries = queryModule.createDisplayListQueries(complete.displayList, loadedRustDisplayListQueryEngine()!);
  const line = queries.visualLinesOnPage(0).at(-1)!;
  queries.dispose();
  const index = new DisplayPositionIndex({ ...arm.engine.geometryReader, selectionText: arm.engine.selectionText, resolveComment: arm.engine.resolveComment });
  const at = displayPositionToYrsLoc(index.inputMap('body')!, line.from)!;
  arm.state.selection = selection(at.paraId, at.offset);
  const navigation: DocxNavigationIntent = { unit: 'line', direction: 'down', extend: false, goalX: line.x };
  const expected = runNavigation(arm, complete, navigation).selection;
  expect(expected.focus).not.toEqual(arm.state.selection.focus);
  const released = arm.engine.releaseDisplayPagesFrame(complete.pages.slice(1).map((page) => page.pageIndex), complete.frameEpoch);
  if (!released) throw new Error('Pages were not released');
  const windowed = applyFrameDelta(complete, decodeFrameDelta(released));
  expect(windowed.pages[1]!.page.unbuilt).toBe(true);
  const builds = spyOn(arm.engine, 'buildDisplayPagesFrame');
  const factory: ResidentCommandHooksFactory = (_engine, _init, environment) => ({
    preflight: (intent, ctx) => preflightDocxCommand(arm.engine, arm.state, intent, ctx, environment),
    apply: (prepared, ctx) => applyDocxCommand(arm.engine, arm.state, prepared, ctx),
    selection: () => arm.state.selection,
    editState: () => { throw new Error('Unused edit-state projection'); },
  });
  const source = readFileSync(new URL('./residentEngineWorker.ts', import.meta.url), 'utf8');
  const binding = source.slice(source.indexOf('function bindResidentCommands('), source.indexOf('function replacementCommandInit('));
  const dependencies = {
    session: arm.engine, commandHooksFactory: factory, collaboration: null,
    init: { layoutWithRegions: true, layoutInput: '{}', extras: '{}' } as ResidentCommandInitRequest,
    retainedFrame: windowed, retainedLayoutVersion: null, layoutRevision: 0, commandFrameChanges: null,
    commandInit: null, commandHost: null, releaseCommandHost: null, releaseResidentSelectionReads,
    readResidentFrameGeometry, resetResidentObjectHandles, proposals: null, createProposalRegistry, proposalRevisionPreview,
    sessionFonts: null, ResidentSessionFonts, scope: { fonts: undefined },
    createResidentCommandHost: (options: ResidentCommandHostOptions) => options.hooks,
    displayPositionIndex: () => index, createDisplayListQueries: queryModule.createDisplayListQueries,
    loadRustDisplayListQueryEngine, loadedRustDisplayListQueryEngine, applyFrameDelta, decodeFrameDelta,
    retainedPageByKey: (frame: RetainedFrame, pageId: string) => retainedFramePageById(frame, BigInt(pageId)),
  };
  const bind = () => {
    const worker = new Function(...Object.keys(dependencies),
      new Bun.Transpiler({ loader: 'ts' }).transformSync(binding) +
      '\nreturn { hooks: bindResidentCommands(init), release: () => { releaseCommandHost?.(); releaseCommandHost = null; } };'
    )(...Object.values(dependencies)) as { hooks: ResidentCommandHooks; release(): void };
    releasesAfterTest.push(worker.release);
    return worker.hooks;
  };
  let hooks = bind();
  try {
    for (const [direction, start, destination] of [
      ['down', arm.state.selection, expected], ['up', expected, arm.state.selection],
    ] as const) {
      if (direction === 'up') {
        const current = applyFrameDelta(null, decodeFrameDelta(arm.engine.buildDisplayListFrame('{}', 0)));
        const release = arm.engine.releaseDisplayPagesFrame([0, 1], current.frameEpoch);
        if (!release) throw new Error('Caret and destination pages were not released');
        dependencies.retainedFrame = applyFrameDelta(current, decodeFrameDelta(release));
        expect(dependencies.retainedFrame.pages[0]!.page.unbuilt).toBe(true);
        expect(dependencies.retainedFrame.pages[1]!.page.unbuilt).toBe(true);
        hooks = bind();
      }
      arm.state.selection = start;
      arm.engine.setSelection(start.anchor, start.focus);
      const intent = command(arm, 'navigate', { ...navigation, direction });
      const checked = await hooks.preflight(intent, context(arm, intent));
      expect(checked.status).toBe('ready');
      if (checked.status !== 'ready') throw new Error('Navigation refused');
      expect(checked.prepared.selection?.focus).toEqual(destination.focus);
      const selected = await hooks.apply(checked.prepared, context(arm, intent));
      expect(selected.selection?.focus).toMatchObject(destination.focus);
    }
    expect(builds.mock.calls.some(([pages]) => pages.includes(1))).toBe(true);
    expect(builds.mock.calls.some(([pages]) => pages.includes(0))).toBe(true);
    const down = command(arm, 'navigate', navigation);
    const checked = await hooks.preflight(down, context(arm, down));
    if (checked.status !== 'ready') throw new Error('Navigation refused');
    const selected = await hooks.apply(checked.prepared, context(arm, down));
    const typing = command(arm, 'insertText', { text: '!' }, { afterSeq: down.seq });
    const typed = execute(arm, typing, selected.selection);
    expect(typed.status).toBe('applied');
    expect(typed.selection?.focus).toMatchObject({ ...expected.focus, offset: expected.focus.offset + 1 });
  } finally {
    builds.mockRestore();
  }
}, TIMEOUT);

test('d2a_worker_arrow_down_builds_the_actual_split_row_caret_page_and_neighbour', async () => {
  const cell = 'body:t0:r0c0';
  const arm = fakeArm({ body: ['Before'], [cell]: ['abcdefghijklmnopqrst'] });
  const segments = arm.engine.geometryReader.storySegments;
  arm.engine.geometryReader.storySegments = (story) => story === 'body' ? [
    ...segments(story),
    { kind: 'embed', embedKind: 'table', attributes: {}, payload: { rows: [{ cells: [{ story: cell }] }] } },
  ] : segments(story);
  arm.state.selection = selection(`${cell}-0`, 10, 10, cell);
  const index = new DisplayPositionIndex({ ...arm.engine.geometryReader, selectionText: arm.engine.selectionText, resolveComment: arm.engine.resolveComment });
  const position = index.positionOf(arm.state.selection.focus, 'body')!;
  let frame = fakeFrame([0, 1, 2].map<DisplayPage>((pageIndex) => ({
    pageIndex, width: 600, height: 800, primitives: [], unbuilt: true,
    positionSpan: pageIndex < 2 ? [position - 10, position + 10] : [position + 1, position + 10],
  })));
  const builds: number[][] = [];
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockImplementation((list) => fakeQueries({
    isReady: () => true, dispose: () => {},
    caretRect: () => list.pages[1]!.unbuilt ? null : { pageIndex: 1, x: 20, y: 40, width: 0, height: 10 },
    verticalMove: (at: number) => ({
      position: list.pages[1]!.unbuilt || list.pages[2]!.unbuilt ? at : at + 1, goalX: 20,
    }),
  }));
  try {
    Object.assign(arm.engine, { releaseDisplayPagesFrame: (pages: number[], frameEpoch: number) => {
      expect(frameEpoch).toBe(frame.frameEpoch);
      frame = fakeFrame(frame.displayList.pages.map((page) => pages.includes(page.pageIndex) ? { ...page, unbuilt: true } : page), frameEpoch + 1);
      return frame;
    } });
    const worker = workerSelectionHooks(arm, frame, (pages, frameEpoch) => {
      expect(frameEpoch).toBe(frame.frameEpoch);
      builds.push([...pages]);
      frame = fakeFrame(frame.displayList.pages.map((page) => pages.includes(page.pageIndex) ? { ...page, unbuilt: false } : page), frameEpoch + 1);
      return frame;
    });
    const intent = command(arm, 'navigate', { unit: 'line', direction: 'down', extend: false, goalX: 20 });
    const checked = await worker.hooks.preflight(intent, context(arm, intent));
    expect(checked.status).toBe('ready');
    if (checked.status !== 'ready') throw new Error('Navigation refused');
    expect(checked.prepared.selection).toEqual(selection(`${cell}-0`, 11, 11, cell));
    expect(frame.pages[1]!.page.unbuilt).toBe(false);
    expect(frame.pages[2]!.page.unbuilt).toBe(false);
    expect(builds).toEqual([[0], [1], [2]]);
  } finally {
    queries.mockRestore();
  }
});

test('d2a_worker_navigation_reuses_queries_within_a_frame_and_refreshes_after_install', async () => {
  const arm = fakeArm({ body: ['Active'] });
  arm.state.selection = selection('body-0', 2);
  const frame = fakeFrame(Array.from({ length: 30 }, (_, pageIndex) => ({ pageIndex, width: 600, height: 800, primitives: [] })));
  let disposed = 0;
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockImplementation(() => fakeQueries({
    isReady: () => true, dispose: () => { disposed++; },
    caretRect: () => ({ pageIndex: 12, x: 20, y: 40, width: 0, height: 10 }),
    verticalMove: (position: number, direction: 'up' | 'down') => ({ position: position + (direction === 'up' ? -1 : 1), goalX: 20 }),
  }));
  try {
    const worker = workerSelectionHooks(arm, frame);
    const navigate = async (direction: 'up' | 'down') => {
      const intent = command(arm, 'navigate', { unit: 'line', direction, extend: false, goalX: 20 });
      const checked = await worker.hooks.preflight(intent, context(arm, intent));
      expect(checked.status).toBe('ready');
      if (checked.status !== 'ready') throw new Error('Navigation refused');
      expect(checked.prepared.selection).toEqual(selection('body-0', direction === 'up' ? 1 : 3));
    };
    for (const direction of ['down', 'up', 'down', 'up'] as const) await navigate(direction);
    expect(queries).toHaveBeenCalledTimes(2);
    expect(disposed).toBe(0);
    const next = { ...frame, frameEpoch: 2 };
    worker.install(next, 1);
    await navigate('down');
    expect(queries).toHaveBeenCalledTimes(4);
    expect(disposed).toBe(2);
    await navigate('up');
    expect(queries).toHaveBeenCalledTimes(4);
    expect(disposed).toBe(2);
    worker.install(next, 2);
    await navigate('down');
    expect(queries).toHaveBeenCalledTimes(5);
    expect(disposed).toBe(3);
  } finally {
    queries.mockRestore();
  }
});

test('d2a_repeated_multiline_headers_navigate_within_the_active_rendering', async () => {
  await loadRustDisplayListQueryEngine();
  const arm = fakeArm({ body: ['Body'], 'hf:head': ['First', 'Last'] });
  const primitives: TextRunPrimitive[] = ['First', 'Last'].map((text, ordinal) => ({
    kind: 'text', text, paraId: `hf:head-${ordinal}`, docStart: ordinal === 0 ? 1 : 8, docEnd: ordinal === 0 ? 6 : 12,
    x: 20, baselineY: 40 + ordinal * 20, width: 60, font: '11px sans-serif', color: '#000000',
  }));
  const frame = fakeFrame([0, 1, 2].map((pageIndex) => ({
    pageIndex, width: 600, height: 800, primitives: [],
    header: { kind: 'header', rId: 'head', y: 0, height: 80, primitives },
  })));
  const point = command(arm, 'selectPoint', {
    page: { pageId: '3', pageIndex: 2 }, x: 40, y: 57, layoutRevision: 0, frameEpoch: 1, clickCount: 1,
    modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  }) as DocxSelectionCommandIntent;
  arm.state.selection = resolveDocxSelectionIntent({
    engine: arm.engine, state: arm.state, command: point, context: context(arm, point), readLayout: () => ({ frame, layoutRevision: 0 }),
  }).selection;
  expect(arm.state.selection.focus).toMatchObject({ story: 'hf:head', paraId: 'hf:head-1' });
  const create = queryModule.createDisplayListQueries;
  const scoped = spyOn(queryModule, 'createDisplayListQueries').mockImplementation((list, ...args) => create(list, ...args));
  try {
    expect(runNavigation(arm, frame, { unit: 'line', direction: 'down', extend: false, goalX: 20 }).selection).toEqual(arm.state.selection);
    expect(runNavigation(arm, frame, { unit: 'line', direction: 'up', extend: false, goalX: 20 }).selection.focus.paraId).toBe('hf:head-0');
    expect(scoped.mock.calls[0]![0].pages.map((page) => page.pageIndex)).toEqual([2]);
  } finally {
    scoped.mockRestore();
  }
});

test('d2a_repeated_end_of_document_drag_and_line_results_use_indexed_positions', () => {
  const cells = Object.fromEntries(Array.from({ length: 100 }, (_, ordinal) => [`body:t0:r0c${ordinal}`, ['Cell']]));
  const arm = fakeArm({ body: Array.from({ length: 2000 }, () => 'Paragraph'), ...cells });
  const segments = arm.engine.geometryReader.storySegments;
  arm.engine.geometryReader.storySegments = (story) => story === 'body' ? [
    ...segments(story),
    { kind: 'embed', embedKind: 'table', attributes: {}, payload: { rows: [{ cells: Object.keys(cells).map((story) => ({ story })) }] } },
  ] : segments(story);
  const cell = 'body:t0:r0c99';
  const index = new DisplayPositionIndex({ ...arm.engine.geometryReader, selectionText: arm.engine.selectionText, resolveComment: arm.engine.resolveComment });
  const position = index.positionOf({ story: cell, paraId: `${cell}-0`, offset: 2 }, 'body')!;
  const retained = fakeFrame(Array.from({ length: 100 }, (_, pageIndex) => ({ pageIndex, width: 600, height: 800, primitives: [] })));
  const frame: RetainedFrame = { ...retained, pages: new Proxy(retained.pages, {
    get: (pages, key, receiver) => {
      if (key === 'find') throw new Error('Linear retained-page lookup');
      return Reflect.get(pages, key, receiver);
    },
  }) };
  const targets = spyOn(YrsPositionProjection.prototype, 'targetAt');
  const queries = spyOn(queryModule, 'createDisplayListQueries').mockReturnValue(fakeQueries({
    isReady: () => true, dispose: () => {},
    hitTestRegions: () => ({ region: 'body', pos: position, target: 'text' }),
    caretRect: () => ({ pageIndex: 99, x: 0, y: 0, width: 0, height: 10 }),
    visualLinesOnPage: () => [fakeLine(position - 2, position + 2, 99)],
  }));
  const point = command(arm, 'selectPoint', {
    page: { pageId: '100', pageIndex: 99 }, x: 20, y: 40, layoutRevision: 0, frameEpoch: 1, clickCount: 1,
    modifiers: { shift: false, alt: false, ctrl: false, meta: false },
  }) as DocxSelectionCommandIntent;
  const drag = () => resolveDocxSelectionIntent({
    engine: arm.engine, state: arm.state, command: point, context: context(arm, point), readLayout: () => ({ frame, layoutRevision: 0 }),
  });
  try {
    expect(drag().selection).toEqual(selection(`${cell}-0`, 2, 2, cell));
    const coldReads = targets.mock.calls.length;
    for (let repeat = 0; repeat < 20; repeat++) {
      arm.state.selection = drag().selection;
      expect(runNavigation(arm, frame, { unit: 'line', direction: 'forward', extend: false }).selection).toEqual(selection(`${cell}-0`, 4, 4, cell));
    }
    expect(targets).toHaveBeenCalledTimes(coldReads);
  } finally {
    targets.mockRestore();
    queries.mockRestore();
  }
});

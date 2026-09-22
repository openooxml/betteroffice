/**
 * docx end-to-end: open pinned corpus documents, run an editing session through
 * the resident layout engine against each and check the results, timing every
 * operation across the wasm boundary together with the stage breakdown the
 * engine reports for it.
 */

import { afterAll, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../../packages/docx/src/docx/rezip';
import { buildResidentRegionLayoutRequest } from '../../packages/docx/src/editor/computeLayout';
import type {
  ResidentFontRequirement,
  ResidentMeasurementConfig,
} from '../../packages/docx/src/layout/measure';
import type { Layout } from '../../packages/docx/src/layout/pagination';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../../packages/docx/src/layout/render/frameDelta';
import type { RetainedFrame } from '../../packages/docx/src/layout/render/frameDelta';
import type { Document } from '../../packages/docx/src/types/document';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { preloadOpcWasm } from '../../packages/docx/src/wasm/opc';
import { preloadParseWasm } from '../../packages/docx/src/wasm/parse';
import { createYrsSession, yrsToDocument } from '../../packages/docx/src/yrs';
import type {
  YrsEngineApplyProfile,
  YrsParagraph,
  YrsParagraphLength,
  YrsSession,
  YrsStoryRange,
  YrsTextMatch,
} from '../../packages/docx/src/yrs';
import { loadSample, samplesFor } from './corpus';
import { SampleRecorder, e2eEnabled, finishFormat } from './harness';
import type { SampleRun, StageProfile } from './harness';

const WASM = resolve(import.meta.dir, '../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(import.meta.dir, '../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const STORY = 'body';
const PAGE_GAP = 24;
const WORD = 'Zephyr';
const PHRASE = 'quartz lantern drifts over the harbor';
const TYPED = `${WORD} ${PHRASE}`;
/** Longer than the engine's 500ms undo capture window, so the next edit is its own step. */
const UNDO_STEP_GAP_MS = 600;

interface RegionLayout {
  layout: Layout;
  notesConverged: boolean;
}

const suite = e2eEnabled() ? describe : describe.skip;
const runs: SampleRun[] = [];

suite('docx end-to-end', () => {
  beforeAll(() =>
    Promise.all([
      preloadEditWasm(new Uint8Array(readFileSync(WASM))),
      preloadOpcWasm(),
      preloadParseWasm(),
    ])
  );
  afterAll(() => finishFormat('docx', runs));

  for (const sample of samplesFor('docx')) {
    it(sample.id, async () => {
      const bytes = await loadSample(sample);
      const session = await createYrsSession();
      try {
        const started = performance.now();
        const opened = session.openDocx(bytes, true).document;
        const recorder = new SampleRecorder('docx', sample.id, performance.now() - started);
        await scenario(session, opened, recorder);
        runs.push(recorder.finish());
      } finally {
        session.destroy();
      }
    });
  }
});

function applyStages(profile: YrsEngineApplyProfile): StageProfile {
  return {
    selection: profile.selectionMs,
    edit: profile.editMs,
    lower: profile.lowerMs,
    measure: profile.measureMs,
    paginate: profile.paginateMs,
    displayInput: profile.displayInputMs,
    displayBuild: profile.displayBuildMs,
    displayFinalize: profile.displayFinalizeMs,
    encode: profile.encodeMs,
  };
}

/** Every requirement resolves to the one registered face, so measurement never depends on the host. */
function measurement(requirements: ResidentFontRequirement[], fontId: number): ResidentMeasurementConfig {
  return {
    fontChains: Object.fromEntries(requirements.map((requirement) => [requirement.key, [fontId]])),
    defaults: { fontSize: 11, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
}

interface TypingTarget {
  paraId: string;
  text: string;
  /** Paragraph length in story units, so the caret lands after any trailing embed. */
  end: number;
}

/** The first paragraph with visible, unbolded text: typing there inherits no bold, so the toggle adds it. */
function typingTarget(
  session: YrsSession,
  paragraphs: YrsParagraph[],
  spans: YrsParagraphLength[]
): TypingTarget {
  const lengths = new Map(spans.map((span) => [span.paraId, span.length]));
  for (const paragraph of paragraphs) {
    if (!/\S/.test(paragraph.text)) continue;
    const end = lengths.get(paragraph.paraId)!;
    const whole: YrsStoryRange = {
      story: STORY,
      start: { paraId: paragraph.paraId, offset: 0 },
      end: { paraId: paragraph.paraId, offset: end },
    };
    if (session.selectionContext(whole).bold === false) {
      return { paraId: paragraph.paraId, text: paragraph.text, end };
    }
  }
  throw new Error('no unbolded body paragraph to type into');
}

async function scenario(session: YrsSession, opened: Document, recorder: SampleRecorder): Promise<void> {
  const fontBytes = new Uint8Array(readFileSync(FONT));
  const fontId = recorder.op('registerFont', () => session.registerFont(fontBytes));
  expect(fontId).toBeGreaterThanOrEqual(0);

  const paragraphs = recorder.op('paragraphs:body', () => session.paragraphs(STORY));
  expect(paragraphs.length).toBeGreaterThan(0);
  const spans = recorder.op('paragraphSpans:body', () => session.paragraphSpans(STORY));
  const target = typingTarget(session, paragraphs, spans);

  const request = buildResidentRegionLayoutRequest(opened, PAGE_GAP, {});
  const requirements = recorder.op(
    'layoutFontRequirements',
    () => JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[]
  );
  expect(requirements.length).toBeGreaterThan(0);
  request.measurement = measurement(requirements, fontId);
  const layoutInput = JSON.stringify(request);

  const initial = regionLayout(session, recorder, 'layoutDocumentWithRegions:initial', layoutInput);
  expect(initial.layout.pages.length).toBeGreaterThan(0);

  const caret = { story: STORY, paraId: target.paraId, offset: target.end };
  recorder.op('setSelection', () => session.setSelection(caret));
  expect(session.encodeSelection()?.story).toBe(STORY);

  let frame = displayFrame(session, recorder, 'displayListFrame:initial', null);
  expect(frame.pages.length).toBe(initial.layout.pages.length);
  expect(frame.displayList.pages[0].primitives.length).toBeGreaterThan(0);

  const snapshot = recorder.op('residentCaretSnapshot', () => session.residentCaretSnapshot());
  expect(snapshot.frameEpoch).toBe(frame.frameEpoch);
  expect(snapshot.caretRect).not.toBeNull();

  frame = typing(session, recorder, 'applyInput:word', WORD, frame);
  expectTyped(session, caret, WORD);
  expect(frameText(frame).toLowerCase()).toContain(WORD.toLowerCase());

  frame = typing(session, recorder, 'applyInput:space', ' ', frame);
  expectTyped(session, caret, `${WORD} `);

  frame = typing(session, recorder, 'applyInput:phrase', PHRASE, frame);
  expectTyped(session, caret, TYPED);
  expect(frameText(frame).toLowerCase()).toContain('harbor');

  const split = recorder.op('splitParagraph', () => session.splitParagraph(caret));
  expect(split.firstParaId).toBe(target.paraId);
  expect(paragraphText(session, target.paraId)).toBe(target.text);
  const typedParaId = split.secondParaId;
  expect(paragraphText(session, typedParaId)).toBe(TYPED);

  const storiesBefore = session.storyIds().length;
  const table = recorder.op('insertTable', () =>
    session.insertTable({ story: STORY, paraId: typedParaId, offset: 0 }, 2, 2)
  );
  expect(table).toMatchObject({ rows: 2, columns: 2, deletedTable: false });
  expect(table.createdStoryIds).toHaveLength(4);
  expect(session.storyIds()).toHaveLength(storiesBefore + 4);
  expect(paragraphText(session, typedParaId)).toBe(TYPED);

  const range = wordRange(recorder.op('searchText:typed', () => session.searchText(WORD)), typedParaId);

  expect(recorder.op('selectionContext', () => session.selectionContext(range)).bold).toBe(false);
  await Bun.sleep(UNDO_STEP_GAP_MS);
  recorder.op('toggleMark:bold', () => session.toggleMark(range, { type: 'bold' }));
  expect(session.selectionContext(range).bold).toBe(true);

  expect(recorder.op('undo:toggleMark', () => session.undo())).toBe(true);
  expect(session.selectionContext(range).bold).toBe(false);
  expect(paragraphText(session, typedParaId)).toBe(TYPED);
  expect(recorder.op('redo:toggleMark', () => session.redo())).toBe(true);
  expect(session.selectionContext(range).bold).toBe(true);

  const edited = regionLayout(session, recorder, 'layoutDocumentWithRegions:afterEdits', layoutInput);
  expect(edited.layout.pages.length).toBeGreaterThanOrEqual(initial.layout.pages.length);
  frame = displayFrame(session, recorder, 'displayListFrame:afterEdits', frame);
  expect(frame.pages.length).toBe(edited.layout.pages.length);
  expect(frameText(frame).toLowerCase()).toContain(WORD.toLowerCase());

  const materialized = recorder.op('materializeDocx', () => session.materializeDocx());
  expect(materialized).not.toBeNull();
  expect(blocks(materialized!, 'paragraph')).toBeGreaterThan(0);

  const projected = recorder.op('save:project', () => yrsToDocument(session, materialized!));
  expect(blocks(projected, 'paragraph')).toBe(blocks(materialized!, 'paragraph') + 1);
  expect(blocks(projected, 'table')).toBe(blocks(materialized!, 'table') + 1);

  const saved = await timed(recorder, 'save:repack', () => repackDocx(projected));
  expect(saved.byteLength).toBeGreaterThan(0);

  const reopened = await createYrsSession();
  try {
    recorder.op('reopen', () => reopened.openDocx(new Uint8Array(saved), true));
    expect(reopened.storyIds()).toHaveLength(session.storyIds().length);
    const persisted = recorder
      .op('paragraphs:afterReopen', () => reopened.paragraphs(STORY))
      .find((paragraph) => paragraph.text === TYPED);
    expect(persisted).toBeDefined();
    const found = recorder.op('searchText:afterReopen', () => reopened.searchText(WORD));
    expect(reopened.selectionContext(wordRange(found, persisted!.paraId)).bold).toBe(true);
  } finally {
    reopened.destroy();
  }
}

function regionLayout(session: YrsSession, recorder: SampleRecorder, op: string, input: string): RegionLayout {
  return recorder.op(
    op,
    () => JSON.parse(session.layoutDocumentWithRegionsRetainedJson(input)) as RegionLayout
  );
}

/** Build the next frame against the one the host holds and fold it in, as the editor does. */
function displayFrame(
  session: YrsSession,
  recorder: SampleRecorder,
  op: string,
  previous: RetainedFrame | null
): RetainedFrame {
  return recorder.op(op, () =>
    applyFrameDeltaOwned(
      previous,
      decodeFrameDelta(session.buildDisplayListFrame('{}', previous?.frameEpoch ?? 0))
    )
  );
}

function typing(
  session: YrsSession,
  recorder: SampleRecorder,
  op: string,
  text: string,
  previous: RetainedFrame
): RetainedFrame {
  let profile: YrsEngineApplyProfile | undefined;
  return recorder.op(
    op,
    () => {
      const result = session.applyInputProfiled(text, previous.frameEpoch);
      profile = result.profile;
      return applyFrameDeltaOwned(previous, decodeFrameDelta(result.frame));
    },
    () => applyStages(profile!)
  );
}

/** Typing at `caret` appended `text` and left the head after it. */
function expectTyped(session: YrsSession, caret: { paraId: string; offset: number }, text: string): void {
  expect(paragraphText(session, caret.paraId).endsWith(text)).toBe(true);
  expect(session.selection()?.head).toEqual({
    story: STORY,
    paraId: caret.paraId,
    offset: caret.offset + text.length,
  });
}

function paragraphText(session: YrsSession, paraId: string): string {
  const paragraph = session.paragraphs(STORY).find((entry) => entry.paraId === paraId);
  if (!paragraph) throw new Error(`paragraph ${paraId} is gone`);
  return paragraph.text;
}

/** The typed word's range in `paraId`, taken from search so embeds ahead of it do not matter. */
function wordRange(matches: YrsTextMatch[], paraId: string): YrsStoryRange {
  const match = matches.find((entry) => entry.story === STORY && entry.paraId === paraId);
  expect(match).toBeDefined();
  expect(match!.end - match!.start).toBe(WORD.length);
  return {
    story: STORY,
    start: { paraId, offset: match!.start },
    end: { paraId, offset: match!.end },
  };
}

function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages
    .flatMap((page) => page.primitives)
    .map((primitive) => ('text' in primitive && typeof primitive.text === 'string' ? primitive.text : ''))
    .join('');
}

function blocks(document: Document, type: 'paragraph' | 'table'): number {
  return document.package.document.content.filter((block) => block.type === type).length;
}

/** Time an awaited operation the way `SampleRecorder.op` times a synchronous one. */
async function timed<T>(recorder: SampleRecorder, op: string, run: () => Promise<T>): Promise<T> {
  const started = performance.now();
  const value = await run();
  recorder.ops.push({ op, e2eMs: performance.now() - started });
  return value;
}

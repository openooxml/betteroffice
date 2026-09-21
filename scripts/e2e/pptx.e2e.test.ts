/**
 * pptx end-to-end: open pinned corpus decks, run an editing session against
 * each and check the results, timing every operation across the wasm boundary
 * together with the stage breakdown the renderer and the edit boundary report.
 */

import { afterAll, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { initWasm, openPresentation } from '../../packages/pptx/src/wasm/loader';
import type { PresentationHandle } from '../../packages/pptx/src/wasm/loader';
import type {
  EditProfile,
  HistoryProfile,
  LayoutProfile,
  PptxFontFace,
  Profiled,
  ShapeSnapshot,
  SlideDisplayList,
  StorySnapshot,
  TextBoxPrimitive,
} from '../../packages/pptx/src/types';
import { loadSample, samplesFor } from './corpus';
import { SampleRecorder, e2eEnabled, finishFormat } from './harness';
import type { SampleRun, StageProfile } from './harness';

const WASM = resolve(import.meta.dir, '../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm');
const FONT = resolve(import.meta.dir, '../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const SLIDE = 0;
const MARKER = 'E2E ';
const NOTE = 'Added by the e2e run';
/** EMU, inside every slide of the corpus. */
const BOX = { x: 914_400, y: 914_400, width: 3_657_600, height: 914_400 };
const MOVED = { x: 1_828_800, y: 2_743_200 };

const suite = e2eEnabled() ? describe : describe.skip;
const runs: SampleRun[] = [];
let fonts: PptxFontFace[] = [];

suite('pptx end-to-end', () => {
  beforeAll(() => {
    fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(readFileSync(FONT)) }];
    return initWasm(new Uint8Array(readFileSync(WASM)));
  });
  afterAll(() => finishFormat('pptx', runs));

  for (const sample of samplesFor('pptx')) {
    it(sample.id, async () => {
      const bytes = await loadSample(sample);
      const started = performance.now();
      const handle = openPresentation(bytes, { fonts });
      const recorder = new SampleRecorder('pptx', sample.id, performance.now() - started);
      try {
        scenario(handle, recorder);
      } finally {
        handle.dispose();
      }
      runs.push(recorder.finish());
    });
  }
});

function layoutStages(profile: LayoutProfile): StageProfile {
  return { scope: profile.scopeMs, layout: profile.layoutMs, serialize: profile.serializeMs };
}

function editStages(profile: EditProfile): StageProfile {
  return { parse: profile.parseMs, apply: profile.applyMs, serialize: profile.serializeMs };
}

function historyStages(profile: HistoryProfile): StageProfile {
  return { undo: profile.undoMs, snapshot: profile.snapshotMs, serialize: profile.serializeMs };
}

function scenario(handle: PresentationHandle, recorder: SampleRecorder): void {
  const deck = recorder.op('snapshot', () => handle.snapshot());
  expect(deck.slides.length).toBeGreaterThan(0);
  const slide = deck.slides[SLIDE];
  const story = firstStory(slide.shapes);
  const original = storyText(story);

  const initial = profiledLayout(handle, recorder, 'layoutSlide:initial', SLIDE);
  expect(initial.primitives.length).toBeGreaterThan(0);
  expect(runCount(initial)).toBeGreaterThan(0);
  expect(laidOutText(textBoxOf(initial, story.id)).length).toBeGreaterThan(0);

  const inserted = profiled(
    recorder,
    'insertText',
    () => handle.insertTextProfiled(story.id, 0, MARKER),
    editStages
  );
  expect(inserted).toMatchObject({ storyId: story.id, start: 0, end: MARKER.length });
  expect(storyText(recorder.op('story:afterInsertText', () => handle.story(story.id)))).toBe(
    MARKER + original
  );
  const afterInsert = profiledLayout(handle, recorder, 'layoutSlide:afterInsertText', SLIDE);
  expect(laidOutText(textBoxOf(afterInsert, story.id))).toContain(MARKER.trim());

  const added = profiled(
    recorder,
    'addTextBox',
    () =>
      handle.addTextBoxProfiled(slide.id, {
        name: 'E2E note',
        rect: BOX,
        text: NOTE,
        style: { fontSizePt: 18, fontFamily: 'Liberation Sans' },
      }),
    editStages
  );
  expect(added.slideId).toBe(slide.id);
  const withBox = recorder.op('snapshot:afterAddTextBox', () => handle.snapshot()).slides[SLIDE];
  expect(withBox.shapes.length).toBe(slide.shapes.length + 1);
  const box = withBox.shapes.find((shape) => shape.id === added.shapeId);
  expect(box).toMatchObject(BOX);
  const boxStory = box!.textStories[0].id;
  const afterAdd = profiledLayout(handle, recorder, 'layoutSlide:afterAddTextBox', SLIDE);
  const boxBefore = textBoxOf(afterAdd, boxStory);
  expect(laidOutText(boxBefore)).toContain(NOTE.split(' ')[0]);

  const moved = profiled(
    recorder,
    'moveShape',
    () => handle.moveShapeProfiled(slide.id, added.shapeId, MOVED.x, MOVED.y),
    editStages
  );
  expect(moved.before).toMatchObject({ x: BOX.x, y: BOX.y });
  expect(moved.after).toMatchObject({ ...MOVED, width: BOX.width, height: BOX.height });
  const afterMove = profiledLayout(handle, recorder, 'layoutSlide:afterMoveShape', SLIDE);
  const boxAfter = textBoxOf(afterMove, boxStory);
  expect(boxAfter.x).toBeGreaterThan(boxBefore.x);
  expect(boxAfter.y).toBeGreaterThan(boxBefore.y);

  const count = deck.slides.length;
  const appended = profiled(
    recorder,
    'insertSlide',
    () => handle.insertSlideProfiled(count),
    editStages
  );
  expect(appended.toIndex).toBe(count);
  const grown = recorder.op('snapshot:afterInsertSlide', () => handle.snapshot());
  expect(grown.slides.length).toBe(count + 1);
  expect(grown.slides[count].id).toBe(appended.slideId);
  const blank = profiledLayout(handle, recorder, 'layoutSlide:insertedSlide', count);
  expect(blank).toMatchObject({ width: initial.width, height: initial.height });

  const deleted = profiled(
    recorder,
    'deleteText',
    () => handle.deleteTextProfiled(story.id, 0, MARKER.length),
    editStages
  );
  expect(deleted.text).toBe(MARKER);
  expect(storyText(recorder.op('story:afterDeleteText', () => handle.story(story.id)))).toBe(
    original
  );

  const undone = profiled(recorder, 'undo:deleteText', () => handle.undoProfiled(), historyStages);
  expect(undone.applied).toBe(true);
  expect(undone.snapshot.slides.length).toBe(count + 1);
  expect(storyText(recorder.op('story:afterUndo', () => handle.story(story.id)))).toBe(
    MARKER + original
  );

  const saved = recorder.op('save', () => handle.save());
  expect(saved.byteLength).toBeGreaterThan(0);

  const reopened = recorder.op('reopen', () => openPresentation(saved, { fonts }));
  try {
    const again = recorder.op('snapshot:afterReopen', () => reopened.snapshot());
    expect(again.slides.length).toBe(count + 1);
    expect(again.slides[SLIDE].shapes.length).toBe(slide.shapes.length + 1);
    const matches = recorder.op('searchText:afterReopen', () => reopened.searchText(MARKER.trim()));
    expect(matches.some((match) => match.slideIndex === SLIDE && match.start === 0)).toBe(true);
  } finally {
    reopened.dispose();
  }
}

function profiled<T, P>(
  recorder: SampleRecorder,
  op: string,
  run: () => Profiled<T, P>,
  stages: (profile: P) => StageProfile
): T {
  let profile: P | undefined;
  return recorder.op(
    op,
    () => {
      const result = run();
      profile = result.profile;
      return result.receipt;
    },
    () => stages(profile!)
  );
}

function profiledLayout(
  handle: PresentationHandle,
  recorder: SampleRecorder,
  op: string,
  slideIndex: number
): SlideDisplayList {
  let profile: LayoutProfile | undefined;
  return recorder.op(
    op,
    () => {
      const result = handle.layoutSlideProfiled(slideIndex);
      profile = result.profile;
      return result.layout;
    },
    () => layoutStages(profile!)
  );
}

/** The first visible top-level shape's story that carries text. */
function firstStory(shapes: ShapeSnapshot[]): StorySnapshot {
  for (const shape of shapes) {
    if (shape.hidden) continue;
    const story = shape.textStories.find((story) =>
      story.paragraphs.some((paragraph) => paragraph.runs.length > 0)
    );
    if (story) return story;
  }
  throw new Error('slide has no text story');
}

function storyText(story: StorySnapshot): string {
  return story.paragraphs.flatMap((paragraph) => paragraph.runs.map((run) => run.text)).join('');
}

function textBoxOf(layout: SlideDisplayList, storyId: string): TextBoxPrimitive {
  const box = layout.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' && primitive.storyId === storyId
  );
  expect(box, `text box for ${storyId}`).toBeDefined();
  return box!;
}

function laidOutText(box: TextBoxPrimitive): string {
  return box.lines.flatMap((line) => line.runs.map((run) => run.text)).join('');
}

function runCount(layout: SlideDisplayList): number {
  let runs = 0;
  for (const primitive of layout.primitives) {
    if (primitive.kind !== 'textBox') continue;
    for (const line of primitive.lines) runs += line.runs.length;
  }
  return runs;
}

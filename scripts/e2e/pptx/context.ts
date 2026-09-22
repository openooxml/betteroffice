/** Shared plumbing for the pptx scenarios: wasm init, fonts, handle lifetime, stage mapping. */

import { expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import {
  initWasm,
  openPresentation,
} from '../../../packages/pptx/src/wasm/loader';
import type { PresentationHandle } from '../../../packages/pptx/src/wasm/loader';
import type {
  DeckSnapshot,
  EditProfile,
  HistoryProfile,
  LayoutProfile,
  PptxFontFace,
  Profiled,
  ShapeSnapshot,
  SlideDisplayList,
  StorySnapshot,
  TextBoxPrimitive,
} from '../../../packages/pptx/src/types';
import type { PinnedSample } from '../corpus';
import type {
  ActorRecorder,
  ScenarioRecorder,
  StageProfile,
  Timer,
} from '../harness';
import type { Scenario } from '../suite';

const WASM = resolve(
  import.meta.dir,
  '../../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm'
);
const FONT = resolve(
  import.meta.dir,
  '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
export const FONT_FAMILY = 'Liberation Sans';

export interface PptxCtx {
  sample: PinnedSample;
  bytes: Uint8Array;
  recorder: ScenarioRecorder;
  fonts: PptxFontFace[];
  /** Opens a handle that is disposed with the scenario. */
  open(bytes?: Uint8Array, clientId?: number): PresentationHandle;
  dispose(): void;
}

export type PptxScenario = Scenario<PptxCtx>;

let fonts: PptxFontFace[] = [];

export function setup(): Promise<void> {
  fonts = [{ family: FONT_FAMILY, bytes: new Uint8Array(readFileSync(FONT)) }];
  return initWasm(new Uint8Array(readFileSync(WASM)));
}

export function context(
  sample: PinnedSample,
  bytes: Uint8Array,
  recorder: ScenarioRecorder
): PptxCtx {
  const handles: PresentationHandle[] = [];
  return {
    sample,
    bytes,
    recorder,
    fonts,
    open(source = bytes, clientId?: number) {
      const handle = openPresentation(source, { fonts, clientId });
      handles.push(handle);
      return handle;
    },
    dispose() {
      for (const handle of handles) {
        try {
          handle.dispose();
        } catch {}
      }
    },
  };
}

export function layoutStages(profile: LayoutProfile): StageProfile {
  return {
    scope: profile.scopeMs,
    layout: profile.layoutMs,
    serialize: profile.serializeMs,
  };
}

export function editStages(profile: EditProfile): StageProfile {
  return {
    parse: profile.parseMs,
    apply: profile.applyMs,
    serialize: profile.serializeMs,
  };
}

export function historyStages(profile: HistoryProfile): StageProfile {
  return {
    undo: profile.undoMs,
    snapshot: profile.snapshotMs,
    serialize: profile.serializeMs,
  };
}

/** Records a `*Profiled` edit as one op and hands back its receipt. */
export function profiled<T, P>(
  timer: Timer,
  op: string,
  run: () => Profiled<T, P>,
  stages: (profile: P) => StageProfile
): T {
  let profile: P | undefined;
  return timer.op(
    op,
    () => {
      const result = run();
      profile = result.profile;
      return result.receipt;
    },
    () => stages(profile!)
  );
}

export function profiledLayout(
  handle: PresentationHandle,
  timer: Timer,
  op: string,
  slideIndex: number
): SlideDisplayList {
  let profile: LayoutProfile | undefined;
  return timer.op(
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
export function firstStory(shapes: ShapeSnapshot[]): StorySnapshot {
  for (const shape of shapes) {
    if (shape.hidden) continue;
    const story = shape.textStories.find((story) =>
      story.paragraphs.some((paragraph) => paragraph.runs.length > 0)
    );
    if (story) return story;
  }
  throw new Error('slide has no text story');
}

export function storyText(story: StorySnapshot): string {
  return story.paragraphs
    .flatMap((paragraph) => paragraph.runs.map((run) => run.text))
    .join('');
}

export function textBoxOf(
  layout: SlideDisplayList,
  storyId: string
): TextBoxPrimitive {
  const box = layout.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' && primitive.storyId === storyId
  );
  expect(box, `text box for ${storyId}`).toBeDefined();
  return box!;
}

export function laidOutText(box: TextBoxPrimitive): string {
  return box.lines.flatMap((line) => line.runs.map((run) => run.text)).join('');
}

export function runCount(layout: SlideDisplayList): number {
  let runs = 0;
  for (const primitive of layout.primitives) {
    if (primitive.kind !== 'textBox') continue;
    for (const line of primitive.lines) runs += line.runs.length;
  }
  return runs;
}

/** An 8x8 PNG, small enough to embed and still a real decodable image. */
export const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAbElEQVR4nA3JQQEAMAgDMZzgpE7qhMf5wAlO6mbLN1VFFypcTLHFFSmqmm7UuJlmm2vSP0QLCYsRK05EP0wbGZsxa87EP4YeNHiYYYcbMj+WXrR4mWWXW7I/jj50+JhjjztyP0IHBYcJGy4kPDtbVkGLfGO6AAAAAElFTkSuQmCC';

/** One collaborative replica with the update bytes its own edits produced. */
export interface Replica {
  name: string;
  handle: PresentationHandle;
  timer: ActorRecorder;
  outbox: Uint8Array[];
}

export function replicas(ctx: PptxCtx, names: string[]): Replica[] {
  return names.map((name, index) => {
    const timer = ctx.recorder.as(name);
    const handle = timer.load(() => ctx.open(ctx.bytes, index + 1));
    const outbox: Uint8Array[] = [];
    handle.onUpdate((update, origin) => {
      if (origin === 'local') outbox.push(update);
    });
    return { name, handle, timer, outbox };
  });
}

/** Delivers every pending update to every other replica, one recorded op each. */
export function exchange(peers: Replica[], op = 'applyUpdate'): number {
  let bytes = 0;
  for (const from of peers) {
    const pending = from.outbox.splice(0);
    for (const update of pending) {
      bytes += update.byteLength;
      for (const to of peers) {
        if (to === from) continue;
        to.timer.op(op, () => to.handle.applyUpdate(update), undefined, {
          from: from.name,
          bytes: update.byteLength,
        });
      }
    }
  }
  return bytes;
}

/** Canonical complete deck snapshot. */
export function fingerprint(deck: DeckSnapshot): string {
  return JSON.stringify(deck, (_key, value) =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? Object.fromEntries(
          Object.entries(value).sort(([a], [b]) => a.localeCompare(b))
        )
      : value
  );
}

export function shapeText(shape: ShapeSnapshot): string {
  return shape.textStories.map(storyText).join('');
}

/** The first paragraph's text; text ranges may not cross paragraph boundaries. */
export function firstParagraphText(story: StorySnapshot): string {
  return story.paragraphs[0].runs.map((run) => run.text).join('');
}

/** Paragraph texts joined the way the python binding reports a story. */
export function storyLines(story: StorySnapshot): string {
  return story.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

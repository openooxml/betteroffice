/** Shared plumbing for the docx scenarios: wasm preloads, sessions, layout requests, stage mapping. */

import { expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../../../packages/docx/src/docx/rezip';
import { buildResidentRegionLayoutRequest } from '../../../packages/docx/src/editor/computeLayout';
import type { ResidentFontRequirement, ResidentMeasurementConfig } from '../../../packages/docx/src/layout/measure';
import type { Layout } from '../../../packages/docx/src/layout/pagination';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../../../packages/docx/src/layout/render/frameDelta';
import type { RetainedFrame } from '../../../packages/docx/src/layout/render/frameDelta';
import type { Document } from '../../../packages/docx/src/types/document';
import { preloadEditWasm } from '../../../packages/docx/src/wasm/edit';
import { preloadOpcWasm } from '../../../packages/docx/src/wasm/opc';
import { preloadParseWasm } from '../../../packages/docx/src/wasm/parse';
import { createYrsSession, yrsToDocument } from '../../../packages/docx/src/yrs';
import type {
  YrsEngineApplyProfile,
  YrsParagraph,
  YrsParagraphLength,
  YrsSession,
  YrsStoryRange,
  YrsTextMatch,
} from '../../../packages/docx/src/yrs';
import type { PinnedSample } from '../corpus';
import type { ActorRecorder, ScenarioRecorder, StageProfile, Timer } from '../harness';
import type { Scenario } from '../suite';

const WASM = resolve(import.meta.dir, '../../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(import.meta.dir, '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
export const STORY = 'body';
export const PAGE_GAP = 24;
/** Longer than the engine's 500ms undo capture window, so the next edit is its own step. */
export const UNDO_STEP_GAP_MS = 600;

export interface Editor {
  session: YrsSession;
  document: Document;
  /** The layout request with every font requirement pinned to the test face. */
  layoutInput: string;
}

export interface DocxCtx {
  sample: PinnedSample;
  bytes: Uint8Array;
  recorder: ScenarioRecorder;
  /** Opens a session that is destroyed with the scenario. */
  open(source?: Uint8Array): Promise<Editor>;
  /** A second replica of `origin`, hydrated from its yrs state rather than re-seeded. */
  hydrate(origin: Editor): Promise<Editor>;
  dispose(): void;
}

export type DocxScenario = Scenario<DocxCtx>;

export interface RegionLayout {
  layout: Layout;
  notesConverged: boolean;
}

let fontBytes: Uint8Array;

export function setup(): Promise<unknown> {
  fontBytes = new Uint8Array(readFileSync(FONT));
  return Promise.all([preloadEditWasm(new Uint8Array(readFileSync(WASM))), preloadOpcWasm(), preloadParseWasm()]);
}

export function context(sample: PinnedSample, bytes: Uint8Array, recorder: ScenarioRecorder): DocxCtx {
  const sessions: YrsSession[] = [];
  return {
    sample,
    bytes,
    recorder,
    async open(source = bytes) {
      const session = await createYrsSession();
      sessions.push(session);
      const document = session.openDocx(source, true).document;
      const fontId = session.registerFont(fontBytes);
      expect(fontId).toBeGreaterThanOrEqual(0);
      const request = buildResidentRegionLayoutRequest(document, PAGE_GAP, {});
      const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(request))) as ResidentFontRequirement[];
      expect(requirements.length).toBeGreaterThan(0);
      request.measurement = measurement(requirements, fontId);
      return { session, document, layoutInput: JSON.stringify(request) };
    },
    async hydrate(origin) {
      const session = await createYrsSession();
      sessions.push(session);
      const document = session.openDocx(bytes, false).document;
      session.loadState(origin.session.encodeState());
      session.registerFont(fontBytes);
      session.beginUndoCapture();
      return { session, document, layoutInput: origin.layoutInput };
    },
    dispose() {
      for (const session of sessions) {
        try {
          session.destroy();
        } catch {}
      }
    },
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

export function applyStages(profile: YrsEngineApplyProfile): StageProfile {
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

export function regionLayout(editor: Editor, timer: Timer, op: string): RegionLayout {
  return timer.op(op, () => JSON.parse(editor.session.layoutDocumentWithRegionsRetainedJson(editor.layoutInput)) as RegionLayout);
}

/** Build the next frame against the one the host holds and fold it in, as the editor does. */
export function displayFrame(session: YrsSession, timer: Timer, op: string, previous: RetainedFrame | null): RetainedFrame {
  return timer.op(op, () => applyFrameDeltaOwned(previous, decodeFrameDelta(session.buildDisplayListFrame('{}', previous?.frameEpoch ?? 0))));
}

/** One profiled keystroke: apply the input and fold the frame delta it returned. */
export function typing(session: YrsSession, timer: Timer, op: string, text: string, previous: RetainedFrame): RetainedFrame {
  let profile: YrsEngineApplyProfile | undefined;
  return timer.op(
    op,
    () => {
      const result = session.applyInputProfiled(text, previous.frameEpoch);
      profile = result.profile;
      return applyFrameDeltaOwned(previous, decodeFrameDelta(result.frame));
    },
    () => applyStages(profile!)
  );
}

export interface TypingTarget {
  paraId: string;
  text: string;
  /** Paragraph length in story units, so the caret lands after any trailing embed. */
  end: number;
}

/** The first paragraph with visible, unbolded text: typing there inherits no bold, so a toggle adds it. */
export function typingTarget(session: YrsSession, paragraphs: YrsParagraph[], spans: YrsParagraphLength[]): TypingTarget {
  const lengths = new Map(spans.map((span) => [span.paraId, span.length]));
  for (const paragraph of paragraphs) {
    if (!/\S/.test(paragraph.text)) continue;
    const end = lengths.get(paragraph.paraId)!;
    const whole: YrsStoryRange = { story: STORY, start: { paraId: paragraph.paraId, offset: 0 }, end: { paraId: paragraph.paraId, offset: end } };
    if (session.selectionContext(whole).bold === false) return { paraId: paragraph.paraId, text: paragraph.text, end };
  }
  throw new Error('no unbolded body paragraph to type into');
}

export function paragraphText(session: YrsSession, paraId: string): string {
  const paragraph = session.paragraphs(STORY).find((entry) => entry.paraId === paraId);
  if (!paragraph) throw new Error(`paragraph ${paraId} is gone`);
  return paragraph.text;
}

/** A found word's range in `paraId`, taken from search so embeds ahead of it do not matter. */
export function matchRange(matches: YrsTextMatch[], paraId: string, length: number): YrsStoryRange {
  const match = matches.find((entry) => entry.story === STORY && entry.paraId === paraId);
  expect(match).toBeDefined();
  expect(match!.end - match!.start).toBe(length);
  return { story: STORY, start: { paraId, offset: match!.start }, end: { paraId, offset: match!.end } };
}

export function frameText(frame: RetainedFrame): string {
  return frame.displayList.pages
    .flatMap((page) => page.primitives)
    .map((primitive) => ('text' in primitive && typeof primitive.text === 'string' ? primitive.text : ''))
    .join('');
}

export function blocks(document: Document, type: 'paragraph' | 'table'): number {
  return document.package.document.content.filter((block) => block.type === type).length;
}

/** Project the session onto the opened package and repack it, as saving does. */
export async function save(editor: Editor, timer: Timer): Promise<Uint8Array> {
  const materialized = timer.op('materializeDocx', () => editor.session.materializeDocx());
  expect(materialized).not.toBeNull();
  const projected = timer.op('save:project', () => yrsToDocument(editor.session, materialized!));
  const bytes = await timer.opAsync('save:repack', () => repackDocx(projected));
  expect(bytes.byteLength).toBeGreaterThan(0);
  return new Uint8Array(bytes);
}

/** One collaborative session with the update bytes its own edits produced. */
export interface Replica {
  name: string;
  editor: Editor;
  timer: ActorRecorder;
  outbox: Uint8Array[];
}

export async function replicas(ctx: DocxCtx, names: string[]): Promise<Replica[]> {
  const peers: Replica[] = [];
  for (const name of names) {
    const timer = ctx.recorder.as(name);
    const editor = await timer.loadAsync(() => (peers.length === 0 ? ctx.open() : ctx.hydrate(peers[0].editor)));
    const outbox: Uint8Array[] = [];
    editor.session.onUpdate((update, origin) => {
      if (origin === 'local') outbox.push(update);
    });
    peers.push({ name, editor, timer, outbox });
  }
  return peers;
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
        to.timer.op(op, () => to.editor.session.applyUpdate(update), undefined, { from: from.name, bytes: update.byteLength });
      }
    }
  }
  return bytes;
}

/** Every body paragraph's text: the convergence fingerprint of a session. */
export function fingerprint(session: YrsSession): string {
  return session
    .paragraphs(STORY)
    .map((paragraph) => `${paragraph.paraId}=${paragraph.text}`)
    .join('|');
}

import type { YrsEngineApplyProfile, YrsLoc, YrsSession } from './index';
import type { CollaborationUpdateOrigin } from '../collaboration/types';
import { createEditSession, preloadEditWasm } from './wasm/index';

export type ResidentEngineSession = Pick<
  YrsSession,
  | 'applyDelete'
  | 'applyDeleteProfiled'
  | 'applyInput'
  | 'applyInputProfiled'
  | 'applyUpdate'
  | 'buildDisplayListFrame'
  | 'clearFonts'
  | 'destroy'
  | 'layoutDocumentJson'
  | 'layoutFontRequirementsJson'
  | 'layoutDocumentWithRegionsJson'
  | 'loadState'
  | 'measureParagraphJson'
  | 'onUpdate'
  | 'outlineGlyphJson'
  | 'registerFont'
  | 'setSelection'
  | 'yrsBlocksForStory'
>;

export async function createResidentEngineSession(): Promise<ResidentEngineSession> {
  await preloadEditWasm();
  const session = createEditSession(randomClientId());
  const listeners = new Set<
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  let observing = false;
  let destroyed = false;
  let undoStory: string | null = null;
  let selectionHead: YrsLoc | null = null;

  const ensureUndo = (story: string): void => {
    if (undoStory === story) return;
    session.track_undo(story);
    undoStory = story;
  };

  const ensureObserver = (): void => {
    if (observing) return;
    session.set_update_observer((update: Uint8Array, origin: number) => {
      if (origin !== 0 && origin !== 1) return;
      for (const listener of [...listeners]) {
        listener(update, origin === 0 ? 'local' : 'remote');
      }
    });
    observing = true;
  };

  return {
    registerFont: (bytes) => session.register_measure_font(bytes),
    clearFonts: () => session.clear_measure_fonts(),
    measureParagraphJson: (input) => session.measure_paragraph_json(input),
    layoutDocumentJson: (input) => session.layout_document_json(input),
    layoutFontRequirementsJson: (input) => session.layout_font_requirements_json(input),
    layoutDocumentWithRegionsJson: (input) => session.layout_document_with_regions_json(input),
    buildDisplayListFrame: (input, expectedFrameEpoch) =>
      session.build_display_list_frame(input, expectedFrameEpoch),
    applyInput: (text, expectedFrameEpoch) => {
      ensureUndo(selectionHead?.story ?? 'body');
      return session.apply_input(text, expectedFrameEpoch);
    },
    applyDelete: (direction, expectedFrameEpoch) => {
      ensureUndo(selectionHead?.story ?? 'body');
      return session.apply_delete(direction, expectedFrameEpoch);
    },
    applyInputProfiled: (text, expectedFrameEpoch) => {
      ensureUndo(selectionHead?.story ?? 'body');
      const frame = session.apply_input_profiled(text, expectedFrameEpoch);
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    applyDeleteProfiled: (direction, expectedFrameEpoch) => {
      ensureUndo(selectionHead?.story ?? 'body');
      const frame = session.apply_delete_profiled(direction, expectedFrameEpoch);
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    outlineGlyphJson: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),
    loadState: (update) => session.load(update),
    applyUpdate: (update) => session.apply_update(update),
    onUpdate: (listener) => {
      listeners.add(listener);
      ensureObserver();
      return () => listeners.delete(listener);
    },
    setSelection: (anchor, head = anchor) => {
      if (anchor.story !== head.story) throw new Error('yrs selection must stay inside one story');
      session.set_selection(anchor.story, anchor.paraId, anchor.offset, head.paraId, head.offset);
      selectionHead = { ...head };
    },
    yrsBlocksForStory: (story, env = {}) =>
      JSON.parse(session.yrs_blocks_for_story(story, JSON.stringify(env))) as unknown[],
    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      listeners.clear();
      if (observing) session.clear_update_observer();
      session.free();
    },
  };
}

function randomClientId(): number {
  if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
    const buffer = new Uint32Array(1);
    crypto.getRandomValues(buffer);
    return buffer[0];
  }
  return Math.floor(Math.random() * 0xffffffff);
}

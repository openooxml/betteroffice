import { transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import { isCommandEnvelope } from '../../../../shared/office-session/commands/types';
import { extendTextRange, graphemeBoundaries, textRangeAt } from '../../../../shared/textSelection';
import { SessionFailure, type MethodHandlers, type MethodPolicies } from '../../../../shared/office-session/types';
import type { PptxEditRequest, PptxEditResult, PptxStoryTarget, PptxStoryText, PptxTextField } from '../edits';
import { DEFAULT_EDITOR_TEXT_STYLE, insertionEchoRun } from '../textFormatting';
import type { ParagraphAlignment, PptxCaretAnchor, SlideDisplayList, ShapeRect, ShapeDraft, PresetShapeDraft, PictureDraft, ShapeReceipt, ShapeSnapshot, StorySnapshot, TextStyle, TextStyleSnapshot } from '../types';
import { canMoveShape, canResizeShape, effectiveShapeRect, findShape, findTopLevelShape, MIN_SHAPE_SIZE_EMU } from './commandShapes';
import {
  decodeTiffImage, initWasm, openPresentation, presentationMetadata, presentationStoryFormatting,
  presentationDeleteTextJoiningParagraphs, type PresentationCommandReadResult,
  type PresentationHandle, type PresentationHistoryReceipt,
} from '../wasm/loader';
import { buildCommandFrame } from './commandFrame';
import { prepareCommandOperation } from './commandOperations';
import { createPresentationCommandProjection,
  type PresentationSharedAck, type PresentationSharedMethods } from './commandSharedAdapter';
import { findLine, hitGeometry, indexHitGeometry, navigateGeometry, transformedRect,
  type CommandHitGeometryIndex } from './commandGeometry';
import {
  CommandOutcomeUnknown, PRESENTATION_COMMAND_ADMISSION,
  type PresentationCommandAck,
  type PresentationCommandEditState,
  type PresentationCommandEnvelope,
  type PresentationCommandFence,
  type PresentationCommandFrame,
  type PresentationCommandFrameTag,
  type PresentationCommandIdentity,
  type PresentationCommandMethods,
  type PresentationCommandOutcome,
  type PresentationCommandRefusalReason,
  type PresentationCommandSelection,
  type PresentationCommandObjectSelection,
  type PresentationCommandObjectTarget,
} from './commandProtocol';

const MIN_FONT_SIZE_PT = 1;
const MAX_FONT_SIZE_PT = 4000;

function normalizeFontSize(size: number): number {
  return Math.round(size * 100) / 100;
}

export interface PresentationCommandHostOptions {
  initWasm?: (source?: ArrayBuffer | WebAssembly.Module) => Promise<void>;
}

export interface PresentationCommandRegistration {
  handlers: MethodHandlers<PresentationCommandMethods & PresentationSharedMethods, null>;
  policies: MethodPolicies<PresentationCommandMethods & PresentationSharedMethods>;
  dispose(): void;
}

export function registerPresentationCommandHandlers(
  options: PresentationCommandHostOptions = {}
): PresentationCommandRegistration {
  type AnchoredSelection = {
    selection: PresentationCommandSelection;
    anchor: PptxCaretAnchor;
    focus: PptxCaretAnchor;
  };
  type Admission =
    | { kind: 'observer'; enabled: boolean }
    | { kind: 'invoke'; run(): unknown; result?: unknown; method: string; receipt?: PresentationHistoryReceipt }
    | { kind: 'remote'; bytes: ArrayBuffer; receipt?: PresentationHistoryReceipt }
    | { kind: 'batch'; request: PptxEditRequest; result?: PptxEditResult; receipt?: PresentationHistoryReceipt }
    | { kind: 'font'; font: { family: string; bytes: ArrayBuffer; bold?: boolean; italic?: boolean }; fallback: boolean; fontId?: number }
    | { kind: 'insertion'; slideId: string; draft: ShapeDraft | PresetShapeDraft | PictureDraft; receipt?: ShapeReceipt }
    | { kind: 'object'; object: PresentationCommandObjectSelection }
    | { kind: 'transform'; object: PresentationCommandObjectSelection; rect: ShapeRect; move: boolean }
    | { kind: 'history'; direction: 'undo' | 'redo'; receipt?: PresentationHistoryReceipt }
    | { kind: 'activation'; slideId: string }
    | { kind: 'mode'; readOnly: boolean }
    | { kind: 'selection'; selection: PresentationCommandSelection | null }
    | { kind: 'compositionStart'; id: string; selection: PresentationCommandSelection }
    | { kind: 'compositionEnd' }
    | { kind: 'noop' }
    | { kind: 'format'; selection: PresentationCommandSelection; patch: TextStyle }
    | { kind: 'alignment'; selection: PresentationCommandSelection; alignment: ParagraphAlignment | null }
    | { kind: 'text'; selection: PresentationCommandSelection; start: number; end: number;
      text: string; style?: TextStyle };
  type Mutation = { seq: number; story: string; start: number; deleted: number; inserted: number;
    version: string; beforeText: string };
  type FormattingProjection = { formatting: TextStyleSnapshot; alignment: ParagraphAlignment | null };
  type CachedStory = { text: PptxStoryText; snapshot: StorySnapshot; echoRanges: { start: number; end: number }[];
    formatting?: StorySnapshot;
    toolbar?: { story: StorySnapshot; selection: PresentationCommandSelection; marks: TextStyle;
      result: FormattingProjection };
    source: Map<string, { text: string; fields: PptxTextField[]; editable: boolean }> };
  let offUpdates: (() => void) | undefined;
  const pendingUpdates: NonNullable<PresentationCommandAck['updates']> = [];
  let handle: PresentationHandle | undefined;
  let identity: PresentationCommandIdentity | undefined;
  let metadata: ReturnType<typeof presentationMetadata> | undefined;
  let metadataVersion = '';
  const slideIndices = new Map<string, number>();
  let disposed = false;
  let opening = false;
  let throughSeq = 0;
  let version = '';
  const documentEpoch = 1;
  let frameEpoch = 0;
  let activeSlideId: string | null = null;
  let selection: PresentationCommandSelection | null = null;
  let objectSelection: PresentationCommandObjectSelection | null = null;
  let nextObjectToken = 0;
  const objectTokens = new Map<string, { signature: string; token: string }>();
  let pointerRange: { group: string; selection: AnchoredSelection } | undefined;
  let readOnly = false;
  let observeChanges = false;
  let failedSeq: number | undefined;
  const storedMarks: TextStyle = {};
  const draftStyles = new Map<string, TextStyleSnapshot>();
  let lastEditState: PresentationCommandEditState | undefined;
  let composition: { id: string; selection: PresentationCommandSelection; committed: boolean } | undefined;
  let typing: { target: string; group: string; time: number; kind: string; position: number } | undefined;
  let initialUpdate = false;
  let historyThroughSeq = 0;
  const versions = new Map<string, number>();
  const projection = createPresentationCommandProjection();
  let retainedBasis: { version: string; seq: number } | undefined;
  const outcomes = new Map<number, PresentationCommandOutcome>();
  const sharedReasons = new Map<number, string>();
  const mutations = new Map<number, Mutation>();
  const storyMutations = new Map<string, Map<number, Mutation>>();
  const stories = new Map<string, CachedStory>();
  const notes = new Map<string, string>();
  const slideStoryKeys = new Map<string, Set<string>>();
  const selections = new Map<number, AnchoredSelection | null>();
  const revisions = new Map<string, number>();
  const sentMedia = new Set<string>();
  const frames = new Map<string, {
    tag: PresentationCommandFrameTag; geometry: CommandHitGeometryIndex; version: string;
  }>();

  const layouts = new Map<string, { slideIndex: number; revision: number;
    list: SlideDisplayList; geometry: CommandHitGeometryIndex }>();

  function clearMarks(): void {
    for (const key of Object.keys(storedMarks) as (keyof TextStyle)[]) delete storedMarks[key];
  }

  function closeTyping(): void {
    if (typing) presentation().addUndoBoundary();
    typing = undefined;
  }

  function cacheStories(content: Extract<PresentationCommandReadResult, { ok: true }>): void {
    for (const slide of content.slides) notes.set(slide.id, slide.notes ?? '');
    const sources = new Map(content.sources.map((source) => [source.storyId, source.paragraphs]));
    for (const story of content.stories) {
      const key = storyKey(story);
      const keys = slideStoryKeys.get(story.slideId) ?? new Set<string>();
      keys.add(key);
      slideStoryKeys.set(story.slideId, keys);
      stories.set(key, { text: story, snapshot: presentation().story(story.storyId), echoRanges: [],
        source: new Map((sources.get(story.storyId) ?? []).map((paragraph) => {
          const live = story.paragraphs.find((value) => value.paragraphId === paragraph.paragraphId)!;
          const text = story.text.slice(live.start, live.end);
          const [prefix, suffix] = commonEnds(paragraph.text, text);
          const fields = paragraph.fields.filter((field) => {
            const shift = field.end <= prefix ? 0
              : field.start >= paragraph.text.length - suffix ? text.length - paragraph.text.length : undefined;
            return shift !== undefined && live.fields.some((current) =>
              current.start === live.start + field.start + shift && current.end === live.start + field.end + shift &&
              current.fieldType === field.fieldType);
          });
          return [paragraph.paragraphId, { text: paragraph.text, fields,
            editable: paragraph.editable && live.editable && !initialUpdate }];
        })) });
    }
  }

  function refreshHistory(seq: number, receipt: PresentationHistoryReceipt, selectAppeared = false): string[] {
    const opened = presentation();
    const dirty = new Set(receipt.changedSlideIds);
    if (receipt.deckSizeChanged || receipt.slideOrderChanged) {
      const index = metadata!.slides.findIndex((slide) => slide.id === activeSlideId);
      const previous = new Map(metadata!.slides.map((slide) => [slide.id, slide.index]));
      metadata = presentationMetadata(opened);
      metadataVersion = opened.version();
      slideIndices.clear();
      const ids = new Set(metadata.slides.map((slide) => slide.id));
      for (const slide of metadata.slides) {
        slideIndices.set(slide.id, slide.index);
        if (receipt.deckSizeChanged || previous.get(slide.id) !== slide.index) dirty.add(slide.id);
        if (!revisions.has(slide.id)) revisions.set(slide.id, 0);
      }
      for (const id of previous.keys()) if (!ids.has(id)) dirty.add(id);
      for (const id of revisions.keys()) {
        if (!ids.has(id)) { revisions.delete(id); layouts.delete(id); }
      }
      if (activeSlideId === null || !ids.has(activeSlideId)) {
        activeSlideId = metadata.slides[Math.max(0, Math.min(index, metadata.slides.length - 1))]?.id ?? null;
      }
      const appeared = selectAppeared ? metadata.slides.find((slide) => !previous.has(slide.id)) : undefined;
      if (appeared) activeSlideId = appeared.id;
    }
    const content = opened.readCommandContent({ slideIds: [...dirty].filter((id) => revisions.has(id)) });
    if (!content.ok) throw new Error(`Presentation history read refused: ${content.failure.message}`);
    for (const id of dirty) {
      for (const key of slideStoryKeys.get(id) ?? []) stories.delete(key);
      slideStoryKeys.delete(id);
    }
    cacheStories(content);
    versions.clear();
    retainedBasis = undefined;
    mutations.clear();
    storyMutations.clear();
    selections.clear();
    for (const [key, frame] of frames) {
      if (dirty.has(frame.tag.slideId)) frames.delete(key);
    }
    historyThroughSeq = seq;
    objectTokens.clear();
    return [...dirty];
  }

  function objectToken(slideId: string, shape: ShapeSnapshot): string {
    const key = JSON.stringify([slideId, shape.id]);
    const signature = JSON.stringify([shape.sourceId, shape.kind, shape.textStories.map((story) => story.id)]);
    let retained = objectTokens.get(key);
    if (!retained || retained.signature !== signature) {
      retained = { signature, token: `${identity!.incarnation}:${++nextObjectToken}` };
      objectTokens.set(key, retained);
    }
    return retained.token;
  }

  function selectedObject(slideId: string, shapeId: string): PresentationCommandObjectSelection | null {
    const deck = presentation().snapshot();
    const slide = deck.slides.find((slide) => slide.id === slideId);
    const shape = slide && findShape(slide.shapes, shapeId);
    const geometry = authoritativeGeometry(slideId);
    if (!shape || shape.id !== shapeId || !geometry) return null;
    const ids = new Set<string>();
    function collect(shape: ShapeSnapshot): void {
      ids.add(shape.id);
      for (const child of shape.children) collect(child);
    }
    collect(shape);
    const rectangles = geometry.hits.filter((entry) => entry.primitive.shapeId && ids.has(entry.primitive.shapeId))
      .map((entry) => transformedRect(entry.transforms, { x: entry.primitive.x, y: entry.primitive.y,
        width: entry.primitive.w, height: entry.primitive.h }));
    const rect = effectiveShapeRect(shape);
    const listSize = geometry.size;
    if (!listSize) return null;
    const units = { x: deck.widthEmu / listSize.width, y: deck.heightEmu / listSize.height };
    if (![units.x, units.y].every((value) => Number.isFinite(value) && value > 0)) return null;
    let bounds: ShapeRect | null = null;
    if (rectangles.length) {
      const x = Math.min(...rectangles.map((rect) => rect.x));
      const y = Math.min(...rectangles.map((rect) => rect.y));
      bounds = { x, y, width: Math.max(...rectangles.map((rect) => rect.x + rect.width)) - x,
        height: Math.max(...rectangles.map((rect) => rect.y + rect.height)) - y };
    } else if (rect) bounds = { x: rect.x / units.x, y: rect.y / units.y,
      width: rect.width / units.x, height: rect.height / units.y };
    if (!bounds) return null;
    const story = shape.textStories[0];
    return { kind: 'object', slideId, shapeId, storyId: story?.id ?? '',
      anchor: 0, focus: 0, affinity: 'downstream', token: objectToken(slideId, shape), bounds, rect,
      canMove: slide!.shapes.includes(shape) && canMoveShape(shape),
      canResize: slide!.shapes.includes(shape) && canResizeShape(shape),
      aspectRatioLocked: shape.aspectRatioLocked ?? false, units };
  }

  function objectStory(object: PresentationCommandObjectSelection): PresentationCommandSelection | null {
    const story = stories.get(storyKey(object))?.text;
    return story ? { slideId: object.slideId, shapeId: object.shapeId, storyId: object.storyId,
      kind: 'shape', anchor: 0, focus: story.text.length, affinity: 'downstream' } : null;
  }

  function objectFrame(tag: PresentationCommandFrameTag): boolean {
    return Boolean(tag && tag.slideId === activeSlideId && tag.incarnation === identity!.incarnation &&
      tag.documentEpoch === documentEpoch && tag.slideRevision === revisions.get(tag.slideId) &&
      tag.coversThroughSeq <= throughSeq && tag.coversThroughSeq >= historyThroughSeq && frames.has(frameKey(tag)));
  }

  function transformTarget(value: PresentationCommandObjectTarget): PresentationCommandObjectSelection | null {
    if (!value || typeof value.slideId !== 'string' || typeof value.shapeId !== 'string' ||
      typeof value.token !== 'string') return null;
    const object = selectedObject(value.slideId, value.shapeId);
    return object?.token === value.token && objectSelection?.token === value.token &&
      objectSelection.slideId === value.slideId && objectSelection.shapeId === value.shapeId ? object : null;
  }

  function marksAt(at: PresentationCommandSelection): TextStyle {
    return selection && storyKey(selection) === storyKey(at) && selection.anchor === at.anchor &&
      selection.focus === at.focus && selection.affinity === at.affinity && selection.kind === at.kind ? storedMarks : {};
  }

  function retainVersions(): void {
    projection.retain([...versions.keys(), ...[...frames.values()].map((entry) => entry.version),
      ...(retainedBasis ? [retainedBasis.version] : []), version]);
    for (const seq of sharedReasons.keys()) {
      if (!outcomes.has(seq)) sharedReasons.delete(seq);
    }
  }

  function sharedAck(result: TransferResult<PresentationCommandAck>): TransferResult<PresentationSharedAck> {
    const value = projection.ack(result.value, retainedBasis?.version);
    value.outcomes = value.outcomes.map((outcome) => outcome.reason === 'stale-basis' && sharedReasons.has(outcome.seq)
      ? { ...outcome, reason: sharedReasons.get(outcome.seq) } : outcome);
    return transferable(value, result.transfer);
  }

  function frameKey(value: PresentationCommandFrameTag): string {
    return JSON.stringify([value.slideId, value.slideRevision, value.coversThroughSeq,
      value.documentEpoch, value.incarnation]);
  }

  function presentation(): PresentationHandle {
    if (failedSeq !== undefined) throw new CommandOutcomeUnknown(failedSeq, 'Presentation command session has failed');
    if (disposed) throw new Error('Presentation command session is disposed');
    if (!handle || !identity || !metadata) throw new Error('Presentation command session is not open');
    return handle;
  }

  function matches(value: PresentationCommandIdentity): boolean {
    return value?.session === identity?.session && value?.incarnation === identity?.incarnation;
  }

  function fence(value: PresentationCommandFence): void {
    presentation();
    if (!matches(value) || !Number.isSafeInteger(value.seq) || value.seq < 0 || value.seq > throughSeq) {
      throw new Error('Presentation command fence differs');
    }
  }

  function anchor(value: PresentationCommandSelection): AnchoredSelection {
    const opened = presentation();
    return { selection: { ...value }, anchor: opened.anchorCaret(value.storyId, value.anchor),
      focus: opened.anchorCaret(value.storyId, value.focus) };
  }

  function resolve(value: AnchoredSelection): PresentationCommandSelection | null {
    const opened = presentation();
    const anchor = opened.resolveCaretAnchor(value.anchor);
    const focus = opened.resolveCaretAnchor(value.focus);
    if (anchor === null || focus === null) return null;
    const result = { ...value.selection, anchor, focus };
    if (anchor !== value.selection.anchor || focus !== value.selection.focus) delete result.goalX;
    return result;
  }

  function target(value: PresentationCommandSelection): PptxStoryText | undefined {
    if (!value || typeof value.slideId !== 'string' ||
      typeof value.shapeId !== 'string' || typeof value.storyId !== 'string') return undefined;
    return stories.get(storyKey(value))?.text;
  }

  function refreshStory(value: PresentationCommandSelection, start: number, end: number): void {
    const cached = stories.get(storyKey(value))!;
    for (const paragraph of cached.text.paragraphs) {
      const source = cached.source.get(paragraph.paragraphId);
      if (source) source.fields = source.fields.filter((_, index) => {
        const field = paragraph.fields[index]!;
        return field.end <= start || field.start >= end;
      });
    }
    const snapshot = presentation().story(value.storyId);
    let text = '';
    const paragraphs = snapshot.paragraphs.map((paragraph, index) => {
      if (index > 0) text += '\n';
      const start = text.length;
      const live = paragraph.runs.map((run) => run.text).join('');
      text += live;
      const source = cached.source.get(paragraph.id);
      const fields: PptxTextField[] = [];
      let editable = source?.editable ?? true;
      if (source && source.fields.length > 0) {
        const [prefix, suffix] = commonEnds(source.text, live);
        for (const field of source.fields) {
          const shift = field.end <= prefix ? 0
            : field.start >= source.text.length - suffix ? live.length - source.text.length : undefined;
          if (shift === undefined) {
            editable = false;
            fields.length = 0;
            break;
          }
          fields.push({ ...field, start: start + field.start + shift, end: start + field.end + shift });
        }
      }
      const lineBreaks: number[] = [];
      for (let offset = start; offset < text.length; offset += 1) {
        if (text[offset] === '\n') lineBreaks.push(offset);
      }
      return { paragraphId: paragraph.id, start, end: text.length, lineBreaks, fields, editable };
    });
    if (text.length + 1 !== snapshot.length) throw new Error('Story text projection differs');
    const inserted = text.length - cached.text.text.length + end - start;
    if (start !== end) cached.echoRanges = [];
    else cached.echoRanges = cached.echoRanges.map((range) => ({
      start: range.start >= start ? range.start + inserted : range.start,
      end: range.end >= start ? range.end + inserted : range.end,
    }));
    if (inserted > 0) {
      cached.echoRanges.push({ start, end: start + inserted });
      cached.echoRanges.sort((a, b) => a.start - b.start);
      cached.echoRanges = cached.echoRanges.reduce<typeof cached.echoRanges>((ranges, range) => {
        const last = ranges[ranges.length - 1];
        if (last && last.end >= range.start) last.end = Math.max(last.end, range.end);
        else ranges.push({ ...range });
        return ranges;
      }, []);
    }
    cached.snapshot = snapshot;
    cached.formatting = undefined;
    cached.text = { ...cached.text, text, paragraphs };
  }

  function formattingAt(at: PresentationCommandSelection): FormattingProjection {
    const cached = stories.get(storyKey(at))!;
    const story = cached.formatting ??= presentationStoryFormatting(presentation(), at);
    const marks = marksAt(at);
    // Boxes inserted here type at their draft style, not the inherited cascade of their empty paragraphs.
    const draft = draftStyles.get(storyKey(at));
    const previous = cached.toolbar;
    if (previous?.story === story && previous.selection.anchor === at.anchor && previous.selection.focus === at.focus &&
      previous.selection.affinity === at.affinity && previous.selection.kind === at.kind &&
      Object.keys(previous.marks).length === Object.keys(marks).length &&
      (Object.keys(marks) as (keyof TextStyle)[]).every((key) => previous.marks[key] === marks[key])) return previous.result;
    const start = at.kind === 'shape' ? 0 : Math.min(at.anchor, at.focus);
    const end = at.kind === 'shape' ? story.length : Math.max(at.anchor, at.focus);
    let offset = 0;
    let caretStyle: TextStyleSnapshot | undefined;
    let precedingStyle: TextStyleSnapshot | undefined;
    let firstStyle: TextStyleSnapshot | undefined;
    const styles: TextStyleSnapshot[] = [];
    const emptyStyles: TextStyleSnapshot[] = [];
    const alignments = new Set<string | null>();
    for (const paragraph of story.paragraphs) {
      const paragraphStart = offset;
      for (const run of paragraph.runs) {
        if (draft && run.text.length === 0) continue;
        const runEnd = offset + run.text.length;
        const style: TextStyleSnapshot = { ...run.style,
          bold: run.style.bold ?? DEFAULT_EDITOR_TEXT_STYLE.bold,
          italic: run.style.italic ?? DEFAULT_EDITOR_TEXT_STYLE.italic,
          underline: run.style.underline ?? 'none',
          fontFamily: run.style.fontFamily ?? DEFAULT_EDITOR_TEXT_STYLE.fontFamily,
          fontSizePt: normalizeFontSize(run.style.fontSizePt ?? DEFAULT_EDITOR_TEXT_STYLE.fontSizePt),
          color: run.style.color ?? DEFAULT_EDITOR_TEXT_STYLE.color };
        firstStyle ??= style;
        if (start >= offset && start < runEnd || start === offset && runEnd === offset) caretStyle ??= style;
        if (runEnd <= start && runEnd > offset) precedingStyle = style;
        if (runEnd > offset && start < runEnd && end > offset) styles.push(style);
        if (runEnd === offset && start <= offset && end > offset) emptyStyles.push(style);
        offset = runEnd;
      }
      if (start === end ? start >= paragraphStart && start <= offset
        : start <= offset && end > paragraphStart) alignments.add(paragraph.alignment ?? 'l');
      offset += 1;
    }
    styles.push(...emptyStyles);
    if (draft && styles.length === 0) styles.push(draft);
    let formatting: TextStyleSnapshot = { ...DEFAULT_EDITOR_TEXT_STYLE };
    if (start === end && at.kind !== 'shape') {
      formatting = { ...formatting, ...(caretStyle ?? precedingStyle ??
        (draft ? firstStyle ?? draft : story.paragraphs[0]?.runs[0]?.style)), ...marks };
      if (formatting.fontSizePt !== null && formatting.fontSizePt !== undefined) {
        formatting.fontSizePt = normalizeFontSize(formatting.fontSizePt);
      }
    } else if (styles.length > 0) {
      formatting = { ...styles[0]! };
      const keys = new Set(styles.flatMap((style) => Object.keys(style))) as Set<keyof TextStyleSnapshot>;
      for (const key of keys) {
        if (styles.some((style) => style[key] !== styles[0]![key])) formatting[key] = null;
      }
    }
    const value = alignments.size === 1 ? [...alignments][0] : null;
    const result: FormattingProjection = { formatting,
      alignment: value === 'l' || value === 'ctr' || value === 'r' || value === 'just' ? value : null };
    cached.toolbar = { story, selection: { ...at }, marks: { ...marks }, result };
    return result;
  }

  function editState(): PresentationCommandEditState {
    const opened = presentation();
    const { formatting, alignment } = selection ? formattingAt(selection) : { formatting: null, alignment: null };
    let echoRun: 'before' | 'after' | null = null;
    if (selection && selection.kind !== 'shape') {
      const cached = stories.get(storyKey(selection))!;
      const neighbour = insertionEchoRun(cached.snapshot, selection.anchor, selection.focus, storedMarks);
      if (!readOnly && neighbour && cached.echoRanges.some((range) => range.start <= neighbour.start && range.end >= neighbour.end)
        && cached.text.paragraphs.some((paragraph) => paragraph.paragraphId === neighbour.paragraphId
          && paragraph.editable && paragraph.fields.length === 0)) echoRun = neighbour.neighbour;
    }
    const slide = objectSelection ? opened.snapshot().slides.find((slide) => slide.id === objectSelection!.slideId) : null;
    const shape = slide && slide.shapes.find((shape) => shape.id === objectSelection!.shapeId);
    const color = (value: string | null | undefined) => value && /^[#]?[0-9a-f]{6}$/i.test(value) ? `#${value.replace(/^#/, '')}` : null;
    lastEditState = { readOnly, proposals: opened.listProposals(), notes: activeSlideId ? { slideId: activeSlideId, text: notes.get(activeSlideId) ?? '' } : null,
      object: shape && slide ? { formattable: shape.kind === 'shape', geometry: shape.geometry ?? null,
        fill: shape.fill?.type === 'none' ? null : color(shape.resolvedFillColor), stroke: color(shape.resolvedOutlineColor),
        strokeWidth: shape.resolvedOutlineColor ? (shape.outline?.width ?? 12700) / 12700 : null,
        adjustments: shape.adjustValues, order: { index: slide.shapes.findIndex((value) => value.id === shape.id), count: slide.shapes.length } } : null, canUndo: opened.canUndo(), canRedo: opened.canRedo(),
      toolbar: { target: selection ? { slideId: selection.slideId, shapeId: selection.shapeId,
        storyId: selection.storyId } : null,
      textContext: selection ? selection.kind === 'shape' ? 'shape'
        : selection.anchor === selection.focus ? 'caret' : 'range' : null,
      storedMarks: { ...storedMarks }, formatting, alignment,
      hyperlink: null, table: null, canEditText: selection !== null && !readOnly, echoRun } };
    return lastEditState;
  }

  function refreshMetadata(): ReturnType<typeof presentationMetadata> {
    metadata = presentationMetadata(handle!);
    metadataVersion = handle!.version();
    slideIndices.clear();
    for (const [index, slide] of metadata.slides.entries()) slideIndices.set(slide.id, index);
    return metadata;
  }

  function refreshTopology(): void {
    if (presentation().version() !== metadataVersion) refreshMetadata();
  }

  function frame(slideId: string, retain = true): TransferResult<PresentationCommandFrame> {
    const opened = presentation();
    refreshTopology();
    const slideIndex = slideIndices.get(slideId);
    if (slideIndex === undefined) throw new RangeError('Slide is not in the command session');
    const revision = revisions.get(slideId)!;
    const cached = layouts.get(slideId);
    const layout = cached?.revision === revision && cached.slideIndex === slideIndex ? cached : undefined;
    let geometry = layout?.geometry;
    const result = buildCommandFrame(opened, { slideId, slideIndex,
      slideRevision: revision, coversThroughSeq: throughSeq,
      documentEpoch, incarnation: identity!.incarnation, version, epoch: ++frameEpoch }, sentMedia,
    layout ? undefined : (list) => {
      geometry = indexHitGeometry(list);
      layouts.set(slideId, { slideIndex, revision, list, geometry });
    }, layout?.list);
    if (!geometry || !retain) return result;
    const { slideId: id, slideRevision, coversThroughSeq, documentEpoch: epoch, incarnation } = result.value;
    const key = frameKey(result.value);
    frames.delete(key);
    frames.set(key, { tag: { slideId: id, slideRevision, coversThroughSeq, documentEpoch: epoch, incarnation },
      geometry, version });
    while (frames.size > 257) frames.delete(frames.keys().next().value!);
    return result;
  }

  function authoritativeGeometry(slideId: string): CommandHitGeometryIndex | undefined {
    let geometry = [...frames.values()].find((entry) => entry.tag.slideId === slideId &&
      entry.tag.coversThroughSeq === throughSeq && entry.tag.slideRevision === revisions.get(slideId))?.geometry;
    if (!geometry) {
      refreshTopology();
      const slideIndex = slideIndices.get(slideId);
      if (slideIndex === undefined) return undefined;
      frame(slideId);
      geometry = layouts.get(slideId)?.geometry;
    }
    return geometry;
  }

  function ack(
    commandOutcomes: PresentationCommandOutcome[] = [], dirtySlideIds: string[] = [], history = false, creation?: ShapeReceipt, editResult?: PptxEditResult, fontId?: number
  ): TransferResult<PresentationCommandAck> {
    if (objectSelection) {
      const object = selectedObject(objectSelection.slideId, objectSelection.shapeId);
      objectSelection = object?.token === objectSelection.token ? object : null;
      selection = objectSelection ? objectStory(objectSelection) : null;
    }
    const edit = editState();
    refreshTopology();
    const result = activeSlideId === null || !slideIndices.has(activeSlideId) ? null : frame(activeSlideId);
    retainVersions();
    const updates = pendingUpdates.splice(0);
    const reply = transferable<PresentationCommandAck>({ ...identity!, epoch: documentEpoch, throughSeq, version,
      ...(updates.length ? { updates } : {}),
      ...(observeChanges && commandOutcomes.some((outcome) => outcome.committed) ? { change: { ...identity!, throughSeq, version, snapshot: presentation().snapshot() } } : {}),
      selection: objectSelection ?? (selection ? { ...selection } : null), editState: edit, outcomes: commandOutcomes,
      dirtySlideIds, activeSlideId, historyThroughSeq, ...(creation ? { creation } : {}), ...(editResult ? { editResult } : {}), ...(fontId !== undefined ? { fontId } : {}),
      ...(history ? { slides: metadata!.slides, size: metadata!.size } : {}),
      frame: result?.value ?? null }, [...(result?.transfer ?? []), ...updates.map((update) => update.bytes)]);
    for (const { assetId } of reply.value.frame?.media ?? []) sentMedia.add(assetId);
    return reply;
  }

  function admit(command: PresentationCommandEnvelope): Admission | PresentationCommandRefusalReason {
    if (!command.basis || command.basis.epoch !== documentEpoch ||
      typeof command.basis.version !== 'string' || command.basis.version.length === 0) return 'stale-basis';
    const payload = command.payload;
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return 'invalid-payload';
    if (command.kind !== 'activateSlide' && (!Number.isFinite(command.payload.inputTime) ||
      typeof command.payload.undoGroup !== 'string' || command.payload.undoGroup.length === 0)) return 'invalid-payload';
    const admission = PRESENTATION_COMMAND_ADMISSION[command.kind];
    if (!admission) return 'invalid-payload';
    if (command.afterSeq !== undefined && (!Number.isSafeInteger(command.afterSeq) ||
      command.afterSeq < 1 || command.afterSeq >= command.seq)) return 'invalid-payload';
    if (command.kind === 'selectAtPoint') {
      const { frame: tag, x, y, extend, granularity, drag } = command.payload;
      if (!tag || typeof tag !== 'object' || Array.isArray(tag) ||
        typeof tag.slideId !== 'string' || tag.slideId.length === 0 ||
        typeof tag.incarnation !== 'string' || tag.incarnation.length === 0 ||
        !Number.isSafeInteger(tag.slideRevision) || tag.slideRevision < 0 ||
        !Number.isSafeInteger(tag.coversThroughSeq) || tag.coversThroughSeq < 0 ||
        !Number.isSafeInteger(tag.documentEpoch) || tag.documentEpoch < 1 ||
        !Number.isFinite(x) || !Number.isFinite(y) || typeof extend !== 'boolean' ||
        granularity !== undefined && !['character', 'word', 'paragraph'].includes(granularity) ||
        drag !== undefined && typeof drag !== 'boolean') return 'invalid-payload';
    }
    const reference = 'at' in payload ? payload.at : undefined;
    if ('at' in payload && (!reference || typeof reference !== 'object' || Array.isArray(reference) ||
      reference.source !== 'confirmed' && reference.source !== 'afterSeq')) return 'invalid-payload';
    if (reference?.source === 'confirmed') {
      const at = reference.selection;
      if (!at || typeof at !== 'object' || Array.isArray(at) || typeof at.slideId !== 'string' ||
        typeof at.shapeId !== 'string' || typeof at.storyId !== 'string' ||
        typeof at.anchor !== 'number' || typeof at.focus !== 'number' ||
        at.affinity !== 'upstream' && at.affinity !== 'downstream') return 'invalid-payload';
      if (!Number.isSafeInteger(at.anchor) || !Number.isSafeInteger(at.focus) ||
        at.anchor < 0 || at.focus < 0) return 'invalid-offset';
    }
    if (reference?.source === 'afterSeq' && (!Number.isSafeInteger(reference.seq) || reference.seq < 1)) return 'invalid-payload';
    if (reference?.source === 'afterSeq' ? command.afterSeq !== reference.seq
      : command.afterSeq !== undefined) return 'invalid-payload';
    if (command.afterSeq !== undefined) {
      if (!outcomes.has(command.afterSeq)) return 'stale-basis';
      const dependency = outcomes.get(command.afterSeq)!;
      if (dependency.status === 'refused' || dependency.status === 'cancelled') return 'dependency-refused';
    }
    if (readOnly && admission === 'local-write') return 'read-only';
    const retainedSeq = retainedBasis && command.basis.version === retainedBasis.version
      ? retainedBasis.seq : versions.get(command.basis.version);
    const predecessorRelative = reference?.source === 'afterSeq' &&
      (command.kind !== 'setSelection' || command.payload.selection === null);
    const basisSeq = retainedSeq ?? (predecessorRelative ? throughSeq : undefined);
    if (basisSeq === undefined) return 'stale-basis';
    if (command.kind === 'activateSlide') {
      refreshMetadata();
      const { slideId } = command.payload;
      return typeof slideId === 'string' && revisions.has(slideId) && slideIndices.has(slideId)
        ? { kind: 'activation', slideId } : 'bad-target';
    }
    if (command.kind === 'setChangeObserver') return typeof command.payload.enabled === 'boolean' ? { kind: 'observer', enabled: command.payload.enabled } : 'invalid-payload';
    if (command.kind === 'invoke') {
      if (composition) return 'invalid-composition';
      if (command.basis.version !== version) return 'stale-basis';
      const run = prepareCommandOperation(presentation(), command.payload);
      return run ? { kind: 'invoke', run, method: command.payload.method } : 'invalid-payload';
    }
    if (command.kind === 'applyRemoteUpdate') {
      const bytes = command.payload.bytes;
      return bytes instanceof ArrayBuffer && bytes.byteLength > 0 && bytes.byteLength <= 16 * 1024 * 1024
        ? { kind: 'remote', bytes } : 'invalid-payload';
    }
    if (command.kind === 'applyEdits') {
      if (composition) return 'invalid-composition';
      try { presentation().validateEdits(command.payload.request); }
      catch { return 'invalid-payload'; }
      return { kind: 'batch', request: command.payload.request };
    }
    if (command.kind === 'registerFont') {
      const { font, fallback } = command.payload;
      if (!font || !validText(font.family) || !font.family || !(font.bytes instanceof ArrayBuffer) ||
        font.bytes.byteLength === 0 || font.bytes.byteLength > 32 * 1024 * 1024 ||
        font.bold !== undefined && typeof font.bold !== 'boolean' || font.italic !== undefined && typeof font.italic !== 'boolean' ||
        fallback !== undefined && typeof fallback !== 'boolean') return 'invalid-payload';
      return { kind: 'font', font, fallback: fallback ?? false };
    }
    if (command.kind === 'insertTextBox' || command.kind === 'insertShape' || command.kind === 'insertPicture') {
      if (composition) return 'invalid-composition';
      const { slideId } = command.payload;
      if (typeof slideId !== 'string' || !slideIndices.has(slideId) || slideId !== activeSlideId) return 'bad-target';
      if (command.basis.version !== version || presentation().version() !== version) return 'stale-basis';
      const draft = command.kind === 'insertPicture' ? command.payload : command.payload.draft;
      if (!draft || !validText(draft.name) || !draft.rect ||
        ![draft.rect.x, draft.rect.y, draft.rect.width, draft.rect.height].every(Number.isSafeInteger) ||
        draft.rect.width <= 0 || draft.rect.height <= 0) return 'invalid-payload';
      if (command.kind === 'insertTextBox') {
        const draft = command.payload.draft;
        if (!validText(draft.text) || !validStyle(draft.style)) return 'invalid-payload';
        return { kind: 'insertion', slideId, draft };
      }
      if (command.kind === 'insertShape') {
        const draft = command.payload.draft;
        if (!['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid',
          'pentagon', 'hexagon', 'octagon', 'star5', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'chevron'].includes(draft.geometry) ||
          draft.fill !== undefined && draft.fill !== null && !/^#[0-9a-f]{6}$/i.test(draft.fill)) return 'invalid-payload';
        return { kind: 'insertion', slideId, draft };
      }
      const { bytes, contentType } = command.payload;
      if (!(bytes instanceof ArrayBuffer) || bytes.byteLength === 0 || bytes.byteLength > 8 * 1024 * 1024 ||
        !['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff', 'image/webp', 'image/svg+xml'].includes(contentType)) return 'invalid-payload';
      let binary = '';
      const data = new Uint8Array(bytes);
      for (let offset = 0; offset < data.length; offset += 8192) binary += String.fromCharCode(...data.subarray(offset, offset + 8192));
      return { kind: 'insertion', slideId, draft: { name: draft.name, rect: draft.rect, contentType, mediaBase64: btoa(binary) } };
    }
    if (reference?.source === 'confirmed') {
      const at = reference.selection;
      const story = target(at);
      if (!story) return 'bad-target';
      let text = story.text;
      for (const mutation of storyMutations.get(storyKey(at))?.values() ?? []) {
        if (mutation.seq > basisSeq) { text = mutation.beforeText; break; }
      }
      if (!boundary(text, at.anchor) || !boundary(text, at.focus)) return 'invalid-offset';
    }
    if (retainedSeq !== undefined) {
      prune(basisSeq, command.basis.version, command.afterSeq);
    }
    if (command.kind === 'undo' || command.kind === 'redo') {
      return { kind: 'history', direction: command.kind };
    }
    if (command.kind === 'selectObject') {
      const { frame: tag } = command.payload;
      if (!objectFrame(tag) || command.basis.version !== version || presentation().version() !== version) return 'stale-basis';
      if (composition) return 'invalid-composition';
      if ('shapeId' in command.payload) {
        if (typeof command.payload.shapeId !== 'string') return 'invalid-payload';
        const object = selectedObject(tag.slideId, command.payload.shapeId);
        return object ? { kind: 'object', object } : 'bad-target';
      }
      const { x, y, scale } = command.payload;
      if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(scale) || scale <= 0) return 'invalid-payload';
      const geometry = frames.get(frameKey(tag))!.geometry;
      const hit = hitGeometry(geometry, { x, y }, undefined, scale);
      if (hit?.kind !== 'shape') return 'bad-target';
      const slide = presentation().snapshot().slides.find((slide) => slide.id === tag.slideId);
      const shape = slide && findTopLevelShape(slide, hit.shapeId);
      const object = shape ? selectedObject(tag.slideId, shape.id) : null;
      return object ? { kind: 'object', object } : 'bad-target';
    }
    if (command.kind === 'moveShape' || command.kind === 'setShapeRect') {
      const relative = command.kind === 'moveShape' && 'delta' in command.payload ? command.payload : null;
      if (relative && (relative.at.source !== 'afterSeq' || relative.at.seq !== throughSeq ||
        !relative.delta || !Number.isFinite(relative.delta.x) || !Number.isFinite(relative.delta.y))) return 'invalid-payload';
      const object = relative ? objectSelection?.shapeId === relative.shapeId
        ? transformTarget(objectSelection) : null : 'target' in command.payload ? transformTarget(command.payload.target) : null;
      if (!object) return 'bad-target';
      if (!objectFrame(command.payload.frame) || command.payload.frame.slideId !== object.slideId ||
        command.basis.version !== version || presentation().version() !== version) return 'stale-basis';
      if (composition) return 'invalid-composition';
      const move = command.kind === 'moveShape';
      if (!object.rect || !object.canMove || !move && !object.canResize) return 'unsupported';
      const rect = command.kind === 'moveShape' ? relative
        ? { ...object.rect, x: object.rect.x + Math.round(relative.delta.x * object.units.x),
          y: object.rect.y + Math.round(relative.delta.y * object.units.y) }
        : 'x' in command.payload ? { ...object.rect, x: command.payload.x, y: command.payload.y } : null
        : command.payload.rect;
      if (!rect || ![rect.x, rect.y, rect.width, rect.height].every(Number.isSafeInteger) ||
        rect.width <= 0 || rect.height <= 0 || !move &&
        (rect.width !== object.rect.width && rect.width < MIN_SHAPE_SIZE_EMU ||
          rect.height !== object.rect.height && rect.height < MIN_SHAPE_SIZE_EMU)) return 'invalid-payload';
      return { kind: 'transform', object, rect: { ...rect }, move };
    }
    let pointBasis: number | undefined;
    let point: PresentationCommandSelection | undefined;
    if (command.kind === 'selectAtPoint') {
      if (!reference && command.payload.extend) return 'invalid-payload';
      const { frame: tag, x, y } = command.payload;
      if (tag.slideId !== activeSlideId || tag.incarnation !== identity!.incarnation ||
        tag.documentEpoch !== documentEpoch || !revisions.has(tag.slideId) ||
        tag.slideRevision > revisions.get(tag.slideId)! || tag.coversThroughSeq > throughSeq ||
        tag.coversThroughSeq < historyThroughSeq) return 'stale-basis';
      const retained = frames.get(frameKey(tag));
      const geometry = retained?.geometry ?? authoritativeGeometry(tag.slideId);
      if (!geometry) return 'stale-basis';
      const previous = reference?.source === 'afterSeq' ? selections.get(reference.seq) : undefined;
      const dragTarget = command.payload.drag ? reference?.source === 'confirmed' ? reference.selection
        : previous ? resolve(previous) : null : undefined;
      const hit = hitGeometry(geometry, { x, y }, dragTarget);
      if (hit?.kind !== 'text') return { kind: 'noop' };
      point = { slideId: tag.slideId, shapeId: hit.shapeId, storyId: hit.storyId,
        anchor: hit.position, focus: hit.position, affinity: hit.affinity ?? 'downstream' };
      pointBasis = retained ? tag.coversThroughSeq : throughSeq;
    }
    if (command.kind === 'setReadOnly') {
      return typeof command.payload.readOnly === 'boolean'
        ? { kind: 'mode', readOnly: command.payload.readOnly } : 'invalid-payload';
    }
    if (command.kind === 'compositionEnd') {
      const { compositionId, cancelled } = command.payload;
      if (typeof compositionId !== 'string' ||
        (cancelled !== undefined && typeof cancelled !== 'boolean')) return 'invalid-payload';
      return composition?.id === compositionId ? { kind: 'compositionEnd' } : 'invalid-composition';
    }
    let at: PresentationCommandSelection | null;
    if (command.kind === 'setSelection') {
      at = command.payload.selection;
      if (at === null) return { kind: 'selection', selection: null };
    } else if (point && !reference) {
      at = null;
    } else {
      if (!('at' in payload) || !payload.at) return 'invalid-payload';
      const reference = payload.at;
      if (reference.source === 'confirmed') at = reference.selection;
      else if (reference.source === 'afterSeq') {
        if (reference.seq !== command.afterSeq) return 'invalid-payload';
        if (!selections.has(reference.seq)) return 'stale-basis';
        const previous = selections.get(reference.seq);
        at = previous ? resolve(previous) : null;
      } else return 'invalid-payload';
    }
    const extension = at;
    if (point) at = point;
    if (!at) return 'bad-target';
    const story = target(at);
    if (!story) return 'bad-target';
    if (!Number.isSafeInteger(at.anchor) || !Number.isSafeInteger(at.focus) ||
      at.anchor < 0 || at.focus < 0) return 'invalid-offset';
    if (point || command.kind === 'setSelection' || reference?.source === 'confirmed') {
      const fromSeq = pointBasis ?? basisSeq;
      const records = storyMutations.get(storyKey(at));
      let basisText = story.text;
      for (const mutation of records?.values() ?? []) {
        if (mutation.seq > fromSeq) { basisText = mutation.beforeText; break; }
      }
      if (!boundary(basisText, at.anchor) || !boundary(basisText, at.focus)) return 'invalid-offset';
      at = { ...at };
      for (const mutation of records?.values() ?? []) {
        if (mutation.seq <= fromSeq) continue;
        delete at.goalX;
        const start = Math.min(at.anchor, at.focus);
        const end = Math.max(at.anchor, at.focus);
        if (start !== end && mutation.deleted > 0 &&
          start < mutation.start + mutation.deleted && end > mutation.start) return 'stale-basis';
        const map = (position: number): number => {
          if (position < mutation.start) return position;
          if (mutation.deleted === 0 && position === mutation.start) {
            return position + (start === end || position === start ? mutation.inserted : 0);
          }
          return position >= mutation.start + mutation.deleted
            ? position - mutation.deleted + mutation.inserted
            : mutation.start + (start === end ? mutation.inserted : 0);
        };
        at.anchor = map(at.anchor);
        at.focus = map(at.focus);
      }
    }
    if (!boundary(story.text, at.anchor) || !boundary(story.text, at.focus)) return 'invalid-offset';
    if (at.kind !== undefined && at.kind !== 'shape') return 'invalid-payload';
    if (at.kind === 'shape') at = { ...at, anchor: 0, focus: story.text.length };
    if (at.affinity !== 'upstream' && at.affinity !== 'downstream') return 'invalid-payload';
    if (at.goalX !== undefined && !Number.isFinite(at.goalX)) return 'invalid-payload';
    if (command.kind === 'setSelection') return { kind: 'selection', selection: { ...at } };
    if (command.kind === 'selectAtPoint') {
      const boundaries = [...graphemeBoundaries(story.text)];
      const focus = at.focus;
      at.focus = boundaries.reduce((best, offset) =>
        Math.abs(offset - focus) < Math.abs(best - focus) ? offset : best, 0);
      at.anchor = at.focus;
      const granularity = command.payload.granularity ?? 'character';
      if (granularity !== 'character') {
        const range = textRangeAt(story.text, at.focus, granularity);
        if (command.payload.drag) {
          const initial = pointerRange?.group === command.payload.undoGroup ? resolve(pointerRange.selection) : null;
          if (!initial || storyKey(initial) !== storyKey(at)) return 'bad-target';
          at = { ...at, ...extendTextRange({ start: Math.min(initial.anchor, initial.focus),
            end: Math.max(initial.anchor, initial.focus) }, range) };
        } else at = { ...at, anchor: range.start, focus: range.end };
        const geometry = authoritativeGeometry(at.slideId);
        const entry = geometry && findLine(geometry, { ...at,
          affinity: at.focus > at.anchor ? 'upstream' : 'downstream' });
        if (entry) at.affinity = at.focus === entry.line.end ? 'upstream' : 'downstream';
      } else if (command.payload.extend && extension && extension.kind !== 'shape' && storyKey(extension) === storyKey(at)) {
        let text = story.text;
        if (reference?.source === 'confirmed') {
          for (const mutation of storyMutations.get(storyKey(extension))?.values() ?? []) {
            if (mutation.seq > basisSeq) { text = mutation.beforeText; break; }
          }
        }
        if (!Number.isSafeInteger(extension.anchor) || !boundary(text, extension.anchor)) return 'invalid-offset';
        const anchor = reference?.source === 'confirmed'
          ? mapConfirmedPoint(extension, basisSeq) : extension;
        at.anchor = anchor.anchor;
      }
      return { kind: 'selection', selection: at };
    }
    if (command.kind === 'moveSelection') {
      if (!['left', 'right', 'wordLeft', 'wordRight', 'lineStart', 'lineEnd', 'storyStart', 'storyEnd', 'up', 'down'].includes(command.payload.motion) ||
        typeof command.payload.extend !== 'boolean') return 'invalid-payload';
      const geometry = authoritativeGeometry(at.slideId);
      if (!geometry) return 'bad-target';
      const moved = navigateGeometry(geometry, at,
        command.payload.motion, command.payload.extend, story.text);
      return moved ? { kind: 'selection', selection: moved } : 'bad-target';
    }
    if (command.kind === 'formatText' || command.kind === 'setParagraphAlignment') {
      const start = Math.min(at.anchor, at.focus);
      const end = Math.max(at.anchor, at.focus);
      if (story.paragraphs.some((paragraph) => start <= paragraph.end && end >= paragraph.start &&
        (!paragraph.editable || paragraph.fields.some((field) =>
          start > field.start && start < field.end || end > field.start && end < field.end)))) return 'unsupported';
      if (composition) return 'invalid-composition';
      if (command.kind === 'setParagraphAlignment') {
        const alignment = command.payload.alignment;
        return alignment === null || ['l', 'ctr', 'r', 'just'].includes(alignment)
          ? { kind: 'alignment', selection: { ...at }, alignment } : 'invalid-payload';
      }
      const { patch, toggle, fontSizeStep } = command.payload;
      if (patch === undefined || !validStyle(patch) || toggle !== undefined &&
        !['bold', 'italic', 'underline'].includes(toggle)) return 'invalid-payload';
      const resolved: TextStyle = Object.fromEntries(Object.entries(patch).filter(([, value]) => value !== undefined));
      if (fontSizeStep !== undefined) {
        if (!fontSizeStep || typeof fontSizeStep !== 'object' || Array.isArray(fontSizeStep) || toggle !== undefined ||
          !['increase', 'decrease'].includes(fontSizeStep.direction) || !Array.isArray(fontSizeStep.sizes) ||
          fontSizeStep.sizes.length > 256 || fontSizeStep.sizes.some((size, index, sizes) =>
            typeof size !== 'number' || !Number.isFinite(size) || size < MIN_FONT_SIZE_PT || size > MAX_FONT_SIZE_PT ||
            index > 0 && size <= sizes[index - 1]!)) return 'invalid-payload';
        const size = formattingAt(at).formatting.fontSizePt ?? DEFAULT_EDITOR_TEXT_STYLE.fontSizePt;
        const sizes = fontSizeStep.sizes.map(normalizeFontSize);
        resolved.fontSizePt = fontSizeStep.direction === 'increase'
          ? sizes.find((value) => value > size) ?? Math.max(size, Math.min(MAX_FONT_SIZE_PT, size + 1))
          : sizes.reverse().find((value) => value < size) ?? Math.min(size, Math.max(MIN_FONT_SIZE_PT, size - 1));
      }
      if (toggle) {
        const current = formattingAt(at).formatting;
        if (toggle === 'underline') resolved.underline = current.underline !== null && current.underline !== 'none' ? 'none' : 'sng';
        else resolved[toggle] = current[toggle] !== true;
      }
      return { kind: 'format', selection: { ...at }, patch: resolved };
    }
    if (at.kind === 'shape') return 'unsupported';
    if (command.kind === 'compositionStart' || command.kind === 'compositionCommit') {
      if (typeof command.payload.compositionId !== 'string' || command.payload.compositionId.length === 0) {
        return 'invalid-payload';
      }
      if (command.kind === 'compositionStart') {
        return composition ? 'invalid-composition'
          : { kind: 'compositionStart', id: command.payload.compositionId, selection: { ...at } };
      }
      if (composition?.id !== command.payload.compositionId || composition.committed) return 'invalid-composition';
      const captured = composition.selection;
      if (at.storyId !== captured.storyId || at.slideId !== captured.slideId || at.shapeId !== captured.shapeId) {
        return 'bad-target';
      }
    }
    let start = Math.min(at.anchor, at.focus);
    let end = Math.max(at.anchor, at.focus);
    let text = '';
    let style: TextStyle | undefined;
    switch (command.kind) {
      case 'insertText':
      case 'replaceText':
      case 'compositionCommit':
        text = command.payload.text;
        style = command.payload.style;
        if (!validStyle(style)) return 'invalid-payload';
        break;
      case 'paragraphBreak':
        text = '\n';
        break;
      case 'deleteBackward':
      case 'deleteForward': {
        const count = command.payload.count ?? 1;
        if (command.payload.unit !== 'codePoint' || !Number.isSafeInteger(count) || count < 1) {
          return 'invalid-payload';
        }
        if (start === end) {
          for (let index = 0; index < count; index += 1) {
            if (command.kind === 'deleteBackward') {
              if (start === 0) break;
              start -= 1;
              if (!boundary(story.text, start)) start -= 1;
            } else {
              if (end === story.text.length) break;
              end += 1;
              if (!boundary(story.text, end)) end += 1;
            }
          }
        }
        break;
      }
      default:
        return 'invalid-payload';
    }
    if (!validText(text) || story.text.length - (end - start) + text.length > 0xfffffffe) {
      return 'invalid-payload';
    }
    if (command.kind === 'compositionCommit' && text === '') {
      return { kind: 'text', selection: { ...at }, start, end, text };
    }
    const cached = stories.get(storyKey(at))!;
    const joinsParagraphs = story.paragraphs.some((paragraph) => start <= paragraph.end && paragraph.end < end);
    for (const paragraph of story.paragraphs) {
      if (start <= paragraph.end && end >= paragraph.start &&
        (!paragraph.editable || joinsParagraphs && paragraph.fields.length > 0)) return 'unsupported';
      if (paragraph.fields.some((field) =>
        (start > field.start && start < field.end) || (end > field.start && end < field.end))) return 'unsupported';
      if (text.includes('\n') && paragraph.start <= start && start <= paragraph.end &&
        paragraph.fields.some((field) => field.end <= start)) return 'unsupported';
      const source = cached.source.get(paragraph.paragraphId);
      if (start > paragraph.end || end < paragraph.start || !source || source.fields.length === 0) continue;
      if (!source.editable) return 'unsupported';
      const live = story.text.slice(paragraph.start, paragraph.end);
      const split = text.lastIndexOf('\n');
      const projected = (split < 0 ? live.slice(0, start - paragraph.start) + text : text.slice(split + 1)) +
        live.slice(end - paragraph.start);
      const projectedStart = split < 0 ? paragraph.start : start + split + 1;
      const [prefix, suffix] = commonEnds(source.text, projected);
      for (const [index, field] of source.fields.entries()) {
        const current = paragraph.fields[index];
        if (!current) return 'unsupported';
        if (start <= current.start && current.end <= end) continue;
        const shift = field.end <= prefix ? 0
          : field.start >= source.text.length - suffix ? projected.length - source.text.length : undefined;
        const moved = (current.start >= end ? text.length - (end - start) : 0) - projectedStart;
        if (shift === undefined || field.start + shift !== current.start + moved ||
          field.end + shift !== current.end + moved || source.text.slice(field.start, field.end).includes('\n')) {
          return 'unsupported';
        }
      }
    }
    if (style === undefined && (command.kind === 'insertText' || command.kind === 'replaceText' ||
      command.kind === 'compositionCommit')) {
      style = { ...DEFAULT_EDITOR_TEXT_STYLE,
        ...Object.fromEntries(Object.entries(formattingAt(at).formatting)
          .filter(([key, value]) => key !== 'caps' && value !== null && value !== undefined)),
        ...marksAt(at) };
    }
    return { kind: 'text', selection: { ...at }, start, end, text, style };
  }

  function mapConfirmedPoint(at: PresentationCommandSelection, basisSeq: number): PresentationCommandSelection {
    let position = at.anchor;
    for (const mutation of storyMutations.get(storyKey(at))?.values() ?? []) {
      if (mutation.seq <= basisSeq || position < mutation.start) continue;
      position = position >= mutation.start + mutation.deleted
        ? position - mutation.deleted + mutation.inserted : mutation.start + mutation.inserted;
    }
    return { ...at, anchor: position, focus: position };
  }

  function prune(seq: number, basisVersion: string, dependencySeq?: number): void {
    retainedBasis = { seq, version: basisVersion };
    for (const [key, position] of versions) {
      if (position > seq) break;
      versions.delete(key);
    }
    for (const position of outcomes.keys()) {
      if (position > seq) break;
      if (position === dependencySeq) continue;
      outcomes.delete(position);
    }
    for (const position of selections.keys()) {
      if (position > seq) break;
      if (position === dependencySeq) continue;
      selections.delete(position);
    }
    const mutationFloor = Math.min(seq, ...[...frames.values()].map((entry) => entry.tag.coversThroughSeq));
    for (const [position, mutation] of mutations) {
      if (position > mutationFloor) break;
      mutations.delete(position);
      const story = storyMutations.get(mutation.story)!;
      story.delete(position);
      if (story.size === 0) storyMutations.delete(mutation.story);
    }
  }

  function trim(): void {
    while (versions.size > 256) {
      const [key, seq] = versions.entries().next().value!;
      prune(seq, key);
    }
    while (mutations.size > 512) {
      const mutation = mutations.values().next().value!;
      for (const [key, retained] of frames) {
        if (retained.tag.coversThroughSeq < mutation.seq) frames.delete(key);
      }
      prune(mutation.seq, mutation.version);
    }
    while (outcomes.size > 256) outcomes.delete(outcomes.keys().next().value!);
    while (selections.size > 256) selections.delete(selections.keys().next().value!);
  }

  function apply(command: PresentationCommandEnvelope, admitted: Admission): boolean {
    const opened = presentation();
    if (admitted.kind === 'noop') return false;
    if (command.kind !== 'selectAtPoint') pointerRange = undefined;
    if (admitted.kind === 'history') {
      if (!(admitted.direction === 'undo' ? opened.canUndo() : opened.canRedo())) return false;
      opened.addUndoBoundary();
      const result = opened.historyStep(admitted.direction);
      if (!result.applied) return false;
      admitted.receipt = result;
      typing = undefined;
      composition = undefined;
      selection = null;
      objectSelection = null;
      clearMarks();
      return true;
    }
    if (admitted.kind !== 'text') closeTyping();
    if (admitted.kind === 'activation') {
      clearMarks();
      activeSlideId = admitted.slideId;
      selection = null;
      objectSelection = null;
      composition = undefined;
      return false;
    }
    if (admitted.kind === 'mode') {
      readOnly = admitted.readOnly;
      return false;
    }
    if (admitted.kind === 'observer') { observeChanges = admitted.enabled; return false; }
    if (admitted.kind === 'invoke') {
      const selected = selection ? anchor(selection) : null;
      const object = objectSelection;
      const before = opened.version();
      opened.addUndoBoundary();
      admitted.result = admitted.run();
      opened.addUndoBoundary();
      if (before === opened.version()) return false;
      admitted.receipt = { applied: true, changedSlideIds: [...revisions.keys()], deckSizeChanged: false,
        slideOrderChanged: ['insertSlide', 'deleteSlide', 'moveSlide'].includes(admitted.method) };
      refreshHistory(command.seq, admitted.receipt);
      if (admitted.method === 'insertSlide') activeSlideId = (admitted.result as { slideId: string }).slideId;
      selection = selected ? resolve(selected) : null;
      if (selection && !target(selection)) selection = null;
      objectSelection = object && slideIndices.has(object.slideId) ? selectedObject(object.slideId, object.shapeId) : null;
      clearMarks();
      return true;
    }
    if (admitted.kind === 'remote') {
      const selected = selection ? anchor(selection) : null;
      const composed = composition ? anchor(composition.selection) : null;
      const retained = new Map(selections);
      const before = opened.version();
      opened.applyUpdate(new Uint8Array(admitted.bytes));
      if (before === opened.version()) return false;
      admitted.receipt = { applied: true, changedSlideIds: [...revisions.keys()], deckSizeChanged: true, slideOrderChanged: true };
      refreshHistory(command.seq, admitted.receipt);
      for (const [seq, value] of retained) selections.set(seq, value);
      selection = selected ? resolve(selected) : null;
      if (selection && !target(selection)) selection = null;
      if (composition && composed) {
        const next = resolve(composed);
        if (next && target(next)) composition = { ...composition, selection: next }; else composition = undefined;
      }
      objectSelection = null;
      clearMarks();
      return true;
    }
    if (admitted.kind === 'batch') {
      const anchored = selection ? anchor(selection) : null;
      admitted.result = opened.applyEdits(admitted.request);
      if (!admitted.result.ok || !admitted.result.applied) return false;
      admitted.receipt = { applied: true, changedSlideIds: admitted.result.changedSlides, deckSizeChanged: false, slideOrderChanged: false };
      refreshHistory(command.seq, admitted.receipt);
      selection = anchored ? resolve(anchored) : null;
      objectSelection = null;
      clearMarks();
      return true;
    }
    if (admitted.kind === 'font') {
      const font = { ...admitted.font, bytes: new Uint8Array(admitted.font.bytes) };
      admitted.fontId = admitted.fallback ? opened.registerFallbackFont(font) : opened.registerFont(font);
      layouts.clear();
      for (const story of stories.values()) { story.formatting = undefined; story.toolbar = undefined; }
      return false;
    }
    if (admitted.kind === 'insertion') {
      opened.addUndoBoundary();
      admitted.receipt = command.kind === 'insertTextBox' ? opened.addTextBox(admitted.slideId, admitted.draft as ShapeDraft)
        : command.kind === 'insertShape' ? opened.addShape(admitted.slideId, admitted.draft as PresetShapeDraft)
        : opened.addPicture(admitted.slideId, admitted.draft as PictureDraft);
      opened.addUndoBoundary();
      refreshMetadata();
      const content = opened.readCommandContent();
      if (!content.ok) throw new Error(content.failure.message);
      cacheStories(content);
      layouts.delete(admitted.slideId);
      activeSlideId = admitted.slideId;
      const object = selectedObject(admitted.slideId, admitted.receipt.shapeId);
      objectSelection = command.kind === 'insertTextBox' ? null : object;
      selection = command.kind === 'insertTextBox' && object?.storyId ? { slideId: object.slideId, shapeId: object.shapeId,
        storyId: object.storyId, anchor: 0, focus: 0, affinity: 'downstream' } : object ? objectStory(object) : null;
      if (command.kind === 'insertTextBox' && selection) {
        const style = (admitted.draft as ShapeDraft).style;
        draftStyles.set(storyKey(selection), { ...DEFAULT_EDITOR_TEXT_STYLE,
          ...Object.fromEntries(Object.entries(style).filter(([, value]) => value !== undefined && value !== null)),
          fontSizePt: normalizeFontSize(style.fontSizePt ?? DEFAULT_EDITOR_TEXT_STYLE.fontSizePt) });
      }
      clearMarks();
      return true;
    }
    if (admitted.kind === 'compositionEnd') {
      composition = undefined;
      return false;
    }
    if (admitted.kind === 'selection') {
      objectSelection = null;
      if (!admitted.selection || marksAt(admitted.selection) !== storedMarks) clearMarks();
      selection = admitted.selection;
      if (command.kind === 'selectAtPoint' && !command.payload.drag && selection) {
        pointerRange = { group: command.payload.undoGroup, selection: anchor(selection) };
      }
      if (selection) activeSlideId = selection.slideId;
      return false;
    }
    if (admitted.kind === 'object' || admitted.kind === 'transform') {
      const object = admitted.object;
      clearMarks();
      selection = objectStory(object);
      objectSelection = object;
      activeSlideId = object.slideId;
      if (admitted.kind === 'object') return false;
      if (object.rect && Object.keys(admitted.rect).every((key) =>
        admitted.rect[key as keyof ShapeRect] === object.rect![key as keyof ShapeRect])) return false;
      opened.addUndoBoundary();
      if (admitted.move) opened.moveShape(object.slideId, object.shapeId, admitted.rect.x, admitted.rect.y);
      else opened.setShapeRect(object.slideId, object.shapeId, admitted.rect);
      opened.addUndoBoundary();
      return true;
    }
    if (admitted.kind === 'compositionStart') {
      objectSelection = null;
      if (marksAt(admitted.selection) !== storedMarks) clearMarks();
      selection = admitted.selection;
      activeSlideId = selection.slideId;
      composition = { id: admitted.id, selection: { ...selection }, committed: false };
      return false;
    }
    if (admitted.kind === 'format' || admitted.kind === 'alignment') {
      const at = admitted.selection;
      if (at.kind !== 'shape' || !objectSelection || storyKey(at) !== storyKey(objectSelection)) objectSelection = null;
      if (marksAt(at) !== storedMarks) clearMarks();
      selection = at;
      activeSlideId = at.slideId;
      const start = Math.min(at.anchor, at.focus);
      const end = at.kind === 'shape' ? stories.get(storyKey(at))!.snapshot.length : Math.max(at.anchor, at.focus);
      if (admitted.kind === 'format' && start === end && at.kind !== 'shape') {
        Object.assign(storedMarks, admitted.patch);
        return false;
      }
      opened.addUndoBoundary();
      if (admitted.kind === 'format') opened.formatText(at.storyId, start, end, admitted.patch);
      else opened.setParagraphAlignment(at.storyId, start, end, admitted.alignment);
      opened.addUndoBoundary();
      return true;
    }
    const { start, end, text, style } = admitted;
    const at = admitted.selection;
    objectSelection = null;
    if (marksAt(at) !== storedMarks) clearMarks();
    if (command.kind === 'compositionCommit' && text === '') {
      composition!.committed = true;
      return false;
    }
    if (start !== end || text.length > 0) {
      const joinsParagraphs = stories.get(storyKey(at))!.text.paragraphs
        .some((paragraph) => start <= paragraph.end && paragraph.end < end);
      const grouped = !joinsParagraphs &&
        (command.kind === 'insertText' || command.kind === 'deleteBackward' || command.kind === 'deleteForward');
      const time = command.payload.inputTime!;
      const group = command.payload.undoGroup!;
      if (!grouped || !typing || typing.target !== storyKey(at) || typing.group !== group ||
        typing.kind !== command.kind || time < typing.time || time - typing.time >= 500 ||
        at.anchor !== typing.position || at.focus !== typing.position) opened.addUndoBoundary();
      if (end > start) {
        if (joinsParagraphs) presentationDeleteTextJoiningParagraphs(opened, at.storyId, start, end);
        else opened.deleteText(at.storyId, start, end);
      }
      let offset = start;
      const paragraphs = text.split('\n');
      for (const [index, paragraph] of paragraphs.entries()) {
        if (index > 0) {
          opened.insertParagraphBreak(at.storyId, offset);
          offset += 1;
        }
        if (paragraph.length > 0) {
          opened.insertText(at.storyId, offset, paragraph, style);
          offset += paragraph.length;
        }
      }
      if (grouped) typing = { target: storyKey(at), group, time, kind: command.kind, position: start + text.length };
      else { opened.addUndoBoundary(); typing = undefined; }
    }
    selection = { ...at, anchor: start + text.length, focus: start + text.length, affinity: 'downstream' };
    delete selection.goalX;
    activeSlideId = at.slideId;
    if (command.kind === 'compositionCommit') composition!.committed = true;
    return start !== end || text.length > 0;
  }

  const handlers = {
    async openCommandEditor(_, bytes, input) {
      if (disposed) throw new Error('Presentation command session is disposed');
      if (handle || opening) throw new Error('Presentation command session is already open');
      if (!input || typeof input.session !== 'string' || !input.session ||
        typeof input.incarnation !== 'string' || !input.incarnation ||
        (input.readOnly !== undefined && typeof input.readOnly !== 'boolean')) {
        throw new TypeError('Invalid presentation command open options');
      }
      opening = true;
      try {
        await (options.initWasm ?? initWasm)(input.wasm);
        if (disposed) throw new Error('Presentation command session is disposed');
        handle = openPresentation(new Uint8Array(bytes), { clientId: input.clientId,
          fonts: input.fonts?.map((face) => ({ ...face, bytes: new Uint8Array(face.bytes) })),
          fallbackFonts: input.fallbackFonts?.map((face) => ({ ...face, bytes: new Uint8Array(face.bytes) })),
          initialUpdate: input.initialUpdate });
        offUpdates = handle.onUpdate((update, origin) => {
          pendingUpdates.push({ bytes: update.slice().buffer, origin });
        });
        const openedMetadata = refreshMetadata();
        const index = input.activeSlideIndex ?? 0;
        if (!Number.isSafeInteger(index)) {
          throw new RangeError('Active slide index is out of range');
        }
        identity = { session: input.session, incarnation: input.incarnation };
        readOnly = input.readOnly ?? false;
        observeChanges = input.observeChanges ?? false;
        initialUpdate = input.initialUpdate !== undefined;
        activeSlideId = openedMetadata.slides[Math.max(0, Math.min(index, openedMetadata.slides.length - 1))]?.id ?? null;
        handle.setUndoCaptureMode('manual');
        version = handle.version();
        projection.initialize(version);
        versions.set(version, 0);
        const content = handle.readCommandContent();
        if (!content.ok) throw new SessionFailure('message',
          `Presentation command open refused: ${content.failure.message}`, content.failure.code);
        cacheStories(content);
        for (const slide of openedMetadata.slides) revisions.set(slide.id, 0);
        const result = ack();
        const recoveryUpdate = handle.encodeStateAsUpdate().slice().buffer;
        return transferable({ ...result.value, shared: projection.ack(result.value),
          slides: openedMetadata.slides, size: openedMetadata.size,
          activeSlideId, frame: result.value.frame ?? null, recoveryUpdate }, [...result.transfer, recoveryUpdate]);
      } catch (error) {
        dispose();
        throw error;
      } finally { opening = false; }
    },
    command(_, command) {
      presentation();
      if (command.kind !== 'activateSlide') refreshTopology();
      if (!matches(command) || !Number.isSafeInteger(command.seq) || command.seq < 1) return ack();
      const cached = outcomes.get(command.seq);
      if (cached) return ack([{ ...cached }]);
      if (command.seq <= throughSeq) return ack([{ seq: command.seq, status: 'refused',
        reason: 'stale-basis', committed: false }]);
      if (command.seq !== throughSeq + 1) return ack();
      const admitted = admit(command);
      if (typeof admitted === 'string') {
        if (command.kind === 'compositionCommit' && composition &&
          composition.id === command.payload?.compositionId) composition = undefined;
        const outcome: PresentationCommandOutcome = { seq: command.seq,
          status: admitted === 'dependency-refused' ? 'cancelled' : 'refused',
          reason: admitted, committed: false };
        throughSeq = command.seq;
        outcomes.set(command.seq, outcome);
        trim();
        return ack([outcome]);
      }
      const previous = { version, selection, objectSelection, activeSlideId, editState: lastEditState };
      const changedSlideId = 'selection' in admitted ? admitted.selection?.slideId
        : 'object' in admitted ? admitted.object.slideId : undefined;
      const previousRevision = changedSlideId === undefined ? undefined : revisions.get(changedSlideId);
      try {
        const beforeText = admitted.kind === 'text' ? target(admitted.selection)!.text : '';
        const committed = apply(command, admitted);
        const historyDirtySlideIds = committed && admitted.kind === 'history'
          ? refreshHistory(command.seq, admitted.receipt!, true) : admitted.kind === 'remote' || admitted.kind === 'invoke' ? admitted.receipt?.changedSlideIds : undefined;
        if (committed && admitted.kind === 'text') refreshStory(admitted.selection, admitted.start, admitted.end);
        if (committed && (admitted.kind === 'format' || admitted.kind === 'alignment')) {
          const cached = stories.get(storyKey(admitted.selection))!;
          cached.snapshot = presentation().story(admitted.selection.storyId);
          cached.formatting = undefined;
          cached.echoRanges = [];
        }
        throughSeq = command.seq;
        version = presentation().version();
        metadataVersion = version;
        if (committed) projection.commit(version);
        if (!versions.has(version) && version !== retainedBasis?.version) versions.set(version, throughSeq);
        const outcome: PresentationCommandOutcome = { seq: command.seq,
          status: command.kind === 'compositionCommit' && command.payload.text === '' ? 'cancelled'
            : committed || admitted.kind === 'invoke' && admitted.method === 'propose' ||
            admitted.kind === 'format' || admitted.kind === 'alignment' ||
            admitted.kind === 'font' || admitted.kind === 'activation' || admitted.kind === 'selection' || admitted.kind === 'object' || command.kind === 'setReadOnly' ||
            command.kind === 'compositionStart' || command.kind === 'compositionEnd' ? 'applied' : 'noop',
          committed };
        outcomes.set(command.seq, outcome);
        if (committed && admitted.kind === 'text') {
          const story = storyKey(admitted.selection);
          const mutation: Mutation = { seq: command.seq, story, start: admitted.start,
            deleted: admitted.end - admitted.start, inserted: admitted.text.length, version, beforeText };
          mutations.set(command.seq, mutation);
          const records = storyMutations.get(story) ?? new Map<number, Mutation>();
          records.set(command.seq, mutation);
          storyMutations.set(story, records);
        }
        selections.set(command.seq, selection ? anchor(selection) : null);
        trim();
        const dirtySlideIds = admitted.kind === 'font' ? [...revisions.keys()] : admitted.kind === 'batch' && admitted.result?.ok
          ? admitted.result.changedSlides : historyDirtySlideIds ?? (committed && activeSlideId !== null ? [activeSlideId] : []);
        for (const slideId of dirtySlideIds) {
          const revision = revisions.get(slideId);
          if (revision !== undefined) revisions.set(slideId, revision + 1);
        }
        const result = ack([outcome], dirtySlideIds, admitted.kind === 'history' || admitted.kind === 'remote' || admitted.kind === 'invoke', admitted.kind === 'insertion' ? admitted.receipt : undefined,
          admitted.kind === 'batch' ? admitted.result : undefined, admitted.kind === 'font' ? admitted.fontId : undefined);
        if (admitted.kind === 'invoke') result.value.mutationResult = admitted.result;
        return result;
      } catch (error) {
        if (version !== previous.version) versions.delete(version);
        version = previous.version;
        selection = previous.selection;
        objectSelection = previous.objectSelection;
        activeSlideId = previous.activeSlideId;
        lastEditState = previous.editState;
        if (changedSlideId !== undefined && previousRevision !== undefined) revisions.set(changedSlideId, previousRevision);
        failedSeq = command.seq;
        throughSeq = command.seq;
        readOnly = true;
        outcomes.set(command.seq, { seq: command.seq, status: 'refused',
          reason: 'unknown-outcome', committed: false });
        throw new CommandOutcomeUnknown(command.seq,
          `Command ${command.seq} has an unknown outcome: ${error instanceof Error ? error.message : String(error)}`);
      }
    },
    edit(_, command): TransferResult<PresentationSharedAck> {
      presentation();
      if (!isCommandEnvelope(command as unknown)) {
        if (!Number.isSafeInteger(command?.seq) || command.seq < 1) {
          throw new TypeError('Invalid presentation edit envelope');
        }
        return sharedAck(ack([{ seq: command.seq, status: 'refused',
          reason: 'invalid-payload', committed: false }]));
      }
      if (command.session !== identity!.session || command.incarnation !== 0) {
        const result = sharedAck(ack());
        result.value.outcomes = [{ seq: command.seq, status: 'refused', committed: false,
          reason: command.session !== identity!.session ? 'wrong-session' : 'stale-incarnation' }];
        return result;
      }
      const envelope = projection.envelope(command, identity!);
      if (command.seq === throughSeq + 1) {
        if (command.basis.epoch !== documentEpoch) sharedReasons.set(command.seq, 'stale-epoch');
        else if (command.expectVersion === true && command.basis.version !== projection.version) {
          sharedReasons.set(command.seq, 'stale-version');
          envelope.basis.version = '';
        }
      }
      return sharedAck(handlers.command(null, envelope));
    },
    checkpoint() {
      throw new Error('Presentation command checkpoints are unavailable');
    },
    thumbnailFrame(_, request) {
      presentation();
      if (!matches(request)) throw new Error('Presentation command thumbnail identity differs');
      refreshTopology();
      const slideRevision = revisions.get(request.slideId);
      const base = { ...identity!, epoch: documentEpoch, slideId: request.slideId, slideRevision: slideRevision ?? 0 };
      if (slideRevision === undefined || !slideIndices.has(request.slideId)) {
        return { ...base, status: 'refused' as const, reason: 'bad-target' as const };
      }
      if (request.knownRevision === slideRevision) return { ...base, status: 'unchanged' as const };
      const result = frame(request.slideId, false);
      for (const { assetId } of result.value.media) sentMedia.add(assetId);
      return transferable({ ...base, status: 'frame' as const, frame: result.value }, result.transfer);
    },
    flushCommandEditor(_, request) {
      fence(request);
      return ack();
    },
    saveCommandEditor(_, request) {
      fence(request);
      const bytes = new Uint8Array(presentation().save()).buffer;
      const updates = pendingUpdates.splice(0);
    return transferable({ ...identity!, epoch: documentEpoch, throughSeq, version,
      ...(updates.length ? { updates } : {}), bytes }, [bytes]);
    },
    accessCommandEditor(_, request, method, args) {
      fence(request);
      if (!Array.isArray(args)) throw new TypeError('Invalid access arguments');
      const opened = presentation();
      switch (method) {
        case 'encodeStateVector': return opened.encodeStateVector();
        case 'encodeStateAsUpdate': return opened.encodeStateAsUpdate(args[0] as Uint8Array | undefined);
        case 'snapshot': return opened.snapshot();
        case 'comments': return opened.comments();
        case 'story': return opened.story(args[0] as string);
        case 'findText': return opened.findText(args[0] as Parameters<PresentationHandle['findText']>[0]);
        case 'validateEdits': return readOnly ? { ok: false as const, version,
          failure: { code: 'read-only' as const, message: 'The editor is read-only' } } : opened.validateEdits(args[0] as PptxEditRequest);
        case 'exportStructured': return opened.exportStructured(args[0] as Parameters<PresentationHandle['exportStructured']>[0]);
        case 'exportMarkdown': return opened.exportMarkdown(args[0] as Parameters<PresentationHandle['exportMarkdown']>[0]);
        case 'searchText': return opened.searchText(args[0] as string, args[1] as Parameters<PresentationHandle['searchText']>[1]);
        case 'listProposals': return opened.listProposals();
        case 'previewProposal': return opened.previewProposal(args[0] as string);
        case 'layoutSlide': return opened.layoutSlide(args[0] as number);
        case 'layoutProposalSlide': return opened.layoutProposalSlide(args[0] as string, args[1] as number);
        case 'layoutProposalDiffSlide': return opened.layoutProposalDiffSlide(args[0] as string, args[1] as number);
        default: throw new TypeError('Unsupported access method');
      }
    },
    checkpointCommandEditor(_, request) {
      fence(request);
      const bytes = presentation().save().slice().buffer;
      const recoveryUpdate = presentation().encodeStateAsUpdate().slice().buffer;
      return transferable({ ...identity!, epoch: documentEpoch, throughSeq, version,
        bytes, recoveryUpdate }, [bytes, recoveryUpdate]);
    },
    prepareCommandPicture(_, request, bytes) {
      fence(request);
      if (!(bytes instanceof ArrayBuffer) || bytes.byteLength === 0 || bytes.byteLength > 8 * 1024 * 1024) throw new Error('Invalid picture bytes');
      const result = decodeTiffImage(new Uint8Array(bytes)).slice().buffer;
      return transferable(result, [result]);
    },
    readEditor(_, request, read) {
      fence(request);
      const result = presentation().readContent(read);
      const reply = ack();
      return transferable({ ...reply.value, result }, reply.transfer);
    },
  } satisfies MethodHandlers<PresentationCommandMethods & PresentationSharedMethods, null>;

  function dispose(): void {
    if (disposed) return;
    disposed = true;
    const opened = handle;
    offUpdates?.(); offUpdates = undefined; pendingUpdates.length = 0;
    handle = undefined;
    metadata = undefined;
    slideIndices.clear();
    composition = undefined;
    typing = undefined;
    lastEditState = undefined;
    versions.clear();
    projection.dispose();
    outcomes.clear();
    sharedReasons.clear();
    retainedBasis = undefined;
    mutations.clear();
    storyMutations.clear();
    stories.clear();
    draftStyles.clear();
    slideStoryKeys.clear();
    objectTokens.clear();
    selections.clear();
    revisions.clear();
    sentMedia.clear();
    frames.clear();
    layouts.clear();
    opened?.dispose();
  }

  return { handlers, policies: {
    openCommandEditor: { lane: 'interactive' },
    command: { lane: 'input', mutates: true, userInput: true },
    edit: { lane: 'input', mutates: true, userInput: true },
    checkpoint: { lane: 'interactive' },
    thumbnailFrame: { lane: 'interactive' },
    flushCommandEditor: { lane: 'interactive' },
    saveCommandEditor: { lane: 'interactive' },
    readEditor: { lane: 'interactive' },
    prepareCommandPicture: { lane: 'interactive' },
    accessCommandEditor: { lane: 'interactive' },
    checkpointCommandEditor: { lane: 'interactive' },
  }, dispose };
}

function storyKey(value: PptxStoryTarget): string {
  return JSON.stringify([value.slideId, value.shapeId, value.storyId]);
}

function commonEnds(source: string, target: string): [number, number] {
  let prefix = 0;
  const length = Math.min(source.length, target.length);
  while (prefix < length && source[prefix] === target[prefix]) prefix += 1;
  if (!boundary(source, prefix) || !boundary(target, prefix)) prefix -= 1;
  let suffix = 0;
  while (suffix < length - prefix && source[source.length - suffix - 1] === target[target.length - suffix - 1]) {
    suffix += 1;
  }
  if (!boundary(source, source.length - suffix) || !boundary(target, target.length - suffix)) suffix -= 1;
  return [prefix, suffix];
}

function boundary(text: string, offset: number): boolean {
  return Number.isSafeInteger(offset) && offset >= 0 && offset <= text.length &&
    !(offset > 0 && offset < text.length && text.charCodeAt(offset - 1) >= 0xd800 &&
      text.charCodeAt(offset - 1) <= 0xdbff && text.charCodeAt(offset) >= 0xdc00 &&
      text.charCodeAt(offset) <= 0xdfff);
}

function validText(value: unknown): value is string {
  if (typeof value !== 'string') return false;
  for (const character of value) {
    const point = character.codePointAt(0)!;
    if (!(point === 9 || point === 10 || point === 13 || (point >= 0x20 && point <= 0xd7ff) ||
      (point >= 0xe000 && point <= 0xfffd) || point >= 0x10000)) return false;
  }
  return true;
}

function validStyle(value: unknown): value is TextStyle | undefined {
  if (value === undefined) return true;
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const underlines = ['none', 'words', 'sng', 'dbl', 'heavy', 'dotted', 'dottedHeavy',
    'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy',
    'dotDotDash', 'dotDotDashHeavy', 'wavy', 'wavyHeavy', 'wavyDbl'];
  return Object.entries(value as Record<string, unknown>).every(([key, entry]) => {
    if (!['bold', 'italic', 'fontFamily', 'underline', 'color', 'fontSizePt', 'spacingPt', 'baselinePct'].includes(key)) return false;
    if (entry === undefined) return true;
    switch (key) {
      case 'bold':
      case 'italic': return typeof entry === 'boolean';
      case 'fontFamily': return validText(entry);
      case 'underline': return typeof entry === 'string' && underlines.includes(entry);
      case 'color': return typeof entry === 'string' && /^#?[0-9a-f]{6}$/i.test(entry);
      case 'fontSizePt': return typeof entry === 'number' && Number.isFinite(entry) && entry >= MIN_FONT_SIZE_PT && entry <= MAX_FONT_SIZE_PT;
      case 'spacingPt': return typeof entry === 'number' && Number.isFinite(entry) && Math.abs(entry) <= 4000;
      case 'baselinePct': return typeof entry === 'number' && Number.isFinite(entry) &&
        entry * 1000 >= -2147483648 && entry * 1000 <= 2147483647;
      default: return false;
    }
  });
}

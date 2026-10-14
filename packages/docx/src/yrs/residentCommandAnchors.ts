import type { YrsLoc, YrsStickyPosition, YrsStoryRange } from './index';
import type {
  DocxCommandApplyContext,
  DocxCommandEnvelope,
  DocxCommandPosition,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
  ResolvedDocxCommandSelection,
} from './residentCommandProtocol';

const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u;
const commandStates = new WeakMap<ResidentDocxCommandEngine, ResidentDocxCommandState>();

export function registerCommandAnchorState(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState): void {
  commandStates.set(engine, state);
}

export function readCommandStickyPosition(engine: ResidentDocxCommandEngine, at: YrsLoc): YrsStickyPosition {
  const sticky = engine.encodeStickyPosition(at);
  commandStates.get(engine)?.stickyPositions.set(stickyPositionKey(sticky.story, sticky.encoded), {
    story: sticky.story, encoded: sticky.encoded.slice(),
  });
  return sticky;
}

export function stickyPositionKey(story: string, bytes: Uint8Array): string {
  return `${story}\0${Array.from(bytes).join(',')}`;
}

function position(value: unknown): value is DocxCommandPosition {
  if (!value || typeof value !== 'object' || Array.isArray(value) ||
    Object.keys(value).some((key) => !['story', 'paraId', 'offset', 'sticky'].includes(key))) return false;
  const at = value as DocxCommandPosition;
  return typeof at.story === 'string' && at.story.length > 0 &&
    typeof at.paraId === 'string' && at.paraId.length > 0 &&
    !LONE_SURROGATE.test(at.story) && !LONE_SURROGATE.test(at.paraId) &&
    Number.isSafeInteger(at.offset) && at.offset >= 0 && at.offset <= 0xffffffff &&
    (at.sticky === undefined || (at.sticky instanceof Uint8Array && at.sticky.length > 0));
}

export function commandSelectionRange(
  engine: ResidentDocxCommandEngine,
  selected: DocxCommandSelection
): YrsStoryRange | null {
  if (!selected || typeof selected !== 'object' || Array.isArray(selected) ||
    Object.keys(selected).some((key) => !['anchor', 'focus'].includes(key)) ||
    !position(selected.anchor) || !position(selected.focus) ||
    selected.anchor.story !== selected.focus.story) return null;
  const { anchor, focus } = selected;
  if (!engine.geometryReader.hasStory(anchor.story)) return null;
  if ([anchor, focus].some((at) => engine.geometryReader.paragraphIdCount(at.story, at.paraId) !== 1)) return null;
  if (anchor.paraId !== focus.paraId) {
    const paragraphs = engine.geometryReader.paragraphSpans(anchor.story);
    const indices = [anchor, focus].map((at) => paragraphs.findIndex(({ paraId }) => paraId === at.paraId));
    if (indices.some((index) => index < 0)) return null;
    for (const paragraph of paragraphs.slice(Math.min(...indices), Math.max(...indices) + 1)) {
      if (engine.geometryReader.paragraphIdCount(anchor.story, paragraph.paraId) !== 1) return null;
    }
  }
  const offsets = [anchor, focus].map((at) => {
    const span = engine.geometryReader.locateParagraph(at.story, at.paraId);
    return at.offset <= span.end - span.start ? span.start + at.offset : -1;
  });
  if (offsets.some((offset) => offset < 0)) return null;
  const [start, end] = offsets[0]! <= offsets[1]! ? [anchor, focus] : [focus, anchor];
  return {
    story: start.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
}

export function readCommandSelection(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  selected = state.selection,
  requireSticky = false
): ResolvedDocxCommandSelection {
  if (!selected || typeof selected !== 'object' || Array.isArray(selected) ||
    Object.keys(selected).some((key) => !['anchor', 'focus'].includes(key)) ||
    !position(selected.anchor) || !position(selected.focus) ||
    selected.anchor.story !== selected.focus.story) return { ok: false, reason: 'invalid-selection' };
  const resolve = (at: DocxCommandPosition): DocxCommandPosition | null => {
    if (!at.sticky) return requireSticky ? null : { ...at };
    const issued = state.stickyPositions.get(stickyPositionKey(at.story, at.sticky));
    if (!issued || issued.story !== at.story || !engine.geometryReader.hasStory(issued.story)) return null;
    const resolved = engine.resolveStickyPosition(issued);
    return resolved && resolved.story === at.story ? { ...resolved, sticky: at.sticky.slice() } : null;
  };
  const anchor = resolve(selected.anchor);
  const focus = resolve(selected.focus);
  if (!anchor || !focus) return { ok: false, reason: 'stale-anchor' };
  const selection = { anchor, focus };
  const range = commandSelectionRange(engine, selection);
  return range ? { ok: true, selection, range } : { ok: false, reason: 'invalid-selection' };
}

export function mintCommandSelection(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  selected: DocxCommandSelection
): DocxCommandSelection {
  const mint = (at: YrsLoc): DocxCommandPosition => {
    const sticky = engine.encodeStickyPosition(at);
    state.stickyPositions.set(stickyPositionKey(at.story, sticky.encoded), {
      story: at.story, encoded: sticky.encoded.slice(),
    });
    return { story: at.story, paraId: at.paraId, offset: at.offset, sticky: sticky.encoded.slice() };
  };
  return { anchor: mint(selected.anchor), focus: mint(selected.focus) };
}

export function resolveCommandSelection(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  command: DocxCommandEnvelope,
  context: DocxCommandApplyContext
): ResolvedDocxCommandSelection {
  if (!command.basis || !Number.isSafeInteger(command.basis.epoch) || command.basis.epoch < 0 ||
    !Number.isSafeInteger(command.basis.version) || command.basis.version < 0 ||
    command.basis.epoch !== context.basis.epoch || command.basis.version !== context.basis.version ||
    command.basis.version > context.currentVersion) return { ok: false, reason: 'stale-anchor' };
  const stale = command.basis.version !== context.currentVersion;
  if (stale && command.expectVersion) return { ok: false, reason: 'stale-version' };
  if (!command.payload || typeof command.payload !== 'object' || Array.isArray(command.payload)) {
    return { ok: false, reason: 'invalid-payload' };
  }
  const payload = command.payload as { selection?: DocxCommandSelection };
  const dependent = payload.selection === undefined && command.afterSeq !== undefined;
  if (stale && payload.selection === undefined && !dependent) return { ok: false, reason: 'stale-anchor' };
  if (dependent && (!Number.isSafeInteger(command.afterSeq) || command.afterSeq! <= 0 ||
    command.afterSeq! >= command.seq || !context.predecessorSelection)) {
    return { ok: false, reason: 'stale-anchor' };
  }
  const selected = payload.selection !== undefined ? payload.selection : dependent ? context.predecessorSelection : state.selection;
  return readCommandSelection(engine, state, selected, stale && !dependent);
}

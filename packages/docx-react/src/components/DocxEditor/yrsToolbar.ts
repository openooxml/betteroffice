import { paragraphSpacingAttrs } from '../../commands/paragraphSpacing';
import type { ColorValue, ParagraphAlignment } from '@betteroffice/docx/types/document';
import type {
  YrsInputPositionMap,
  YrsAuthor,
  YrsSelectionContext,
  YrsSession,
  YrsStoryRange,
  YrsTableRange,
} from '@betteroffice/docx/yrs';
import {
  applyListToggle,
  cellSelectionContainsCaret,
  cellSelectionRanges,
  compareYrsLocs,
  listOffAttrs,
  toggleAcross,
} from '@betteroffice/docx/yrs';
import type { YrsStoredFormatting, YrsStoredFormattingAction } from './YrsInput';
import type { TableContextInfo } from './types';
import { currentYrsTableContext, currentYrsTableTarget } from './yrsCommands';

/** A formatting operation over the current selection. */
export type FormattingAction =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'superscript'
  | 'subscript'
  | 'clearFormatting'
  | 'bulletList'
  | 'numberedList'
  | 'indent'
  | 'outdent'
  | 'insertLink'
  | 'setRtl'
  | 'setLtr'
  | { type: 'fontFamily'; value: string }
  | { type: 'fontSize'; value: number }
  | { type: 'textColor'; value: ColorValue | string }
  | { type: 'highlightColor'; value: string }
  | { type: 'alignment'; value: ParagraphAlignment }
  | { type: 'lineSpacing'; value: number }
  | { type: 'paragraphSpacing'; before?: number; after?: number; lineSpacing?: number }
  | { type: 'applyStyle'; value: string };

export interface YrsToolbarSelection {
  context: YrsSelectionContext;
  tableContext: TableContextInfo | null;
  range: YrsStoryRange;
  startParagraphIndex: number;
  endParagraphIndex: number;
}

/** Resolve the session's sticky selection into the paragraph-addressed toolbar range. */
export function currentYrsToolbarSelection(
  session: YrsSession,
  map: YrsInputPositionMap
): YrsToolbarSelection | null {
  const selection = session.selection();
  if (!selection || selection.anchor.story !== map.story || selection.head.story !== map.story) {
    return null;
  }

  const [start, end] =
    compareYrsLocs(map, selection.anchor, selection.head) <= 0
      ? [selection.anchor, selection.head]
      : [selection.head, selection.anchor];
  const startParagraphIndex = map.paragraphs.findIndex((entry) => entry.paraId === start.paraId);
  const endParagraphIndex = map.paragraphs.findIndex((entry) => entry.paraId === end.paraId);
  if (startParagraphIndex < 0 || endParagraphIndex < 0) return null;

  const range: YrsStoryRange = {
    story: map.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
  return {
    context: session.selectionContext(range),
    tableContext: currentYrsTableContext(session),
    range,
    startParagraphIndex,
    endParagraphIndex,
  };
}

function isCollapsed(range: YrsStoryRange): boolean {
  return range.start.paraId === range.end.paraId && range.start.offset === range.end.offset;
}

export function selectedYrsToolbarCells(session: YrsSession, range: YrsStoryRange): YrsTableRange | null {
  if (!isCollapsed(range)) return null;
  const cells = currentYrsTableTarget(session)?.range;
  return cells && cellSelectionContainsCaret(session, cells, range.story) ? cells : null;
}

function paragraphRange(story: string, paraId: string): YrsStoryRange {
  return {
    story,
    start: { paraId, offset: 0 },
    end: { paraId, offset: 0 },
  };
}

function paragraphNumber(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : 0;
}

function effectiveFontSize(context: YrsSelectionContext): number | null {
  if (context.fontSize != null) return context.fontSize;
  const defaults = context.paragraphProperties.defaultTextFormatting;
  if (!defaults || typeof defaults !== 'object') return null;
  const value = (defaults as { fontSize?: unknown }).fontSize;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (value && typeof value === 'object') {
    const size = (value as { size?: unknown }).size;
    if (typeof size === 'number' && Number.isFinite(size)) return size;
  }
  return null;
}

function hasExplicitFontSize(session: YrsSession, selection: YrsToolbarSelection): boolean {
  const { range, startParagraphIndex, endParagraphIndex } = selection;
  const indices = Array.from(
    { length: endParagraphIndex - startParagraphIndex + 1 },
    (_, index) => startParagraphIndex + index
  );
  return session.storySegmentUnits(range.story, indices).some((segments, index) => {
    const start = index === 0 ? range.start.offset : 0;
    const end = index === indices.length - 1 ? range.end.offset : Infinity;
    let offset = 0;
    return segments.some((segment) => {
      const at = offset;
      offset += segment.kind === 'text' ? segment.text.length : 1;
      return segment.kind === 'text' && segment.attributes.fontSize != null &&
        Math.min(offset, end) > Math.max(at, start);
    });
  });
}

/** Runs `edit` as one undo step although it writes several stories. */
function inOneUndoStep(session: YrsSession, edit: () => void): void {
  const mode = session.undoCaptureMode();
  session.setUndoCaptureMode('manual');
  try {
    edit();
  } finally {
    session.setUndoCaptureMode(mode);
  }
}

function paragraphNumPr(value: unknown): { numId?: number; ilvl?: number } | null {
  if (!value || typeof value !== 'object') return null;
  const candidate = value as { numId?: unknown; ilvl?: unknown };
  const numId = typeof candidate.numId === 'number' ? candidate.numId : undefined;
  const ilvl = typeof candidate.ilvl === 'number' ? candidate.ilvl : undefined;
  return numId == null && ilvl == null ? null : { numId, ilvl };
}

function forEachSelectedParagraph(
  session: YrsSession,
  selection: YrsToolbarSelection,
  apply: (range: YrsStoryRange, properties: Record<string, unknown>) => void
): void {
  const paragraphs = session
    .paragraphs(selection.range.story)
    .slice(selection.startParagraphIndex, selection.endParagraphIndex + 1);
  for (const paragraph of paragraphs) {
    apply(paragraphRange(selection.range.story, paragraph.paraId), paragraph.properties);
  }
}

function changeIndent(
  session: YrsSession,
  selection: YrsToolbarSelection,
  direction: 'indent' | 'outdent',
  suggesting?: YrsAuthor
): void {
  const startRange = paragraphRange(selection.range.story, selection.context.paraId);
  const numPr = paragraphNumPr(selection.context.paragraphProperties.numPr);
  if (numPr?.numId) {
    const level = numPr.ilvl ?? 0;
    if (direction === 'indent' && level < 8) {
      session.setParagraphAttrs(startRange, {
        numPr: { ...numPr, ilvl: level + 1 },
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      }, suggesting);
    } else if (direction === 'outdent' && level > 0) {
      session.setParagraphAttrs(startRange, {
        numPr: { ...numPr, ilvl: level - 1 },
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      }, suggesting);
    } else if (direction === 'outdent') {
      session.setParagraphAttrs(startRange, {
        ...listOffAttrs(
          selection.context.paragraphProperties,
          session.paragraphStyleNumbers?.(selection.range.story, selection.context.paraId)
        ),
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      }, suggesting);
    }
    return;
  }

  forEachSelectedParagraph(session, selection, (range, properties) => {
    const current = paragraphNumber(properties.indentLeft);
    const next = direction === 'indent' ? current + 720 : Math.max(0, current - 720);
    if (next !== current) session.setParagraphAttrs(range, { indentLeft: next }, suggesting);
  });
}

function textColorDelta(
  value: Extract<FormattingAction, { type: 'textColor' }>['value']
): { rgb: string } | { themeColor: string } | null {
  if (typeof value === 'string') return { rgb: value.replace(/^#/, '') };
  if ('auto' in value && value.auto) return null;
  if ('themeColor' in value && typeof value.themeColor === 'string') {
    return { themeColor: value.themeColor };
  }
  if ('rgb' in value && typeof value.rgb === 'string') {
    return { rgb: value.rgb.replace(/^#/, '') };
  }
  return null;
}

/** Translate a collapsed inline toolbar action into stored formatting. */
export function storedYrsToolbarFormatting(
  context: YrsSelectionContext,
  action: FormattingAction
): YrsStoredFormattingAction | null {
  if (action === 'bold' || action === 'italic' || action === 'underline') {
    return { type: 'toggle', mark: action, active: context[action] === true };
  }
  if (action === 'strikethrough') {
    return { type: 'toggle', mark: 'strike', active: context.strike === true };
  }
  if (action === 'superscript' || action === 'subscript') {
    return { type: 'toggle', mark: action, active: context[action] === true };
  }
  if (action === 'clearFormatting') return { type: 'clear' };
  if (typeof action !== 'object') return null;
  switch (action.type) {
    case 'fontFamily':
      return {
        type: 'set',
        delta: { fontFamily: { ascii: action.value, hAnsi: action.value } },
      };
    case 'fontSize':
      return { type: 'set', delta: { fontSize: action.value } };
    case 'textColor':
      return { type: 'set', delta: { color: textColorDelta(action.value) } };
    case 'highlightColor':
      return {
        type: 'set',
        delta: { highlight: action.value && action.value !== 'none' ? action.value : null },
      };
    default:
      return null;
  }
}

/** Overlay paragraph-local stored formatting onto the selection read model. */
export function withStoredYrsFormatting(
  selection: YrsToolbarSelection,
  stored: YrsStoredFormatting | null
): YrsToolbarSelection {
  if (!stored) return selection;
  const delta = stored.delta;
  const base = stored.clear
    ? {
        ...selection.context,
        bold: false as const,
        italic: false as const,
        underline: false as const,
        strike: false as const,
        superscript: false as const,
        subscript: false as const,
        fontFamily: null,
        fontSize: null,
        color: null,
        highlight: null,
      }
    : selection.context;
  const fontFamily = delta.fontFamily;
  const color = delta.color;
  const script = delta.other ?? {};
  return {
    ...selection,
    context: {
      ...base,
      bold: delta.bold === undefined ? base.bold : delta.bold === true,
      italic: delta.italic === undefined ? base.italic : delta.italic === true,
      underline:
        delta.underline === undefined
          ? base.underline
          : delta.underline !== false && delta.underline !== null,
      strike:
        delta.strike === undefined ? base.strike : delta.strike !== false && delta.strike !== null,
      superscript:
        script.superscript === undefined ? base.superscript : script.superscript === true,
      subscript: script.subscript === undefined ? base.subscript : script.subscript === true,
      fontFamily:
        fontFamily === undefined
          ? base.fontFamily
          : fontFamily?.ascii ?? fontFamily?.hAnsi ?? null,
      fontSize:
        delta.fontSize === undefined
          ? base.fontSize
          : delta.fontSize == null
            ? null
            : delta.fontSize * 2,
      color:
        color === undefined ? base.color : color?.rgb ?? color?.themeColor ?? null,
      highlight: delta.highlight === undefined ? base.highlight : delta.highlight,
    },
  };
}

/**
 * Apply one toolbar action to the live yrs selection.
 *
 * Returns false for commands outside the current yrs formatting surface. The
 * caller deliberately does not fall back while yrs-authoritative input
 * is active.
 */
export function applyYrsToolbarFormatting(
  session: YrsSession,
  map: YrsInputPositionMap,
  action: FormattingAction,
  suggesting?: YrsAuthor
): boolean {
  const selection = currentYrsToolbarSelection(session, map);
  if (!selection) return false;
  const { context, range } = selection;
  const cells = selectedYrsToolbarCells(session, range);
  const ranges = cells ? cellSelectionRanges(session, cells) : isCollapsed(range) ? [] : [range];
  const format = (edit: () => void): boolean => {
    if (ranges.length === 0) return false;
    if (cells) inOneUndoStep(session, edit);
    else edit();
    return true;
  };
  const each = (apply: (target: YrsStoryRange) => void) => format(() => ranges.forEach(apply));

  if (
    action === 'bold' ||
    action === 'italic' ||
    action === 'underline' ||
    action === 'superscript' ||
    action === 'subscript'
  ) {
    return cells ? format(() => toggleAcross(session, ranges, action))
      : each((target) => session.toggleMark(target, { type: action }));
  }
  if (action === 'strikethrough') {
    return cells ? format(() => toggleAcross(session, ranges, 'strike'))
      : each((target) => session.formatRange(target, { strike: context.strike !== true }));
  }
  if (action === 'clearFormatting') return each((target) => session.clearFormatting(target));
  if (action === 'bulletList' || action === 'numberedList') {
    applyListToggle(session, range, action === 'bulletList' ? 'bullet' : 'numbered', suggesting);
    return true;
  }
  if (action === 'indent' || action === 'outdent') {
    changeIndent(session, selection, action, suggesting);
    return true;
  }
  if (action === 'setRtl' || action === 'setLtr') {
    session.setParagraphAttrs(range, { bidi: action === 'setRtl' }, suggesting);
    return true;
  }

  if (typeof action !== 'object') return false;
  switch (action.type) {
    case 'alignment':
      session.setParagraphAttrs(range, { alignment: action.value }, suggesting);
      return true;
    case 'lineSpacing':
      session.setParagraphAttrs(
        range,
        { lineSpacing: action.value, lineSpacingRule: 'auto' },
        suggesting
      );
      return true;
    case 'paragraphSpacing':
      session.setParagraphAttrs(range, paragraphSpacingAttrs(action), suggesting);
      return true;
    case 'applyStyle':
      session.applyParagraphStyle(range, action.value, suggesting);
      return true;
    case 'fontFamily':
      return each((target) =>
        session.formatRange(target, { fontFamily: { ascii: action.value, hAnsi: action.value } })
      );
    case 'fontSize':
      if (
        !cells &&
        ranges.length > 0 &&
        effectiveFontSize(context) === action.value * 2 &&
        (context.fontSize != null || !hasExplicitFontSize(session, selection))
      ) {
        return false;
      }
      return each((target) => session.formatRange(target, { fontSize: action.value }));
    case 'textColor':
      return each((target) => session.formatRange(target, { color: textColorDelta(action.value) }));
    case 'highlightColor':
      return each((target) =>
        session.formatRange(target, {
          highlight: action.value && action.value !== 'none' ? action.value : null,
        })
      );
  }
}

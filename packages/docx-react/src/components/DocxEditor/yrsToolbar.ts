import type {
  YrsInputPositionMap,
  YrsAuthor,
  YrsParagraphAttrs,
  YrsSelectionContext,
  YrsSession,
  YrsStoryRange,
} from '@betteroffice/docx/yrs';
import { compareYrsLocs } from '@betteroffice/docx/yrs';
import type { FormattingAction } from '../Toolbar';
import type { YrsStoredFormatting, YrsStoredFormattingAction } from './YrsInput';
import type { TableContextInfo } from './types';
import { currentYrsTableContext } from './yrsCommands';

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

function setList(
  session: YrsSession,
  selection: YrsToolbarSelection,
  numId: 1 | 2,
  suggesting?: YrsAuthor
): void {
  const current = paragraphNumPr(selection.context.paragraphProperties.numPr);
  const remove = current?.numId === numId;
  forEachSelectedParagraph(session, selection, (range, properties) => {
    const existing = paragraphNumPr(properties.numPr);
    const attrs: YrsParagraphAttrs = remove
      ? {
          numPr: null,
          listIsBullet: null,
          listNumFmt: null,
          listMarker: null,
        }
      : {
          numPr: { numId, ilvl: existing?.ilvl ?? 0 },
          listIsBullet: numId === 1,
          listNumFmt: numId === 1 ? null : 'decimal',
          listMarker: null,
        };
    session.setParagraphAttrs(range, attrs, suggesting);
  });
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
        numPr: null,
        listIsBullet: null,
        listNumFmt: null,
        listMarker: null,
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
    session.setParagraphAttrs(range, { indentLeft: next > 0 ? next : null }, suggesting);
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

/** Translate a collapsed inline toolbar action into PM-style stored formatting. */
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
        fontFamily: null,
        fontSize: null,
        color: null,
      }
    : selection.context;
  const fontFamily = delta.fontFamily;
  const color = delta.color;
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
    },
  };
}

/**
 * Apply one toolbar action to the live yrs selection.
 *
 * Returns false for commands outside the current yrs formatting surface. The
 * caller deliberately does not fall back to PM while yrs-authoritative input
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

  if (action === 'bold' || action === 'italic' || action === 'underline') {
    if (isCollapsed(range)) return false;
    session.toggleMark(range, { type: action });
    return true;
  }
  if (action === 'strikethrough') {
    if (isCollapsed(range)) return false;
    session.formatRange(range, { strike: context.strike === true ? false : true });
    return true;
  }
  if (action === 'clearFormatting') {
    if (isCollapsed(range)) return false;
    session.clearFormatting(range);
    return true;
  }
  if (action === 'bulletList' || action === 'numberedList') {
    setList(session, selection, action === 'bulletList' ? 1 : 2, suggesting);
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
    case 'applyStyle':
      session.applyParagraphStyle(range, action.value, suggesting);
      return true;
    case 'fontFamily':
      if (isCollapsed(range)) return false;
      session.formatRange(range, {
        fontFamily: { ascii: action.value, hAnsi: action.value },
      });
      return true;
    case 'fontSize':
      if (isCollapsed(range)) return false;
      if (effectiveFontSize(context) === action.value * 2) return false;
      session.formatRange(range, { fontSize: action.value });
      return true;
    case 'textColor':
      if (isCollapsed(range)) return false;
      session.formatRange(range, { color: textColorDelta(action.value) });
      return true;
    case 'highlightColor':
      if (isCollapsed(range)) return false;
      session.formatRange(range, {
        highlight: action.value && action.value !== 'none' ? action.value : null,
      });
      return true;
  }
}

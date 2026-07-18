import { useCallback } from 'react';
import type {
  LineSpacingRule,
  ParagraphAlignment,
  ParagraphFormatting,
  ColorValue,
  Theme,
  TabStop,
} from '@betteroffice/docx/types/document';
import type { SelectionState, TableContextInfo } from '../types';
import { createStyleResolver } from '@betteroffice/docx/styles';
import { resolveColorToHex } from '@betteroffice/docx/utils';
import type { SelectionFormatting } from '../../Toolbar';
import type { YrsToolbarSelection } from '../yrsToolbar';

interface PmImageContext {
  pos: number;
  wrapType: string;
  displayMode: string;
  cssFloat: string | null;
  transform: string | null;
  alt: string | null;
  borderWidth: number | null;
  borderColor: string | null;
  borderStyle: string | null;
  width: number | null;
  height: number | null;
}

interface BorderSpec {
  style: string;
  size: number;
  color: { rgb: string };
}

const PARAGRAPH_ALIGNMENTS = new Set<ParagraphAlignment>([
  'left',
  'center',
  'right',
  'both',
  'distribute',
  'mediumKashida',
  'highKashida',
  'lowKashida',
  'thaiDistribute',
]);

const LINE_SPACING_RULES = new Set<LineSpacingRule>(['auto', 'exact', 'atLeast']);

function finiteNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function yrsSelectionState(selection: YrsToolbarSelection): SelectionState {
  const { context } = selection;
  const properties = context.paragraphProperties;
  const paragraphFormatting: ParagraphFormatting = {};
  if (
    typeof context.alignment === 'string' &&
    PARAGRAPH_ALIGNMENTS.has(context.alignment as ParagraphAlignment)
  ) {
    paragraphFormatting.alignment = context.alignment as ParagraphAlignment;
  }
  paragraphFormatting.indentLeft = finiteNumber(properties.indentLeft);
  paragraphFormatting.indentRight = finiteNumber(properties.indentRight);
  paragraphFormatting.indentFirstLine = finiteNumber(properties.indentFirstLine);
  paragraphFormatting.spaceBefore = finiteNumber(properties.spaceBefore);
  paragraphFormatting.spaceAfter = finiteNumber(properties.spaceAfter);
  paragraphFormatting.lineSpacing = finiteNumber(properties.lineSpacing);
  if (
    typeof properties.lineSpacingRule === 'string' &&
    LINE_SPACING_RULES.has(properties.lineSpacingRule as LineSpacingRule)
  ) {
    paragraphFormatting.lineSpacingRule = properties.lineSpacingRule as LineSpacingRule;
  }
  if (typeof properties.hangingIndent === 'boolean') {
    paragraphFormatting.hangingIndent = properties.hangingIndent;
  }
  if (typeof properties.bidi === 'boolean') paragraphFormatting.bidi = properties.bidi;
  const numPr = properties.numPr;
  if (numPr && typeof numPr === 'object') {
    const numId = finiteNumber(numPr.numId);
    const ilvl = finiteNumber(numPr.ilvl);
    if (numId != null || ilvl != null) paragraphFormatting.numPr = { numId, ilvl };
  }

  const color = context.color;
  const textColor: ColorValue | undefined =
    color == null
      ? undefined
      : /^#?[0-9a-f]{6}$/i.test(color)
        ? { rgb: color.replace(/^#/, '') }
        : { themeColor: color as ColorValue['themeColor'] };

  return {
    hasSelection: context.hasSelection,
    isMultiParagraph: context.isMultiParagraph,
    textFormatting: {
      bold: context.bold === true ? true : undefined,
      italic: context.italic === true ? true : undefined,
      underline: context.underline === true ? { style: 'single' } : undefined,
      strike: context.strike === true ? true : undefined,
      fontFamily: context.fontFamily
        ? { ascii: context.fontFamily, hAnsi: context.fontFamily }
        : undefined,
      fontSize: context.fontSize ?? undefined,
      color: textColor,
    },
    paragraphFormatting,
    styleId: context.styleId,
    startParagraphIndex: selection.startParagraphIndex,
    endParagraphIndex: selection.endParagraphIndex,
  };
}

/** Slice of EditorState that handleSelectionChange writes on every fire. */
export interface SelectionStateDelta {
  selectionFormatting: SelectionFormatting;
  paragraphIndentLeft?: number;
  paragraphIndentRight?: number;
  paragraphFirstLineIndent?: number;
  paragraphHangingIndent?: boolean;
  paragraphTabs?: TabStop[] | null;
  pmTableContext: TableContextInfo | null;
  pmImageContext: PmImageContext | null;
}

/**
 * Selection-change handler: extracts the formatting state ProseMirror
 * sees at the cursor, derives table + image context from the PM
 * selection, syncs the border-spec ref to the cell's actual color,
 * pushes the result into EditorState, refreshes the floating
 * add-comment button, and fans the SelectionState out to consumer-side
 * `onSelectionChange` + the bridge subscribers.
 *
 * Font/size fall back to the paragraph style's resolved values when no
 * explicit run-level mark is present — keeps the toolbar picker showing
 * the right value for unstyled cursor positions.
 */
export function useSelectionTracker({
  borderSpecRef,
  theme,
  historyStateRef,
  getCachedStyleResolver,
  setFloatingCommentBtn,
  applySelectionDelta,
  recomputeFloatingCommentBtn,
  onSelectionChange,
  selectionChangeSubscribersRef,
  canvasA11yNotifyRef,
}: {
  borderSpecRef: React.RefObject<BorderSpec>;
  theme: Theme | null | undefined;
  historyStateRef: React.RefObject<{ package: { styles?: unknown } } | null>;
  getCachedStyleResolver: (
    styles: Parameters<typeof createStyleResolver>[0]
  ) => ReturnType<typeof createStyleResolver>;
  setFloatingCommentBtn: React.Dispatch<React.SetStateAction<{ top: number; left: number } | null>>;
  applySelectionDelta: (delta: SelectionStateDelta) => void;
  recomputeFloatingCommentBtn: () => void;
  onSelectionChange: ((state: SelectionState | null) => void) | undefined;
  selectionChangeSubscribersRef: React.RefObject<Set<(s: SelectionState | null) => void>>;
  /**
   * Canvas-renderer live region notifier (assigned by CanvasA11yLiveRegion).
   * Called on every selection change — body and HF paths both funnel through
   * this handler — so canvas mode announces selection transitions; the
   * announcer edge-triggers internally, keeping uniform caret movement silent.
   */
  canvasA11yNotifyRef?: React.RefObject<(() => void) | null>;
}) {
  const handleSelectionChange = useCallback(
    (selectionState: SelectionState | null, authoritativeTableContext?: TableContextInfo | null) => {
      canvasA11yNotifyRef?.current?.();
      const pmTableCtx: TableContextInfo | null = authoritativeTableContext ?? null;

      // Sync borderSpecRef with the current cell's actual border color so
      // the toolbar's color/width pickers reflect the active cell.
      if (pmTableCtx?.cellBorderColor) {
        const rgb = resolveColorToHex(pmTableCtx.cellBorderColor, theme ?? undefined);
        if (rgb) {
          borderSpecRef.current = { ...borderSpecRef.current, color: { rgb } };
        }
      }

      if (!selectionState) {
        setFloatingCommentBtn(null);
        applySelectionDelta({
          selectionFormatting: {},
          pmTableContext: pmTableCtx,
          pmImageContext: null,
        });
        return;
      }

      const { textFormatting, paragraphFormatting } = selectionState;

      // Font/size fall back to the paragraph style's resolved values when no
      // explicit run-level mark is present.
      let fontFamily = textFormatting.fontFamily?.ascii || textFormatting.fontFamily?.hAnsi;
      let fontSize = textFormatting.fontSize;
      if (!fontFamily || !fontSize) {
        const currentDoc = historyStateRef.current;
        const paraStyleId = selectionState.styleId;
        if (currentDoc?.package.styles && paraStyleId) {
          const resolver = getCachedStyleResolver(
            currentDoc.package.styles as Parameters<typeof createStyleResolver>[0]
          );
          const resolved = resolver.resolveParagraphStyle(paraStyleId);
          if (!fontFamily && resolved.runFormatting?.fontFamily) {
            fontFamily =
              resolved.runFormatting.fontFamily.ascii || resolved.runFormatting.fontFamily.hAnsi;
          }
          if (!fontSize && resolved.runFormatting?.fontSize) {
            fontSize = resolved.runFormatting.fontSize;
          }
        }
      }

      const textColorHex = resolveColorToHex(textFormatting.color, theme ?? undefined);
      const textColor = textColorHex ? `#${textColorHex}` : undefined;

      // Build list state from numPr.
      const numPr = paragraphFormatting.numPr;
      const listState = numPr
        ? {
            type: (numPr.numId === 1 ? 'bullet' : 'numbered') as 'bullet' | 'numbered',
            level: numPr.ilvl ?? 0,
            isInList: true,
            numId: numPr.numId,
          }
        : undefined;

      const formatting: SelectionFormatting = {
        bold: textFormatting.bold,
        italic: textFormatting.italic,
        underline: !!textFormatting.underline,
        strike: textFormatting.strike,
        superscript: textFormatting.vertAlign === 'superscript',
        subscript: textFormatting.vertAlign === 'subscript',
        fontFamily,
        fontSize,
        color: textColor,
        highlight: textFormatting.highlight,
        alignment: paragraphFormatting.alignment,
        lineSpacing: paragraphFormatting.lineSpacing,
        listState,
        styleId: selectionState.styleId ?? undefined,
        indentLeft: paragraphFormatting.indentLeft,
        bidi: !!paragraphFormatting.bidi,
      };

      applySelectionDelta({
        selectionFormatting: formatting,
        paragraphIndentLeft: paragraphFormatting.indentLeft ?? 0,
        paragraphIndentRight: paragraphFormatting.indentRight ?? 0,
        paragraphFirstLineIndent: paragraphFormatting.indentFirstLine ?? 0,
        paragraphHangingIndent: paragraphFormatting.hangingIndent ?? false,
        paragraphTabs: paragraphFormatting.tabs ?? null,
        pmTableContext: pmTableCtx,
        pmImageContext: null,
      });

      recomputeFloatingCommentBtn();

      onSelectionChange?.(selectionState);
      // Fan out to bridge subscribers.
      for (const cb of selectionChangeSubscribersRef.current) {
        try {
          cb(selectionState);
        } catch (e) {
          console.error('selectionChange subscriber threw:', e);
        }
      }
    },
    [
      borderSpecRef,
      theme,
      historyStateRef,
      getCachedStyleResolver,
      setFloatingCommentBtn,
      applySelectionDelta,
      recomputeFloatingCommentBtn,
      onSelectionChange,
      selectionChangeSubscribersRef,
      canvasA11yNotifyRef,
    ]
  );

  const handleYrsSelectionChange = useCallback(
    (selection: YrsToolbarSelection) => {
      handleSelectionChange(yrsSelectionState(selection), selection.tableContext);
    },
    [handleSelectionChange]
  );

  return { handleSelectionChange, handleYrsSelectionChange };
}

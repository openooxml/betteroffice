import type { YrsInlineFormatDelta } from './index';
import type {
  DocxClipboardFragment,
  DocxFragmentSegment,
  DocxTableFragment,
} from './clipboardFragment';

/** @internal The three variants a paste can be re-applied as. */
export type PasteFormattingOption = 'keep-source' | 'merge' | 'text-only';

const KEPT_FORMAT_KEYS = ['bold', 'italic', 'underline', 'strike'] as const;

function mergeFormat(format: YrsInlineFormatDelta | undefined): YrsInlineFormatDelta | undefined {
  if (!format) return undefined;
  const kept: YrsInlineFormatDelta = {};
  if (format.bold) kept.bold = format.bold;
  if (format.italic) kept.italic = format.italic;
  if (format.underline) kept.underline = format.underline;
  if (format.strike) kept.strike = format.strike;
  return Object.keys(kept).length > 0 ? kept : undefined;
}

function mergeAttributes(attributes: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  if (!attributes) return undefined;
  const kept: Record<string, unknown> = {};
  for (const key of [...KEPT_FORMAT_KEYS, 'hyperlink']) {
    if (attributes[key] !== undefined && attributes[key] !== null && attributes[key] !== false) kept[key] = attributes[key];
  }
  return Object.keys(kept).length > 0 ? kept : undefined;
}

function mergeSegment(segment: DocxFragmentSegment): DocxFragmentSegment {
  switch (segment.kind) {
    case 'text': {
      const format = mergeFormat(segment.format);
      const attributes = mergeAttributes(segment.attributes);
      return {
        kind: 'text',
        text: segment.text,
        ...(attributes ? { attributes } : {}),
        ...(format ? { format } : {}),
      };
    }
    case 'pilcrow': {
      const properties = segment.properties ? { ...segment.properties } : undefined;
      if (properties) {
        delete properties.pStyle;
        delete properties.styleId;
        delete properties.pStyleId;
      }
      const attributes = mergeAttributes(segment.attributes);
      return {
        kind: 'pilcrow',
        ...(properties && Object.keys(properties).length > 0 ? { properties } : {}),
        ...(attributes ? { attributes } : {}),
      };
    }
    case 'table':
      return {
        kind: 'table',
        table: mergeTable(segment.table),
        ...(segment.attributes ? { attributes: { ...segment.attributes } } : {}),
      };
    default:
      return { ...segment };
  }
}

function mergeTable(table: DocxTableFragment): DocxTableFragment {
  return {
    tblPr: { ...table.tblPr },
    grid: [...table.grid],
    rows: table.rows.map(row => ({
      trPr: { ...row.trPr },
      cells: row.cells.map(cell => ({
        tcPr: { ...cell.tcPr },
        content: cell.content.map(mergeSegment),
        ...(cell.shading !== undefined ? { shading: cell.shading } : {}),
      })),
    })),
  };
}

function textOnlySegment(segment: DocxFragmentSegment): DocxFragmentSegment {
  switch (segment.kind) {
    case 'text':
      return { kind: 'text', text: segment.text };
    case 'pilcrow':
      return { kind: 'pilcrow' };
    case 'table':
      return { kind: 'table', table: textOnlyTable(segment.table) };
    default:
      return { kind: segment.kind };
  }
}

function textOnlyTable(table: DocxTableFragment): DocxTableFragment {
  return {
    tblPr: { ...table.tblPr },
    grid: [...table.grid],
    rows: table.rows.map(row => ({
      trPr: { ...row.trPr },
      cells: row.cells.map(cell => ({
        tcPr: { ...cell.tcPr },
        content: cell.content.map(textOnlySegment),
        ...(cell.shading !== undefined ? { shading: cell.shading } : {}),
      })),
    })),
  };
}

function cloneFragment(fragment: DocxClipboardFragment): DocxClipboardFragment {
  return structuredClone(fragment) as DocxClipboardFragment;
}

/**
 * @internal Merge-formatting variant of a pasted fragment: bold, italic,
 * underline, strikethrough, hyperlinks and list numbering survive; font
 * family, size, colour, highlight and paragraph style are dropped so the
 * destination paragraph's formatting applies.
 */
export function mergeFormattingFragment(fragment: DocxClipboardFragment): DocxClipboardFragment {
  const source = cloneFragment(fragment);
  return {
    paragraphs: source.paragraphs.map(paragraph => ({
      ...(paragraph.tables ? { tables: paragraph.tables.map(mergeTable) } : {}),
      text: paragraph.text,
      ...(paragraph.inlines ? { inlines: paragraph.inlines } : {}),
      ...(paragraph.runs
        ? {
            runs: paragraph.runs.map(run => {
              const format = mergeFormat(run.format);
              return {
                text: run.text,
                ...(format ? { format } : {}),
                ...(run.href ? { href: run.href } : {}),
              };
            }),
          }
        : {}),
      ...(paragraph.list ? { list: paragraph.list } : {}),
    })),
  };
}

/**
 * @internal Text-only variant of a pasted fragment: one plain paragraph per
 * source paragraph, with no run, paragraph or list formatting of its own.
 */
export function textOnlyFragment(fragment: DocxClipboardFragment): DocxClipboardFragment {
  const source = cloneFragment(fragment);
  return {
    paragraphs: source.paragraphs.map(paragraph => ({
      ...(paragraph.tables ? { tables: paragraph.tables.map(textOnlyTable) } : {}),
      text: paragraph.text,
    })),
  };
}

/** @internal Whether a fragment carries any formatting a paste variant could change. */
export function clipboardFragmentHasFormatting(fragment: DocxClipboardFragment | null): boolean {
  if (!fragment) return false;
  const segmentHas = (segment: DocxFragmentSegment): boolean => {
    if (segment.kind === 'text') {
      return (segment.format !== undefined && Object.keys(segment.format).length > 0) ||
        (segment.attributes !== undefined && Object.keys(segment.attributes).length > 0);
    }
    if (segment.kind === 'table') return segment.table.rows.some(row => row.cells.some(cell => cell.content.some(segmentHas)));
    return false;
  };
  return fragment.paragraphs.some(paragraph =>
    paragraph.pStyle !== undefined ||
    paragraph.list !== undefined ||
    (paragraph.tables ?? []).length > 0 ||
    (paragraph.inlines ?? []).some(inline => 'image' in inline) ||
    (paragraph.runs ?? []).some(run =>
      run.href !== undefined || (run.format !== undefined && Object.keys(run.format).length > 0)) ||
    (paragraph.tables ?? []).some(table => table.rows.some(row => row.cells.some(cell => cell.content.some(segmentHas)))));
}

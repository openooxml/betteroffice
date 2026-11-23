import {
  MAX_PASTE_FORMATTED_RUNS,
  clipboardFragmentHasPictures,
  clipboardFragmentRunCount,
  validClipboardFragment,
  type DocxClipboardFragment,
  type DocxTableFragment,
} from '@betteroffice/docx/yrs';
import { clipboardFragmentHasFormatting, type PasteFormattingOption } from './clipboardFragmentTransforms';

/**
 * @internal The Paste Options a paste offers: none unless it has formatting to choose about within the run limit, text to keep and no
 * pictures, and Keep Text Only only when its text-only fragment is one the worker accepts.
 */
export function pasteOptionsOffered(fragment: DocxClipboardFragment): PasteFormattingOption[] {
  if (clipboardFragmentHasPictures(fragment) || clipboardFragmentRunCount(fragment) > MAX_PASTE_FORMATTED_RUNS ||
    !clipboardFragmentHasFormatting(fragment)) return [];
  const plain = clipboardFragmentPlainText(fragment);
  if (plain.trim() === '') return [];
  return validClipboardFragment(plainTextClipboardFragment(plain)) ? ['keep-source', 'merge', 'text-only'] : ['keep-source', 'merge'];
}

/** @internal Fragment for a plain-text string, one paragraph per line. */
export function plainTextClipboardFragment(plain: string): DocxClipboardFragment {
  const lines = plain.split(/\r\n|[\r\n\u000B\u{2028}\u{2029}]/u);
  return { paragraphs: lines.map((text) => ({ text })) };
}

function cellFragmentText(content: DocxTableFragment['rows'][number]['cells'][number]['content']): string {
  let text = '';
  let breakBefore = false;
  const append = (piece: string) => {
    if (piece === '') return;
    if (breakBefore) text += '\n';
    breakBefore = false;
    text += piece;
  };
  content.forEach((segment, index) => {
    if (segment.kind === 'table') {
      breakBefore = text !== '';
      append(tableFragmentText(segment.table));
      breakBefore = true;
    } else if (segment.kind === 'pilcrow') {
      const last = index === content.length - 1;
      if (!last && content[index - 1]?.kind !== 'table' && content[index + 1]?.kind !== 'table') append(' ');
    } else append(segment.kind === 'text' ? segment.text : segment.kind === 'tab' ? '\t' : ' ');
  });
  return text;
}

function tableFragmentText(table: DocxTableFragment): string {
  return table.rows.map((row) => row.cells.map((cell) => cellFragmentText(cell.content)).join('\t')).join('\n');
}

/** @internal Plain text a fragment was pasted from: tables row by row with tab-separated cells, then the paragraph text. */
export function clipboardFragmentPlainText(fragment: DocxClipboardFragment): string {
  return fragment.paragraphs.map((paragraph) =>
    [...(paragraph.tables ?? []).map(tableFragmentText), ...(paragraph.text !== '' || !paragraph.tables?.length ? [paragraph.text] : [])].join('\n')
  ).join('\n');
}

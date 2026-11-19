import {
  clipboardFragmentHasPictures,
  validClipboardFragment,
  type DocxClipboardFragment,
  type DocxTableFragment,
} from '@betteroffice/docx/yrs';
import { clipboardFragmentHasFormatting, type PasteFormattingOption } from './clipboardFragmentTransforms';

/**
 * @internal The Paste Options a paste offers: none unless it has formatting to choose about, text to keep and no
 * pictures, and Keep Text Only only when its text-only fragment is one the worker accepts.
 */
export function pasteOptionsOffered(fragment: DocxClipboardFragment): PasteFormattingOption[] {
  if (clipboardFragmentHasPictures(fragment) || !clipboardFragmentHasFormatting(fragment)) return [];
  const plain = clipboardFragmentPlainText(fragment);
  if (plain.trim() === '') return [];
  return validClipboardFragment(plainTextClipboardFragment(plain)) ? ['keep-source', 'merge', 'text-only'] : ['keep-source', 'merge'];
}

/** @internal Fragment for a plain-text string, one paragraph per line. */
export function plainTextClipboardFragment(plain: string): DocxClipboardFragment {
  const lines = plain.split(/\r\n|[\r\n\u000B\u{2028}\u{2029}]/u);
  return { paragraphs: lines.map((text) => ({ text })) };
}

function tableFragmentText(table: DocxTableFragment): string {
  return table.rows.map((row) => row.cells.map((cell) => cell.content.map((segment) =>
    segment.kind === 'text' ? segment.text
      : segment.kind === 'tab' ? '\t'
        : segment.kind === 'table' ? tableFragmentText(segment.table)
          : ' '
  ).join('').trim()).join('\t')).join('\n');
}

/** @internal Plain text a fragment was pasted from: tables row by row with tab-separated cells, then the paragraph text. */
export function clipboardFragmentPlainText(fragment: DocxClipboardFragment): string {
  return fragment.paragraphs.map((paragraph) =>
    [...(paragraph.tables ?? []).map(tableFragmentText), ...(paragraph.text !== '' || !paragraph.tables?.length ? [paragraph.text] : [])].join('\n')
  ).join('\n');
}

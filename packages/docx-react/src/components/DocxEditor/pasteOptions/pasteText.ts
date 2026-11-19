import {
  clipboardFragmentHasFormatting,
  clipboardFragmentHasPictures,
  type DocxClipboardFragment,
  type DocxTableFragment,
} from '@betteroffice/docx/yrs';

/** @internal Whether a paste offers Paste Options: it has formatting to choose about, text to keep and no pictures. */
export function pasteOptionsApply(fragment: DocxClipboardFragment, plainText: string): boolean {
  return plainText.trim() !== '' && !clipboardFragmentHasPictures(fragment) && clipboardFragmentHasFormatting(fragment);
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

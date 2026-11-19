import { describe, expect, test } from 'bun:test';
import { validClipboardFragment, type DocxClipboardFragment } from './clipboardFragment';
import {
  clipboardFragmentHasFormatting,
  mergeFormattingFragment,
  textOnlyFragment,
} from './clipboardFragmentTransforms';

function richFragment(): DocxClipboardFragment {
  return {
    paragraphs: [
      {
        text: 'Rich link',
        runs: [
          {
            text: 'Rich ',
            format: {
              other: {
                bold: true,
                italic: true,
                underline: { style: 'single' },
                strike: { double: false },
                textColor: { rgb: 'FF0000' },
                highlight: 'FFFF00',
                fontSize: { size: 40, sizeCs: 40 },
                fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' },
              },
            },
          },
          { text: 'link', format: { other: { bold: true, textColor: { rgb: '0000FF' } } }, href: 'https://example.com' },
        ],
      },
      { text: 'Second', runs: [{ text: 'Second', format: { other: {} } }] },
    ],
  };
}

describe('mergeFormattingFragment', () => {
  test('keeps bold italic underline strike and links but drops font, size, colour and highlight', () => {
    const merged = mergeFormattingFragment(richFragment());
    const paragraph = merged.paragraphs[0]!;
    expect(paragraph.runs![0]!.format).toEqual({
      other: { bold: true, italic: true, underline: { style: 'single' }, strike: { double: false } },
    });
    expect(paragraph.runs![1]).toEqual({ text: 'link', format: { other: { bold: true } }, href: 'https://example.com' });
    expect(merged.paragraphs[1]!.runs).toEqual([{ text: 'Second', format: { other: {} } }]);
    expect(JSON.stringify(merged)).not.toContain('Georgia');
    expect(JSON.stringify(merged)).not.toContain('FF0000');
    expect(JSON.stringify(merged)).not.toContain('0000FF');
    expect(JSON.stringify(merged)).not.toContain('FFFF00');
    expect(validClipboardFragment(merged)).toBe(true);
  });

  test('keeps the tooltip of a linked run', () => {
    const linked: DocxClipboardFragment = { paragraphs: [{ text: 'Link', runs: [
      { text: 'Link', format: { other: { bold: true } }, href: 'https://example.test/', tooltip: 'Tip' },
    ] }] };
    expect(mergeFormattingFragment(linked).paragraphs[0]!.runs).toEqual([
      { text: 'Link', format: { other: { bold: true } }, href: 'https://example.test/', tooltip: 'Tip' },
    ]);
    expect(textOnlyFragment(linked).paragraphs).toEqual([{ text: 'Link' }]);
  });

  test('does not mutate the source fragment', () => {
    const source = richFragment();
    mergeFormattingFragment(source);
    expect(source.paragraphs[0]!.runs![0]!.format?.other?.fontSize).toEqual({ size: 40, sizeCs: 40 });
  });

  test('strips table cell run formatting except the kept marks', () => {
    const fragment: DocxClipboardFragment = {
      paragraphs: [{
        text: '',
        tables: [{
          tblPr: {},
          grid: [1000],
          rows: [{
            trPr: {},
            cells: [{
              tcPr: {},
              content: [
                { kind: 'text', text: 'Cell', format: { bold: true, fontSize: 20, color: { rgb: 'FF0000' } } },
                { kind: 'pilcrow', properties: { pStyle: 'Heading1', alignment: 'center' } },
              ],
            }],
          }],
        }],
      }],
    };
    const merged = mergeFormattingFragment(fragment);
    const content = merged.paragraphs[0]!.tables![0]!.rows[0]!.cells[0]!.content;
    expect(content[0]).toEqual({ kind: 'text', text: 'Cell', format: { bold: true } });
    expect(content[1]).toEqual({ kind: 'pilcrow', properties: { alignment: 'center' } });
  });

  test('keeps the marks an in-editor copy carries as other run attributes', () => {
    const merged = mergeFormattingFragment(editorCopyFragment());
    expect(merged.paragraphs[0]!.runs).toEqual([
      { text: 'Marked', format: { other: { bold: true, italic: true, underline: { style: 'single' }, strike: { double: false } } } },
      { text: ' plain', format: { other: {} } },
    ]);
    expect(JSON.stringify(merged)).not.toContain('FF0000');
    expect(JSON.stringify(merged)).not.toContain('Georgia');
  });
});

function editorCopyFragment(): DocxClipboardFragment {
  return {
    paragraphs: [{
      text: 'Marked plain',
      runs: [
        {
          text: 'Marked',
          format: {
            other: {
              bold: true,
              italic: true,
              underline: { style: 'single' },
              strike: { double: false },
              textColor: { rgb: 'FF0000' },
              fontFamily: { ascii: 'Georgia' },
            },
          },
        },
        { text: ' plain', format: { other: {} } },
      ],
    }],
  };
}

describe('textOnlyFragment', () => {
  test('keeps one plain paragraph per source paragraph with no formatting', () => {
    const textOnly = textOnlyFragment(richFragment());
    expect(textOnly.paragraphs).toEqual([{ text: 'Rich link' }, { text: 'Second' }]);
  });

  test('removes table cell formatting but keeps cell text', () => {
    const fragment: DocxClipboardFragment = {
      paragraphs: [{
        text: '',
        tables: [{
          tblPr: {},
          grid: [1000],
          rows: [{
            trPr: {},
            cells: [{
              tcPr: {},
              content: [
                { kind: 'text', text: 'Cell', format: { bold: true } },
                { kind: 'pilcrow', properties: { pStyle: 'Heading1' } },
              ],
            }],
          }],
        }],
      }],
    };
    const textOnly = textOnlyFragment(fragment);
    const content = textOnly.paragraphs[0]!.tables![0]!.rows[0]!.cells[0]!.content;
    expect(content).toEqual([{ kind: 'text', text: 'Cell' }, { kind: 'pilcrow' }]);
  });

  test('does not mutate the source fragment', () => {
    const source = richFragment();
    textOnlyFragment(source);
    expect(source.paragraphs[0]!.runs).toHaveLength(2);
  });
});

describe('clipboardFragmentHasFormatting', () => {
  test('detects run formatting and links', () => {
    expect(clipboardFragmentHasFormatting(richFragment())).toBe(true);
    expect(clipboardFragmentHasFormatting({
      paragraphs: [{ text: 'x', runs: [{ text: 'x', href: 'https://example.com' }] }],
    })).toBe(true);
  });

  test('is false for plain text, empty fragments and null', () => {
    expect(clipboardFragmentHasFormatting({ paragraphs: [{ text: 'plain' }] })).toBe(false);
    expect(clipboardFragmentHasFormatting({ paragraphs: [{ text: 'x', runs: [{ text: 'x' }] }] })).toBe(false);
    expect(clipboardFragmentHasFormatting(null)).toBe(false);
  });

  test('counts formatting an in-editor copy carries only as other run attributes', () => {
    expect(clipboardFragmentHasFormatting(editorCopyFragment())).toBe(true);
    expect(clipboardFragmentHasFormatting({
      paragraphs: [{ text: 'x', runs: [{ text: 'x', format: { other: { bold: true } } }] }],
    })).toBe(true);
  });

  test('is false for an in-editor copy whose runs carry no attributes', () => {
    expect(clipboardFragmentHasFormatting({
      paragraphs: [{ text: 'ab', runs: [{ text: 'a', format: { other: {} } }, { text: 'b', format: { other: {} } }] }],
    })).toBe(false);
  });
});

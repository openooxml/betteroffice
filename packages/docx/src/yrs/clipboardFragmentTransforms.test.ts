import { describe, expect, test } from 'bun:test';
import type { DocxClipboardFragment } from './clipboardFragment';
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
        pStyle: 'Heading1',
        list: 'numbered',
        runs: [
          {
            text: 'Rich ',
            format: {
              bold: true,
              italic: true,
              underline: true,
              strike: true,
              color: { rgb: 'FF0000' },
              highlight: 'FFFF00',
              fontSize: 20,
              fontFamily: { ascii: 'Georgia' },
            },
          },
          { text: 'link', format: { bold: true, color: { rgb: '0000FF' } }, href: 'https://example.com' },
        ],
      },
      { text: 'Second', runs: [{ text: 'Second' }] },
    ],
  };
}

describe('mergeFormattingFragment', () => {
  test('keeps bold italic underline strike, links and lists but drops font, size, colour, highlight and paragraph style', () => {
    const merged = mergeFormattingFragment(richFragment());
    const paragraph = merged.paragraphs[0]!;
    expect(paragraph.pStyle).toBeUndefined();
    expect(paragraph.list).toBe('numbered');
    expect(paragraph.runs![0]!.format).toEqual({ bold: true, italic: true, underline: true, strike: true });
    expect(paragraph.runs![1]).toEqual({ text: 'link', format: { bold: true }, href: 'https://example.com' });
    expect(JSON.stringify(merged)).not.toContain('Georgia');
    expect(JSON.stringify(merged)).not.toContain('FF0000');
    expect(JSON.stringify(merged)).not.toContain('FFFF00');
  });

  test('does not mutate the source fragment', () => {
    const source = richFragment();
    mergeFormattingFragment(source);
    expect(source.paragraphs[0]!.pStyle).toBe('Heading1');
    expect(source.paragraphs[0]!.runs![0]!.format?.fontSize).toBe(20);
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
});

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
  test('detects run formatting, paragraph style, links and lists', () => {
    expect(clipboardFragmentHasFormatting(richFragment())).toBe(true);
    expect(clipboardFragmentHasFormatting({ paragraphs: [{ text: 'x', pStyle: 'Heading1' }] })).toBe(true);
    expect(clipboardFragmentHasFormatting({ paragraphs: [{ text: 'x', list: 'bullet' }] })).toBe(true);
    expect(clipboardFragmentHasFormatting({
      paragraphs: [{ text: 'x', runs: [{ text: 'x', href: 'https://example.com' }] }],
    })).toBe(true);
  });

  test('is false for plain text, empty fragments and null', () => {
    expect(clipboardFragmentHasFormatting({ paragraphs: [{ text: 'plain' }] })).toBe(false);
    expect(clipboardFragmentHasFormatting({ paragraphs: [{ text: 'x', runs: [{ text: 'x' }] }] })).toBe(false);
    expect(clipboardFragmentHasFormatting(null)).toBe(false);
  });
});

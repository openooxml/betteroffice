import { describe, expect, it } from 'bun:test';
import { fromTsv, toTsv } from './index';
import type { CellInput } from './index';

// build a grid of plain-text cells (nothing is a formula).
function text(rows: string[][]): CellInput[][] {
  return rows.map((row) => row.map((input) => ({ input, isFormula: false })));
}

describe('toTsv', () => {
  it('joins fields with tabs and rows with CRLF', () => {
    expect(
      toTsv(
        text([
          ['a', 'b'],
          ['c', 'd'],
        ])
      )
    ).toBe('a\tb\r\nc\td');
  });

  it('quotes fields containing tabs, newlines, or quotes', () => {
    expect(toTsv(text([['a\tb']]))).toBe('"a\tb"');
    expect(toTsv(text([['line1\nline2']]))).toBe('"line1\nline2"');
    expect(toTsv(text([['say "hi"']]))).toBe('"say ""hi"""');
  });

  it('leaves ordinary values unquoted', () => {
    expect(toTsv(text([['42', 'hello world']]))).toBe('42\thello world');
  });
});

describe('formula-injection guard', () => {
  it('defangs text that a spreadsheet would read as a formula', () => {
    expect(toTsv(text([['=1+1']]))).toBe("'=1+1");
    expect(toTsv(text([['+1']]))).toBe("'+1");
    expect(toTsv(text([['-1']]))).toBe("'-1");
    expect(toTsv(text([['@SUM']]))).toBe("'@SUM");
  });

  it('defangs a leading control character that could break the field', () => {
    expect(toTsv(text([['\tinjected']]))).toBe('"\'\tinjected"');
    expect(toTsv(text([['\rinjected']]))).toBe('"\'\rinjected"');
  });

  it('exports genuine formulas verbatim, without a guard', () => {
    const cells: CellInput[][] = [[{ input: '=SUM(A1:A3)', isFormula: true }]];
    expect(toTsv(cells)).toBe('=SUM(A1:A3)');
  });

  it('does not guard a hyphen mid-text or an empty cell', () => {
    expect(toTsv(text([['a-b', '']]))).toBe('a-b\t');
  });
});

describe('fromTsv', () => {
  it('splits fields and rows', () => {
    expect(fromTsv('a\tb\r\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('accepts LF-only rows', () => {
    expect(fromTsv('a\tb\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ]);
  });

  it('ignores a single trailing newline', () => {
    expect(fromTsv('a\tb\r\n')).toEqual([['a', 'b']]);
  });

  it('unquotes fields with embedded tabs, newlines, and doubled quotes', () => {
    expect(fromTsv('"a\tb"\tc')).toEqual([['a\tb', 'c']]);
    expect(fromTsv('"line1\nline2"')).toEqual([['line1\nline2']]);
    expect(fromTsv('"say ""hi"""')).toEqual([['say "hi"']]);
  });

  it('keeps empty trailing fields', () => {
    expect(fromTsv('a\t')).toEqual([['a', '']]);
  });
});

describe('round-trip', () => {
  it('preserves ordinary grids structurally', () => {
    const cells = text([
      ['name', 'score'],
      ['alice', '10'],
      ['bob', 'has\ttab'],
    ]);
    const back = fromTsv(toTsv(cells));
    expect(back).toEqual([
      ['name', 'score'],
      ['alice', '10'],
      ['bob', 'has\ttab'],
    ]);
  });

  it('round-trips a real formula back to its = text', () => {
    const cells: CellInput[][] = [[{ input: '=A1*2', isFormula: true }]];
    expect(fromTsv(toTsv(cells))).toEqual([['=A1*2']]);
  });

  it('round-trips defanged text with its guard quote intact', () => {
    // the guard is part of the literal text now; fromTsv does not strip it.
    expect(fromTsv(toTsv(text([['=danger']])))).toEqual([["'=danger"]]);
  });
});

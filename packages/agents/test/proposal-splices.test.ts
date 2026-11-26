import { expect, test } from 'bun:test';
import { MAX_SPLICE_CELLS, spliceChange } from '../src/proposal-splices';

const change = (oldText: string, newText: string, start = 0) => ({
  ref: 'p1', revision: 1, story: 'body', paragraph: 1, start, oldText, newText,
});

test('splices preserve the leftmost old tokens on ties', () => {
  expect(spliceChange(change('aXa', 'a', 10))).toEqual([change('Xa', '', 11)]);
  expect(spliceChange(change('abab', 'ab'))).toEqual([change('ab', '', 2)]);
  expect(spliceChange(change('X|red tan|Y', 'Z|tan red|W')))
    .toEqual([change('X', 'Z'), change('', 'tan ', 2), change(' tan', '', 5), change('Y', 'W', 10)]);
  expect(spliceChange(change('X|red|Y', 'Z|red blue red|W')))
    .toEqual([change('X', 'Z'), change('', ' blue red', 5), change('Y', 'W', 6)]);
});

test('splices maximize the UTF-16 length of retained tokens', () => {
  expect(spliceChange(change('a b KEEP', 'KEEP c d')))
    .toEqual([change('a b ', ''), change('', ' c d', 8)]);
  expect(spliceChange(change('X|red blue|Y', 'Z|blue red|W')))
    .toEqual([change('X', 'Z'), change('red ', '', 2), change('', ' red', 10), change('Y', 'W', 11)]);
  expect(spliceChange(change('X|KEEP 👩‍💻|Y', 'Z|👩‍💻 KEEP|W')))
    .toEqual([change('X', 'Z'), change('KEEP ', '', 2), change('', ' KEEP', 12), change('Y', 'W', 13)]);
});

test('splices weight a whitespace token once regardless of its length', () => {
  expect(spliceChange(change('X|KEEP     red|Y', 'Z|tan     KEEP|W')))
    .toEqual([change('X', 'Z'), change('', 'tan     ', 2), change('     red', '', 6), change('Y', 'W', 15)]);
});

test('splices merge swapped words across the shared space', () => {
  expect(spliceChange(change('alpha beta', 'BETA ALPHA')))
    .toEqual([change('alpha beta', 'BETA ALPHA')]);
  expect(spliceChange(change('a  b', 'A   B'))).toEqual([change('a  b', 'A   B')]);
});

test('splices remove brackets, retain the unchanged word and space, and merge a replacement with an abutting deletion', () => {
  expect(spliceChange(change('[MRID xxxxxxxx]', 'MRID xxxxxxxx')))
    .toEqual([change('[', ''), change(']', '', 14)]);
  expect(spliceChange(change('[MRID xxxxxxxx]', 'MRID 12345678')))
    .toEqual([change('[', ''), change('xxxxxxxx]', '12345678', 6)]);
});

test('splices keep punctuation inside numbers, dates, IDs and URL components', () => {
  for (const [oldToken, newToken, oldText, newText, offset] of [
    ['1,234.56', '1,923.56', '234', '923', 4],
    ['2024-01-31', '2024-02-31', '1', '2', 8],
    ['tenant-prod-123', 'tenant-dev-123', 'prod', 'dev', 9],
  ] as const) {
    expect(spliceChange(change(`x ${oldToken} a`, `y ${newToken} b`)))
      .toEqual([change('x', 'y'), change(oldText, newText, offset), change('a', 'b', oldToken.length + 3)]);
  }
  expect(spliceChange(change('x alpha.example a', 'y beta.example b')), 'The first gap is just a space and the second contains a.example')
    .toEqual([change('x alph', 'y bet'), change('a', 'b', 16)]);
  expect(spliceChange(change('x https://alpha.example/path a', 'y https://beta.example/path b')))
    .toEqual([change('x', 'y'), change('alph', 'bet', 10), change('a', 'b', 29)]);
});

test('splices trim shared text within a changed token', () => {
  expect(spliceChange(change('u cd.ab v', 'U EF-ab V')), 'The first gap is just a space and the second contains ab')
    .toEqual([change('u cd.', 'U EF-'), change('v', 'V', 8)]);
  expect(spliceChange(change('u cd.e\u0301z v', 'U EF-e\u0300z V')), 'The first gap is just a space and the second contains z')
    .toEqual([change('u cd.e\u0301', 'U EF-e\u0300'), change('v', 'V', 9)]);
});

test('splices retain words between parentheses and other non-joining punctuation', () => {
  for (const [left, right] of [['(', ')'], ['[', ']'], ['!', '?'], [';', '+']]) {
    expect(spliceChange(change(`x${left}KEEP${right}y`, `a${left}KEEP${right}b`)))
      .toEqual([change('x', 'a'), change('y', 'b', 7)]);
  }
});

test('splices keep CJK runs, marks, numbers and underscores inside word tokens', () => {
  expect(spliceChange(change('甲東京乙', '丙東京丁'))).toEqual([change('甲東京乙', '丙東京丁')]);
  expect(spliceChange(change('ae\u0301_27x', 'be\u0301_27y'))).toEqual([change('ae\u0301_27x', 'be\u0301_27y')]);
});

for (const [name, oldCluster, newCluster] of [
  ['emoji ZWJ sequence', '👩‍💻', '👨‍💻'],
  ['flag', '🇺🇸', '🇺🇳'],
  ['combining mark', 'e\u0301', 'e\u0300'],
  ['astral character', '😀', '😁'],
]) {
  test(`splices keep a ${name} whole`, () => {
    expect(spliceChange(change(`a${oldCluster}z`, `a${newCluster}z`)))
      .toEqual([change(oldCluster, newCluster, 1)]);
    expect(spliceChange(change(`x${oldCluster}|a`, `y${newCluster}|b`)))
      .toEqual([change(`x${oldCluster}`, `y${newCluster}`), change('a', 'b', oldCluster.length + 2)]);
    expect(spliceChange(change(`a|${oldCluster}|b`, `A|${oldCluster}|B`)))
      .toEqual([change('a', 'A'), change('b', 'B', oldCluster.length + 3)]);
  });
}

test('splices cap the trimmed token alignment table', () => {
  const overOld = `!K${'!'.repeat(98)}`;
  const overNew = `?K${'?'.repeat(9898)}`;
  expect((overOld.length + 1) * (overNew.length + 1)).toBe(MAX_SPLICE_CELLS + 1);
  expect(spliceChange(change(`p${overOld}s`, `p${overNew}s`))).toEqual([change(overOld, overNew, 1)]);

  const underOld = '!'.repeat(996);
  const underNew = '?'.repeat(998);
  expect((underOld.length + 3) * (underNew.length + 3)).toBe(MAX_SPLICE_CELLS - 1);
  expect(spliceChange(change(`p!K${underOld}s`, `p?K${underNew}s`)))
    .toEqual([change('!', '?', 1), change(underOld, underNew, 3)]);

  const oldWord = 'x'.repeat(1100), newWord = 'y'.repeat(1100);
  expect(spliceChange(change(`${oldWord} ${oldWord}`, `${newWord} ${newWord}`)), 'A shared space does not split the long tokens')
    .toEqual([change(`${oldWord} ${oldWord}`, `${newWord} ${newWord}`)]);
});

test('splices combine adjacent deletions and insertions and preserve pure edits', () => {
  expect(spliceChange(change('aXYZb', 'a123b'))).toEqual([change('XYZ', '123', 1)]);
  expect(spliceChange(change('a|X|b', 'Y!a||b?Z')))
    .toEqual([change('', 'Y!'), change('X', '', 2), change('', '?Z', 5)]);
  expect(spliceChange(change('a', 'a'))).toEqual([]);
});

test('splices split 33 replacements and retain every KEEP word', () => {
  const words = (word: string) => Array(33).fill(word).join(' KEEP ');
  expect(spliceChange(change(`p ${words('a')} z`, `p ${words('b')} z`)))
    .toEqual(Array.from({ length: 33 }, (_, i) => change('a', 'b', 2 + i * 7)));
});

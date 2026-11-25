import { expect, test } from 'bun:test';
import { MAX_SPLICE_CELLS, spliceChange } from '../src/proposal-splices';

const change = (oldText: string, newText: string, start = 0) => ({
  ref: 'p1', revision: 1, story: 'body', paragraph: 1, start, oldText, newText,
});

test('splices preserve the leftmost old tokens on ties', () => {
  expect(spliceChange(change('aXa', 'a', 10))).toEqual([change('Xa', '', 11)]);
  expect(spliceChange(change('abab', 'ab'))).toEqual([change('ab', '', 2)]);
  expect(spliceChange(change('X|red blue|Y', 'Z|blue red|W')))
    .toEqual([change('X', 'Z'), change('', 'blue ', 2), change(' blue', '', 5), change('Y', 'W', 11)]);
  expect(spliceChange(change('X|red|Y', 'Z|red blue red|W')))
    .toEqual([change('X', 'Z'), change('', ' blue red', 5), change('Y', 'W', 6)]);
});

test('splices keep shared space between swapped words', () => {
  expect(spliceChange(change('alpha beta', 'BETA ALPHA')))
    .toEqual([change('alpha', 'BETA'), change('beta', 'ALPHA', 6)]);
  expect(spliceChange(change('a  b', 'A   B'))).toEqual([change('a  b', 'A   B')]);
});

test('splices remove brackets and retain the unchanged word and space', () => {
  expect(spliceChange(change('[MRID xxxxxxxx]', 'MRID 12345678')))
    .toEqual([change('[', ''), change('xxxxxxxx', '12345678', 6), change(']', '', 14)]);
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
  expect(spliceChange(change(`${oldWord} ${oldWord}`, `${newWord} ${newWord}`)))
    .toEqual([change(oldWord, newWord), change(oldWord, newWord, oldWord.length + 1)]);
});

test('splices combine adjacent deletions and insertions and preserve pure edits', () => {
  expect(spliceChange(change('aXYZb', 'a123b'))).toEqual([change('XYZ', '123', 1)]);
  expect(spliceChange(change('a|X|b', 'Y!a||b?Z')))
    .toEqual([change('', 'Y!'), change('X', '', 2), change('', '?Z', 5)]);
  expect(spliceChange(change('a', 'a'))).toEqual([]);
});

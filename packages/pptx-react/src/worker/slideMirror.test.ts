import { expect, test } from 'bun:test';
import type { SlideDisplayList, SlidePrimitive } from '@betteroffice/pptx';
import { slideMirrorItems, slideObjectName, slideTitle } from './slideMirror';

const at = { objectId: 1, x: 10, y: 20, w: 30, h: 40 };
const text = (shapeId: string | undefined, ...lines: string[]): SlidePrimitive => ({
  ...at, kind: 'textBox', shapeId, anchor: 'top', lines: [],
  paragraphs: lines.map((line) => ({ level: 0, runs: line.split('|').map((part) => ({ text: part, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' })) })),
});
const shape = (shapeId: string, name = 'Rectangle 1'): SlidePrimitive => ({ ...at, kind: 'shape', shapeId, name, geometry: 'rect', path: [] });
const image = (shapeId: string | undefined, name = 'Picture 2'): SlidePrimitive => ({ ...at, kind: 'image', shapeId, name });
const list = (primitives: SlidePrimitive[], altText?: SlideDisplayList['altText']): SlideDisplayList =>
  ({ contractVersion: 1, width: 960, height: 540, primitives, altText });
const items = (value: SlideDisplayList) => slideMirrorItems(value, 'Picture');

test('a paragraph reads as one string across its runs', () => {
  const [item] = items(list([text('a', 'Quarterly |results| are in', '', 'Second\u000bline')]));
  expect(item!.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Quarterly results are in', 'Second line']);
  expect(item).toMatchObject({ x: 10, y: 20, w: 30, h: 40 });
});

test('a shape and its text are one item, and a shape with neither text nor alt text is none', () => {
  expect(items(list([shape('a'), text('a', 'Label'), shape('b')])).map((item) => item.key)).toEqual(['a']);
});

test('alt text labels pictures, and an undescribed picture falls back to its name', () => {
  const value = list([image('a'), image('b'), image('c', '')], { a: { title: 'Logo', description: 'A blue square' } });
  expect(items(value).map((item) => item.label)).toEqual(['Logo: A blue square', 'Picture 2', 'Picture']);
});

test('a decorative picture is skipped, but text in a decorative shape stays', () => {
  const value = list([image('a'), shape('b'), text('b', 'Kept')], { a: { decorative: true }, b: { decorative: true, description: 'Dropped' } });
  expect(items(value)).toMatchObject([{ key: 'b', label: null, paragraphs: [{ text: 'Kept' }] }]);
});

test('a described group speaks once, and its unlabelled pictures add nothing', () => {
  const value = list([image('a'), image('b'), text('c', 'Inside')], {
    outer: { description: 'Whole diagram' }, inner: { description: 'Left half', group: 'outer' },
    a: { group: 'inner' }, b: { group: 'inner' }, c: { group: 'outer' },
  });
  expect(items(value).map((item) => [item.key, item.label])).toEqual([['outer', 'Whole diagram'], ['inner', 'Left half'], ['c', null]]);
});

test('a picture painted without a shape id keeps its alt text, its place and one item', () => {
  const op = (name: string): SlidePrimitive => ({ ...at, kind: 'shape', name, geometry: 'custom', path: [] });
  const value = list([text('a', 'Before'), op('Vector'), op('Vector'), op('Bare'), text('b', 'After')],
    { v: { description: 'A described drawing', primitive: 1 }, w: { primitive: 3 } });
  expect(items(value).map((item) => [item.key, item.label])).toEqual([['a', null], ['v', 'A described drawing'], ['w', 'Bare'], ['b', null]]);
  expect(slideObjectName(value, 'w', 'Shape')).toBe('Bare');
});

test('a frame description reads once above the shapes it expands into', () => {
  const value = list([shape('f:0'), text('f:0', 'Manager'), shape('f:1'), text('f:1', 'Report')],
    { f: { description: 'Reporting lines' }, 'f:0': { group: 'f' }, 'f:1': { group: 'f' } });
  expect(items(value).map((item) => [item.key, item.label])).toEqual([['f', 'Reporting lines'], ['f:0', null], ['f:1', null]]);
});

test('table cells and chart summaries are reachable', () => {
  const table: SlidePrimitive = { ...at, kind: 'table', shapeId: 't', name: 'Table 1', label: 'Table, 1 rows, 2 columns', primitives: [text(undefined, 'A1'), text(undefined, 'B1')] };
  const chart: SlidePrimitive = { ...at, kind: 'chart', shapeId: 'c', name: 'Chart 1', label: 'Bar chart', primitives: [text(undefined, 'Axis')] };
  expect(items(list([table, chart]))).toMatchObject([
    { label: 'Table, 1 rows, 2 columns', paragraphs: [{ text: 'A1' }, { text: 'B1' }] },
    { label: 'Bar chart', paragraphs: [] },
  ]);
});

test('layout and master shapes read after the slide, and backgrounds not at all', () => {
  const value = list([image(undefined, 'Background'), text('master:m:0', 'Confidential'), text('layout:l:0', 'Footer'), text('a', 'Title')]);
  expect(items(value).map((item) => item.key)).toEqual(['a', 'master:m:0', 'layout:l:0']);
});

test('the title placeholder reads as one heading, wherever it sits in the stack', () => {
  const value = { ...list([text('a', 'Intro'), text('b', 'Two line', 'title')]), titleShape: 'b' };
  expect(items(value).flatMap((item) => item.paragraphs)).toEqual([
    { text: 'Intro', heading: false }, { text: 'Two line title', heading: true },
  ]);
  expect(slideTitle(value)).toBe('Two line title');
});

test('a slide with an empty or missing title has none', () => {
  expect(slideTitle(list([text('a', 'Body')]))).toBeNull();
  expect(slideTitle({ ...list([text('a', '')]), titleShape: 'a' })).toBeNull();
});

test('an announcement prefers alt text, then the name, then the text', () => {
  const value = list([shape('a'), text('a', 'Body'), text('b', 'Only text'), image('c', '')], { a: { description: 'Described' } });
  expect(slideObjectName(value, 'a', 'Shape')).toBe('Described');
  expect(slideObjectName(list([shape('a'), text('a', 'Body')]), 'a', 'Shape')).toBe('Rectangle 1');
  expect(slideObjectName(value, 'b', 'Shape')).toBe('Only text');
  expect(slideObjectName(value, 'c', 'Shape')).toBe('Shape');
  expect(slideObjectName(list([image('d')], { d: { decorative: true, description: 'Silent' } }), 'd', 'Shape')).toBe('Picture 2');
});

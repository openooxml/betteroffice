import { expect, test } from 'bun:test';
import type { DiagramSnapshot, PageDisplayList, PageLayer, PageSnapshot } from '@betteroffice/vsdx';
import { collectDiagnostics, selectionHiddenByLayers, stillSelectable } from './VsdxEditor';

const frame: PageDisplayList = {
  contractVersion: 4,
  width: 1,
  height: 1,
  paintTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
  primitives: [{ kind: 'textBox', id: 'text', zOrder: 0, x: 0, y: 0, width: 1, height: 1, paragraphs: [{ runs: [{ text: 'x', family: 'Arial', sizeIn: 12, bold: false, italic: false, underline: false, smallCaps: false, superscript: false, subscript: false, letterSpacing: 0, color: '#000', diagnostics: [{ category: 'integrity', code: 'missing-media', detail: '' }, { category: 'fidelity', code: 'font-substituted', detail: '' }] }]}], lines: [] }],
};

test('collects structured diagnostics without matching their text', () => {
  expect(collectDiagnostics(frame)).toEqual([
    { category: 'integrity', code: 'missing-media', detail: '' },
    { category: 'fidelity', code: 'font-substituted', detail: '' },
  ]);
});

test('collects defaulted paint diagnostics and tolerates shapes without them', () => {
  const shapes: PageDisplayList = {
    ...frame,
    primitives: [
      { kind: 'shape', id: 'resolved', zOrder: 0, path: [] },
      { kind: 'group', id: 'group', zOrder: 1, primitives: [{ kind: 'shape', id: 'defaulted', zOrder: 2, path: [], diagnostics: [{ category: 'fidelity', code: 'unresolvable-fill-colour', detail: 'unresolvable fill colour: missing colour cell FillForegnd' }] }] },
    ],
  };
  expect(collectDiagnostics(shapes)).toEqual([
    { category: 'fidelity', code: 'unresolvable-fill-colour', detail: 'unresolvable fill colour: missing colour cell FillForegnd' },
  ]);
});

function textBox(frame: PageDisplayList) {
  const box = frame.primitives[0];
  if (box.kind !== 'textBox') throw new Error('fixture must start with a text box');
  return box;
}

function withoutDiagnostics(): PageDisplayList {
  const { diagnostics: _omitted, ...run } = textBox(frame).paragraphs[0].runs[0];
  return { ...frame, primitives: [{ ...textBox(frame), paragraphs: [{ runs: [run] }] }] };
}

test('tolerates text runs that omit diagnostics', () => {
  expect(collectDiagnostics(withoutDiagnostics())).toEqual([]);
});

test('collects diagnostics alongside runs that omit them', () => {
  const mixed = withoutDiagnostics();
  textBox(mixed).paragraphs[0].runs.push({ ...textBox(frame).paragraphs[0].runs[0], diagnostics: [{ category: 'fidelity', code: 'font-substituted', detail: '' }] });
  expect(collectDiagnostics(mixed)).toEqual([{ category: 'fidelity', code: 'font-substituted', detail: '' }]);
});

const layerPage: PageSnapshot = {
  id: 'page:1',
  sourcePartPath: 'visio/pages/page1.xml',
  name: null,
  shapes: [{ id: 'page:1:shape:1', sourceId: 1, name: null, cells: [{ locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: 'LayerMember' }, name: 'LayerMember', formula: null, value: '1' }], children: [] }],
};

const layerSnapshot: DiagramSnapshot = { pages: [layerPage] };

function testLayers(): PageLayer[] {
  return [
    { index: 0, name: 'Trussing', visible: true, print: true, lock: false, active: false, color: '255', status: '0' },
    { index: 1, name: 'Lighting', visible: false, print: true, lock: false, active: false, color: '255', status: '0' },
  ];
}

test('a selection hidden by its layer is no longer selectable', () => {
  const selection = { pageId: 'page:1', shapeId: 'page:1:shape:1', hit: { kind: 'shape' as const, shapeId: 'x' } };
  expect(stillSelectable(layerSnapshot, 0, selection, testLayers())).toBe(false);
  expect(stillSelectable(layerSnapshot, 0, selection, testLayers().map((layer) => ({ ...layer, visible: true })))).toBe(true);
});

test('selection layer visibility matches the engine rule, including hidden groups', () => {
  const selection = { pageId: 'page:1', shapeId: 'page:1:shape:1', hit: { kind: 'shape' as const, shapeId: 'x' } };
  expect(selectionHiddenByLayers(layerPage, testLayers(), selection)).toBe(true);
  expect(selectionHiddenByLayers(layerPage, [], selection)).toBe(false);
  const grouped: PageSnapshot = {
    ...layerPage,
    shapes: [{
      id: 'page:1:shape:9', sourceId: 9, name: null,
      cells: [{ locator: { sheet: { page: 1 }, shapeId: 9, section: null, row: null, cellName: 'LayerMember' }, name: 'LayerMember', formula: null, value: '1' }],
      children: [{ id: 'page:1:shape:1', sourceId: 1, name: null, cells: [], children: [] }],
    }],
  };
  expect(selectionHiddenByLayers(grouped, testLayers(), selection)).toBe(true);
});

test('a shape on no layer stays selectable when layer zero is hidden', () => {
  const page: PageSnapshot = { ...layerPage, shapes: [{ id: 'page:1:shape:1', sourceId: 1, name: null, cells: [], children: [] }] };
  const snapshot: DiagramSnapshot = { pages: [page] };
  const selection = { pageId: 'page:1', shapeId: 'page:1:shape:1', hit: { kind: 'shape' as const, shapeId: 'x' } };
  const layers = testLayers().map((layer) => (layer.index === 0 ? { ...layer, visible: false } : layer));
  expect(selectionHiddenByLayers(page, layers, selection)).toBe(false);
  expect(stillSelectable(snapshot, 0, selection, layers)).toBe(true);
});

test('layer membership parsing matches the engine table', () => {
  const layers: PageLayer[] = [0, 1, 2, 3, 5, 9].map((index) => ({ index, name: `L${index}`, visible: false, print: true, lock: false, active: false, color: '', status: '' }));
  const hidden = (member: string) => {
    const page: PageSnapshot = {
      ...layerPage,
      shapes: [{ id: 'page:1:shape:1', sourceId: 1, name: null, cells: [{ locator: { sheet: { page: 1 }, shapeId: 1, section: null, row: null, cellName: 'LayerMember' }, name: 'LayerMember', formula: null, value: member }], children: [] }],
    };
    return selectionHiddenByLayers(page, layers, { pageId: 'page:1', shapeId: 'page:1:shape:1', hit: { kind: 'shape' as const, shapeId: 'x' } });
  };
  for (const [member, expected] of [['0;2', true], ['1;0;1', true], [' 2 ; 9 ', true], ['', false], ['a;3', true], ['3;', true], [';', false], ['-1;2', true], ['+1;2', true], ['4294967296;5', true], ['-1', false], ['a', false], ['4294967296', false], ['7', false]] as const) {
    expect(hidden(member)).toBe(expected);
  }
});

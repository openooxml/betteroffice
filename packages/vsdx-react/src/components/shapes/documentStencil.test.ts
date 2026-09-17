import { describe, expect, test } from 'bun:test';
import type { DocumentMaster, GeometryPathCommand, PageDisplayList } from '@betteroffice/vsdx';
import { documentMasterDraft, documentStencilEntries, masterPreviewPath } from './documentStencil';

function displayList(path: GeometryPathCommand[]): PageDisplayList {
  return {
    contractVersion: 7,
    width: 96,
    height: 96,
    printWidth: 96,
    printHeight: 96,
    paintTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    primitives: [{ kind: 'shape', id: 'master:1', zOrder: 0, path }],
  };
}

const line = [
  { type: 'move', x: 0, y: 4 },
  { type: 'line', x: 8, y: 4 },
] satisfies GeometryPathCommand[];

describe('documentMasterDraft', () => {
  test('writes placement only so the master keeps its dimensions', () => {
    const draft = documentMasterDraft(7, 4, 5);
    expect(draft.master).toBe(7);
    expect(draft.cells.map((cell) => cell.name)).toEqual(['PinX', 'PinY']);
    expect(draft.cells.map((cell) => cell.formula)).toEqual(['4', '5']);
  });

  test('falls back to the origin for non-finite points', () => {
    expect(documentMasterDraft(1, Number.NaN, Number.POSITIVE_INFINITY).cells.map((cell) => cell.formula)).toEqual(['0', '0']);
  });
});

describe('masterPreviewPath', () => {
  test('renders a horizontal line master instead of an empty preview', () => {
    const preview = masterPreviewPath(displayList(line));
    expect(preview).toBe('M 0 0.5 L 1 0.5');
  });

  test('renders a vertical line master', () => {
    const preview = masterPreviewPath(displayList([
      { type: 'move', x: 3, y: 0 },
      { type: 'line', x: 3, y: 2 },
    ]));
    expect(preview).toBe('M 0.5 0 L 0.5 1');
  });

  test('centres a wide master inside the unit preview box', () => {
    const preview = masterPreviewPath(displayList([
      { type: 'move', x: 0, y: 0 },
      { type: 'line', x: 4, y: 0 },
      { type: 'line', x: 4, y: 1 },
      { type: 'close' },
    ]));
    expect(preview).toBe('M 0 0.375 L 1 0.375 L 1 0.625 Z');
  });

  test('returns an empty path when there is no geometry', () => {
    expect(masterPreviewPath(displayList([]))).toBe('');
  });
});

describe('documentStencilEntries', () => {
  const masters: DocumentMaster[] = [
    { id: 1, name: 'Stencil-Rect', display: displayList(line) },
    { id: 2, name: null, display: null },
  ];

  test('labels masters by name and falls back to the id', () => {
    const entries = documentStencilEntries(masters);
    expect(entries.map((entry) => entry.id)).toEqual(['document-master-1', 'document-master-2']);
    expect(entries.map((entry) => entry.label)).toEqual(['Stencil-Rect', '#2']);
  });

  test('keeps the tile when a master fails to render', () => {
    const [rect, broken] = documentStencilEntries(masters);
    expect(rect.preview).not.toBe('');
    expect(broken.preview).toBe('');
  });

  test('sizes the tile from the rendered master, in inches', () => {
    const [rect, broken] = documentStencilEntries(masters);
    expect(rect.defaultSize).toEqual({ width: 1, height: 1 });
    expect(broken.defaultSize).toEqual({ width: 1, height: 1 });
  });

  test('drafts an instance of the master it came from', () => {
    const [rect] = documentStencilEntries(masters);
    expect(rect.draft(2, 3, 1, 1).master).toBe(1);
  });
});

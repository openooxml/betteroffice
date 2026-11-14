import { expect, test } from 'bun:test';
import type { DisplayList, DrawCmd } from '@betteroffice/xlsx';
import { cellEditorColors } from './cellEditorColors';

const cell = { x: 100, y: 20, w: 80, h: 20 };

function list(...commands: DrawCmd[]): DisplayList {
  return { width: 800, height: 600, commands };
}

const fill = (color: string): DrawCmd => ({ op: 'fillRect', x: 100, y: 20, w: 80, h: 20, color });
const text = (color: string): DrawCmd => ({ op: 'text', x: 103, y: 34, text: 'value', fontSize: 11, color, clip: cell });

test('an unstyled cell edits black on white', () => {
  expect(cellEditorColors(list(fill('#ffffff'), text('#000000')), cell)).toEqual({ color: '#000000', background: '#ffffff' });
  expect(cellEditorColors(null, cell)).toEqual({ color: '#000000', background: '#ffffff' });
});

test('the editor takes the cell font colour and fill', () => {
  expect(cellEditorColors(list(fill('#fff2cc'), text('#c00000')), cell)).toEqual({ color: '#c00000', background: '#fff2cc' });
});

test('white text on a dark fill keeps a white caret on the dark fill', () => {
  expect(cellEditorColors(list(fill('#1f3864'), text('#ffffff')), cell)).toEqual({ color: '#ffffff', background: '#1f3864' });
});

test('an empty dark cell gets a white caret', () => {
  expect(cellEditorColors(list(fill('#000000')), cell)).toEqual({ color: '#ffffff', background: '#000000' });
});

test('text that would vanish against the fill falls back to a contrasting colour', () => {
  expect(cellEditorColors(list(fill('#ffffff'), text('#fefefe')), cell)).toEqual({ color: '#000000', background: '#ffffff' });
});

test('translucent overlays and other cells do not colour the editor', () => {
  const other: DrawCmd = { op: 'fillRect', x: 300, y: 20, w: 80, h: 20, color: '#000000' };
  const band: DrawCmd = { op: 'fillRect', x: 0, y: 0, w: 800, h: 600, color: 'rgba(33, 115, 70, 0.12)' };
  const otherText: DrawCmd = { op: 'text', x: 303, y: 34, text: 'x', fontSize: 11, color: '#ff0000', clip: { x: 300, y: 20, w: 80, h: 20 } };
  expect(cellEditorColors(list(other, band, otherText), cell)).toEqual({ color: '#000000', background: '#ffffff' });
});

test('a mostly opaque fill is composited over the cell beneath it', () => {
  expect(cellEditorColors(list(fill('#000000'), fill('rgba(255, 255, 255, 0.5)'), text('#808080')), cell))
    .toEqual({ color: '#000000', background: '#808080' });
});

import { expect, test } from 'bun:test';
import type { ShapeSnapshot } from '@betteroffice/vsdx';
import { isCellWriteBlocked, isDeleteBlocked, isHandleResizeBlocked } from './commands';

function shapeWith(cells: Array<{ cellName: string; formula: string | null; value: string | null }>): ShapeSnapshot {
  return {
    id: 'shape',
    sourceId: 1,
    name: null,
    children: [],
    cells: cells.map((cell) => ({
      locator: { sheet: 'document' as const, shapeId: 1, section: null, row: null, cellName: cell.cellName },
      name: cell.cellName,
      formula: cell.formula,
      value: cell.value,
    })),
  };
}

function plainShape(): ShapeSnapshot {
  return shapeWith([
    { cellName: 'PinX', formula: '5', value: '5' },
    { cellName: 'PinY', formula: '2', value: '2' },
    { cellName: 'Width', formula: '2', value: '2' },
    { cellName: 'Height', formula: '1', value: '1' },
    { cellName: 'Angle', formula: '0', value: '0' },
  ]);
}

test('a direct GUARD on Angle disables rotation', () => {
  const shape = shapeWith([{ cellName: 'Angle', formula: 'GUARD(0)', value: '0' }]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a GUARD behind one SETATREF hop disables the control', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: 'SETATREF(Vault)', value: '0' },
    { cellName: 'Vault', formula: 'GUARD(0)', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a GUARD at the end of a redirect chain disables the control', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: 'SETATREF(First)', value: '0' },
    { cellName: 'First', formula: 'SETATREF(Second)', value: '0' },
    { cellName: 'Second', formula: 'GUARD(0)', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a redirect cycle disables the control instead of erroring on commit', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: 'SETATREF(Other)', value: '0' },
    { cellName: 'Other', formula: 'SETATREF(Angle)', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a redirect to a missing cell disables the control', () => {
  const shape = shapeWith([{ cellName: 'Angle', formula: 'SETATREF(Nope)', value: '0' }]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a cross-sheet redirect disables the control', () => {
  const shape = shapeWith([{ cellName: 'Angle', formula: 'SETATREF(Sheet.2!Target)', value: '0' }]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a two-argument SETATREF disables the control', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: 'SETATREF(Target,1)', value: '0' },
    { cellName: 'Target', formula: '1', value: '1' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a lowercase setatref redirect is followed', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: 'setatref(vault)', value: '0' },
    { cellName: 'vault', formula: 'GUARD(0)', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('a leading-equals redirect is followed', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: '=SETATREF(Vault)', value: '0' },
    { cellName: 'Vault', formula: 'GUARD(0)', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(true);
});

test('an unguarded redirect target keeps the control enabled', () => {
  const shape = shapeWith([
    { cellName: 'Angle', formula: 'SETATREF(Target)', value: '0' },
    { cellName: 'Target', formula: '0', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(false);
});

test('an unguarded cell stays enabled', () => {
  expect(isCellWriteBlocked(plainShape(), 'Angle')).toBe(false);
});

test('a reference name containing guard or setatref keeps the control enabled', () => {
  const shape = shapeWith([
    { cellName: 'PinX', formula: '5', value: '5' },
    { cellName: 'PinY', formula: '2', value: '2' },
    { cellName: 'Width', formula: 'User.SetatrefWidth', value: '2' },
    { cellName: 'Height', formula: '1', value: '1' },
    { cellName: 'Angle', formula: 'User.GuardAngle', value: '0' },
  ]);
  expect(isCellWriteBlocked(shape, 'Angle')).toBe(false);
  expect(isHandleResizeBlocked(shape)).toBe(false);
});

test('a guarded size redirect disables handle resize', () => {
  const shape = shapeWith([
    { cellName: 'PinX', formula: '5', value: '5' },
    { cellName: 'PinY', formula: '2', value: '2' },
    { cellName: 'Width', formula: 'SETATREF(Locked)', value: '2' },
    { cellName: 'Height', formula: '1', value: '1' },
    { cellName: 'Locked', formula: 'GUARD(2)', value: '2' },
  ]);
  expect(isHandleResizeBlocked(shape)).toBe(true);
});

test('a guarded delete redirect disables delete', () => {
  const shape = shapeWith([
    { cellName: 'LockDelete', formula: 'SETATREF(Vault)', value: '0' },
    { cellName: 'Vault', formula: 'GUARD(0)', value: '0' },
  ]);
  expect(isDeleteBlocked(shape)).toBe(true);
});

test('an unguarded shape keeps resize and delete enabled', () => {
  const shape = plainShape();
  expect(isHandleResizeBlocked(shape)).toBe(false);
  expect(isDeleteBlocked(shape)).toBe(false);
});

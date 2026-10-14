import { expect, test } from 'bun:test';
import { isCellWriteBlocked, isDeleteBlocked, isHandleResizeBlocked, isRotateBlocked } from './commands';
import { initEngineProbe, probeCells as probesFor } from './engineProbe';

await initEngineProbe();

const PLAIN = [
  { cellName: 'PinX', formula: '5' }, { cellName: 'PinY', formula: '2' },
  { cellName: 'Width', formula: '2' }, { cellName: 'Height', formula: '1' },
  { cellName: 'Angle', formula: '0' },
];

test('a direct GUARD on Angle disables rotation', () => {
  expect(isRotateBlocked(probesFor([...PLAIN, { cellName: 'Angle', formula: 'GUARD(0)' }]))).toBe(true);
});

test('a GUARD behind one SETATREF hop disables the control', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'Vault', formula: 'GUARD(0)' }, { cellName: 'Angle', formula: 'SETATREF(Vault)' }]), 'Angle')).toBe(true);
});

test('a GUARD at the end of a redirect chain disables the control', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN,
    { cellName: 'Second', formula: 'GUARD(0)' },
    { cellName: 'First', formula: 'SETATREF(Second)' },
    { cellName: 'Angle', formula: 'SETATREF(First)' },
  ]), 'Angle')).toBe(true);
});

test('a redirect cycle disables the control instead of erroring on commit', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN,
    { cellName: 'Other', formula: 'SETATREF(Angle)' },
    { cellName: 'Angle', formula: 'SETATREF(Other)' },
  ]), 'Angle')).toBe(true);
});

test('a redirect to a missing cell disables the control', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'Angle', formula: 'SETATREF(Nope)' }]), 'Angle')).toBe(true);
});

test('a cross-sheet redirect disables the control', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'Angle', formula: 'SETATREF(Sheet.2!Target)' }]), 'Angle')).toBe(true);
});

test('a two-argument SETATREF disables the control', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'Target', formula: '1' }, { cellName: 'Angle', formula: 'SETATREF(Target,1)' }]), 'Angle')).toBe(true);
});

test('a lowercase setatref redirect is followed', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'vault', formula: 'GUARD(0)' }, { cellName: 'Angle', formula: 'setatref(vault)' }]), 'Angle')).toBe(true);
});

test('a leading-equals redirect is followed', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'Vault', formula: 'GUARD(0)' }, { cellName: 'Angle', formula: '=SETATREF(Vault)' }]), 'Angle')).toBe(true);
});

test('an unguarded redirect target keeps the control enabled', () => {
  expect(isCellWriteBlocked(probesFor([...PLAIN, { cellName: 'Target', formula: '0' }, { cellName: 'Angle', formula: 'SETATREF(Target)' }]), 'Angle')).toBe(false);
});

test('an unguarded cell stays enabled', () => {
  expect(isCellWriteBlocked(probesFor(PLAIN), 'Angle')).toBe(false);
});

test('a reference name containing guard or setatref keeps the control enabled', () => {
  const probes = probesFor([...PLAIN,
    { cellName: 'Width', formula: 'User.SetatrefWidth' },
    { cellName: 'Angle', formula: 'User.GuardAngle' },
  ]);
  expect(isCellWriteBlocked(probes, 'Angle')).toBe(false);
  expect(isHandleResizeBlocked(probes)).toBe(false);
});

test('a guarded size redirect disables handle resize', () => {
  expect(isHandleResizeBlocked(probesFor([...PLAIN,
    { cellName: 'Locked', formula: 'GUARD(2)' },
    { cellName: 'Width', formula: 'SETATREF(Locked)' },
  ]))).toBe(true);
});

test('a guarded delete redirect disables delete', () => {
  expect(isDeleteBlocked(probesFor([...PLAIN, { cellName: 'Vault', formula: 'GUARD(0)' }, { cellName: 'LockDelete', formula: 'SETATREF(Vault)' }]))).toBe(true);
});

test('an unguarded shape keeps resize and delete enabled', () => {
  const probes = probesFor(PLAIN);
  expect(isHandleResizeBlocked(probes)).toBe(false);
  expect(isDeleteBlocked(probes)).toBe(false);
});

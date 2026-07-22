import { describe, expect, test } from 'bun:test';
import { CARET_PAINT_IDLE_MS, PaintedCaretMachine } from './paintedCaret';

describe('painted caret machine', () => {
  test('activates on a painted frame inside the input window', () => {
    const machine = new PaintedCaretMachine();
    expect(machine.shouldPaint(0)).toBe(false);
    machine.noteInput(1000);
    expect(machine.shouldPaint(1000 + CARET_PAINT_IDLE_MS)).toBe(true);
    expect(machine.shouldPaint(1001 + CARET_PAINT_IDLE_MS)).toBe(false);
    expect(machine.framePainted(machine.token())).toBe(true);
    expect(machine.isActive()).toBe(true);
  });

  test('an interrupt invalidates in-flight paints and requests an erase', () => {
    const machine = new PaintedCaretMachine();
    machine.noteInput(0);
    const token = machine.token();
    expect(machine.framePainted(token)).toBe(true);
    expect(machine.interrupt()).toBe(true);
    expect(machine.isActive()).toBe(false);
    expect(machine.framePainted(token)).toBe(false);
    expect(machine.shouldPaint(1)).toBe(false);
    expect(machine.interrupt()).toBe(false);
  });

  test('idle timeout ends active mode once the threshold elapses', () => {
    const machine = new PaintedCaretMachine();
    machine.noteInput(0);
    machine.framePainted(machine.token());
    expect(machine.idleTimeout(CARET_PAINT_IDLE_MS - 1)).toBe(false);
    machine.noteInput(200);
    expect(machine.idleTimeout(CARET_PAINT_IDLE_MS)).toBe(false);
    expect(machine.msUntilIdle(CARET_PAINT_IDLE_MS)).toBe(200);
    expect(machine.idleTimeout(200 + CARET_PAINT_IDLE_MS)).toBe(true);
    expect(machine.isActive()).toBe(false);
    expect(machine.idleTimeout(201 + CARET_PAINT_IDLE_MS)).toBe(false);
  });

  test('an unpainted frame drops active mode without an erase request', () => {
    const machine = new PaintedCaretMachine();
    machine.noteInput(0);
    machine.framePainted(machine.token());
    expect(machine.frameUnpainted()).toBe(true);
    expect(machine.frameUnpainted()).toBe(false);
  });

  test('a fresh token after an interrupt activates again', () => {
    const machine = new PaintedCaretMachine();
    machine.noteInput(0);
    machine.interrupt();
    machine.noteInput(10);
    expect(machine.shouldPaint(20)).toBe(true);
    expect(machine.framePainted(machine.token())).toBe(true);
    expect(machine.isActive()).toBe(true);
  });

  test('a dispatch hold hides the caret until its frame resolves it', () => {
    const machine = new PaintedCaretMachine();
    expect(machine.isHolding(0)).toBe(false);
    machine.noteDispatch(100);
    expect(machine.isHolding(100)).toBe(true);
    expect(machine.isHolding(99 + CARET_PAINT_IDLE_MS)).toBe(true);
    expect(machine.isHolding(100 + CARET_PAINT_IDLE_MS)).toBe(false);
    expect(machine.shouldPaint(100)).toBe(true);
    expect(machine.framePainted(machine.token())).toBe(true);
    expect(machine.isHolding(150)).toBe(false);
    expect(machine.isActive()).toBe(true);
  });

  test('unpainted frames and interrupts drop a pending dispatch hold', () => {
    const machine = new PaintedCaretMachine();
    machine.noteDispatch(0);
    expect(machine.frameUnpainted()).toBe(false);
    expect(machine.isHolding(1)).toBe(false);
    machine.noteDispatch(10);
    expect(machine.interrupt()).toBe(false);
    expect(machine.isHolding(11)).toBe(false);
  });
});

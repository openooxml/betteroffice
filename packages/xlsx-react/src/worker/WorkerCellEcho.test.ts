import type { DisplayList, WorkbookCommandFrame } from '@betteroffice/xlsx';
import { describe, expect, test } from 'bun:test';
import { WorkerCellEcho } from './WorkerCellEcho';
import { PeerlessWorkerPaintSource as WorkerPaintSource, type WorkerPaintRequest } from './PeerlessWorkerPaintSource';

function harness(maxPatches = 8) {
  const request: WorkerPaintRequest = {
    generation: 1, sheet: 0, navigation: 0, width: 192, height: 48, zoom: 1, dpr: 1,
    viewport: { x: 0, y: 0, width: 192, height: 48 },
  };
  const text: string[] = [];
  const clips: number[][] = [];
  const context = {
    save() {}, restore() {}, setTransform() {}, beginPath() {}, clip() {}, clearRect() {},
    rect(...rect: number[]) { clips.push(rect); },
    fillText(value: string) { text.push(value); },
    fillRect() {},
  } as unknown as CanvasRenderingContext2D;
  const errors: unknown[] = [];
  const source = new WorkerPaintSource({
    generation: 1, commandIdentity: { session: 'book', incarnation: 'worker' },
    capture: () => request, sentSequence: () => 0,
    requestFrame: () => { throw new Error('Unexpected frame request'); },
    publish: (painted) => { text.push(`frame:${painted.version}`); },
    onError: (error) => { errors.push(error); },
  });
  const echo = new WorkerCellEcho({ source, context: () => context, maxPatches });
  const install = (coversThroughSeq: number, value = 'final', startRow = 0) => {
    const captured = source.captureCommandFrame()!;
    const displayList: DisplayList = {
      width: 192, height: 48,
      commands: [{ op: 'fillRect', x: 0, y: 0, w: 192, h: 48, color: '#ffffff' },
        { op: 'text', text: value, x: 8, y: 18, fontSize: 11, color: '#123456', bold: true,
          clip: { x: 0, y: 0, w: 96, h: 24 } }],
      grid: { startRow, startCol: 0, rowOffsets: [0, 24, 48], colOffsets: [0, 96, 192, 288] },
    };
    const frame: WorkbookCommandFrame = {
      session: 'book', incarnation: 'worker', coversThroughSeq, sequence: 90, epoch: 1,
      sheet: request.sheet, viewport: request.viewport, displayList, version: value,
      requestGeneration: captured.requestGeneration!, zoom: 1, dpr: 1,
    };
    return source.install(frame, captured);
  };
  expect(install(0, 'initial')).toBe(true);
  text.length = 0;
  clips.length = 0;
  return { source, echo, text, clips, errors, context, install, request };
}

describe('WorkerCellEcho', () => {
  test('paste inspects visible entries after an offscreen prefix within its budget', () => {
    const { echo, request, install, text } = harness(2);
    request.viewport.y = 999 * 24;
    expect(install(0, 'scrolled', 999)).toBe(true);
    text.length = 0;
    const edits = Array.from({ length: 1000 }, (_, row) => ({ row, col: 0, input: `row ${row + 1}` }));
    let inspected = 0;
    echo.paste(1, 0, new Proxy(edits, { get(target, property, receiver) {
      if (typeof property === 'string' && /^\d+$/.test(property)) inspected++;
      return Reflect.get(target, property, receiver);
    } }));
    expect(echo.size).toBe(1);
    expect(text).toContain('row 1000');
    expect(inspected).toBeLessThan(100);
  });

  test('ack retains echo until covering paint', () => {
    const { echo, source, text, install } = harness();
    const authoritative = source.painted;
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: '=1+2' }]);
    expect(text).toEqual(['=1+2']);
    expect(source.painted).toBe(authoritative);
    echo.acknowledge({ throughSeq: 1, statuses: [{ seq: 1, status: 'applied' }] });
    expect(echo.size).toBe(1);
    expect(text).toEqual(['=1+2']);
    text.length = 0;
    expect(install(1, '3')).toBe(true);
    expect(echo.size).toBe(0);
    expect(text).toEqual(['frame:3']);
    expect(source.painted?.displayList.commands[1]).toMatchObject({ text: '3' });
  });

  test('reuses cached appearance without changing authoritative geometry', () => {
    const { echo, source, context, clips } = harness();
    const authoritative = source.painted!;
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'raw' }]);
    expect(context.font).toContain('bold');
    expect(context.fillStyle).toBe('#123456');
    expect(clips[0]).toEqual([0, 0, 96, 24]);
    expect(source.painted).toBe(authoritative);
    expect(source.painted?.geometry).toBe(authoritative.displayList.grid);
    expect(authoritative.displayList.commands[1]).toMatchObject({ text: 'initial' });
  });

  test('predicted navigation uses cached visible geometry', () => {
    const { echo, request, text } = harness();
    request.navigation++;
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'committed' }]);
    expect(echo.size).toBe(1);
    expect(text).toEqual(['committed']);
  });

  test('older coverage cannot retire newer patch', () => {
    const { echo, text, install } = harness();
    echo.patch(2, [{ sheet: 0, row: 0, col: 0, text: 'newer' }]);
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'older' }]);
    expect(text[text.length - 1]).toBe('newer');
    text.length = 0;
    expect(install(1, 'accepted older')).toBe(true);
    expect(echo.size).toBe(1);
    expect(text).toEqual(['frame:accepted older', 'newer']);
    expect(install(0)).toBe(false);
    expect(echo.size).toBe(1);
    echo.refuse(2);
    expect(text[text.length - 1]).toBe('accepted older');
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'late' }]);
    expect(echo.size).toBe(0);
  });

  test('refusal repaints damaged cells', () => {
    const { echo, text, clips } = harness();
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'first' }]);
    echo.patch(2, [{ sheet: 0, row: 0, col: 0, text: 'second' }]);
    text.length = 0;
    clips.length = 0;
    echo.acknowledge({ throughSeq: 2, statuses: [{ seq: 2, status: 'refused', committed: false }] });
    expect(text).toEqual(['first']);
    expect(clips[0]).toEqual([0, 0, 96, 24]);
    expect(echo.size).toBe(1);
    text.length = 0;
    echo.refuse(1);
    expect(text).toEqual(['initial']);
    expect(echo.size).toBe(0);
  });

  test('large paste stays bounded', () => {
    const { echo, text } = harness(2);
    const cells = Array.from({ length: 10_000 }, (_, row) => ({ sheet: 0, row, col: 0, text: 'paste' }));
    echo.patch(1, cells);
    expect(echo.size).toBe(2);
    echo.patch(2, [{ sheet: 1, row: 0, col: 0, text: 'other sheet' },
      { sheet: 0, row: 0, col: 2, text: 'offscreen' },
      { sheet: 0, row: 0, col: 0, text: 'x'.repeat(5000) }]);
    expect(echo.size).toBe(2);
    echo.clear(3, { sheet: 0, top: 0, left: 0, bottom: 1_000_000, right: 1_000_000 });
    expect(echo.size).toBe(2);
    expect(text).not.toContain('offscreen');
    expect(text).not.toContain('other sheet');
    expect(text).not.toContain('x'.repeat(5000));
  });

  test('failed covering paint retains pending echo', () => {
    const { echo, source, install, errors } = harness();
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'pending' }]);
    echo.patch(2, [{ sheet: 0, row: 0, col: 0, text: 'later' }]);
    const authoritative = source.painted;
    const failure = new Error('Canvas missing');
    const unregister = source.registerPaintTransaction(() => { throw failure; });
    expect(install(1)).toBe(false);
    expect(echo.size).toBe(2);
    expect(source.painted).toBe(authoritative);
    expect(errors).toEqual([failure]);
    unregister();
  });

  test('large mostly invisible paste bounds inspected cells', () => {
    const { echo } = harness(8);
    let inspected = 0;
    const cells = (function* () {
      for (let row = 0; row < 10_000; row++) {
        inspected++;
        yield { sheet: 0, row, col: 0, text: 'paste' };
      }
    })();
    echo.patch(1, cells);
    expect(inspected).toBe(8);
    expect(echo.size).toBe(2);
  });

  test('viewport changes discard offscreen echo', () => {
    const { echo, request, install } = harness();
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'pending' }]);
    request.viewport.y = 2400;
    expect(install(0, 'scrolled', 100)).toBe(true);
    expect(echo.size).toBe(0);
    echo.patch(2, [{ sheet: 0, row: 100, col: 0, text: 'scrolled pending' }]);
    expect(echo.size).toBe(1);
    request.sheet = 1;
    expect(install(0)).toBe(true);
    expect(echo.size).toBe(0);
  });

  test('echo redraw failure rolls back covering paint and coverage', () => {
    const { echo, source, context, install, errors } = harness();
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'pending' }]);
    echo.patch(2, [{ sheet: 0, row: 0, col: 0, text: 'later' }]);
    const authoritative = source.painted;
    const failure = new Error('Echo paint failed');
    context.fillText = () => { throw failure; };
    expect(install(1)).toBe(false);
    expect(echo.size).toBe(2);
    expect(source.painted).toBe(authoritative);
    expect(source.coversThroughSeq).toBe(0);
    expect(errors).toEqual([failure]);
  });

  test('dispose restores authoritative pixels and removes the transaction', () => {
    const { echo, text, install } = harness();
    echo.patch(1, [{ sheet: 0, row: 0, col: 0, text: 'pending' }]);
    text.length = 0;
    echo.dispose();
    expect(text).toEqual(['initial']);
    expect(echo.size).toBe(0);
    text.length = 0;
    expect(install(1)).toBe(true);
    expect(text).toEqual(['frame:final']);
    echo.patch(2, [{ sheet: 0, row: 0, col: 0, text: 'late' }]);
    expect(echo.size).toBe(0);
  });
});

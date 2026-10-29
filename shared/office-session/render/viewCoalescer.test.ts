import { expect, test } from 'bun:test';
import { createViewCoalescer } from './viewCoalescer';
import type { ViewMessage } from './types';

test('view updates merge into one send per callback with increasing generations', () => {
  const turns = new Map<number, () => void>(), sent: ViewMessage[] = [];
  let handle = 0;
  const coalescer = createViewCoalescer({
    requestFrame(callback) { const id = handle++; turns.set(id, callback); return id; },
    cancelFrame(id: number) { turns.delete(id); },
    send(view) { sent.push(view); },
  });
  const surfaces: ViewMessage['surfaces'] = [{ key: { kind: 'pane', id: 'body' },
    visibleRect: { x: 0, y: 0, width: 100, height: 100 }, priority: 0 }];
  for (let index = 0; index < 30; index += 1) coalescer.update({ scale: index + 1 });
  coalescer.update({ surfaces }); coalescer.update({ overscan: 12 });
  expect(turns.size).toBe(1);
  expect(sent).toEqual([]);
  turns.get(0)!(); turns.delete(0);
  expect(sent).toEqual([{ generation: 1, scale: 30, surfaces, overscan: 12 }]);
  expect(turns.size).toBe(0);
  coalescer.update({ scale: 2 });
  turns.get(1)!(); turns.delete(1);
  expect(sent[1]).toEqual({ generation: 2, scale: 2, surfaces, overscan: 12 });
  coalescer.update({ overscan: 0 }); coalescer.dispose(); coalescer.update({ scale: 5 });
  expect(turns.size).toBe(0);
  expect(sent.length).toBe(2);
});

import { afterEach, beforeEach, expect, spyOn, test } from 'bun:test';
import { attachPptxWorkerRender, pptxRasterOwnership, type PptxRenderReply, type PptxRenderRequest } from './pptxWorkerRender';
import type { SessionTransport } from '../transport';

class Port {
  onmessage: ((event: { data: PptxRenderReply }) => void) | null = null;
  onmessageerror: (() => void) | null = null;
  sent: PptxRenderRequest[] = [];
  closed = false;
  postMessage(message: PptxRenderRequest) { this.sent.push(message); }
  addEventListener() {}
  close() { this.closed = true; }
}
let original: PropertyDescriptor | undefined;
let originalOptions: PropertyDescriptor | undefined;
let channels: { port1: Port; port2: Port }[];
let disposers: (() => void)[];
beforeEach(() => {
  original = Object.getOwnPropertyDescriptor(globalThis, 'MessageChannel');
  originalOptions = Object.getOwnPropertyDescriptor(globalThis, '__betterofficePptxWorkerRender');
  Reflect.deleteProperty(globalThis, '__betterofficePptxWorkerRender');
  channels = []; disposers = [];
  Object.defineProperty(globalThis, 'MessageChannel', { configurable: true, value: class {
    port1 = new Port(); port2 = new Port();
    constructor() { channels.push(this); }
  } });
});
afterEach(() => {
  for (const dispose of disposers) dispose();
  if (original) Object.defineProperty(globalThis, 'MessageChannel', original);
  else Reflect.deleteProperty(globalThis, 'MessageChannel');
  if (originalOptions) Object.defineProperty(globalThis, '__betterofficePptxWorkerRender', originalOptions);
  else Reflect.deleteProperty(globalThis, '__betterofficePptxWorkerRender');
});
function host(failAttach = false) {
  const posted: unknown[] = [], errors = new Set<(error: unknown) => void>();
  const transport: SessionTransport = { post(message) { if (failAttach) throw new Error('unavailable'); posted.push(message); },
    listen: () => () => {}, onError(listener) { errors.add(listener); return () => { errors.delete(listener); }; }, close() {},
  };
  const client = attachPptxWorkerRender(transport);
  if (client) disposers.push(() => client.dispose());
  return { client, posted, errors, port: channels[channels.length - 1]!.port1 };
}
function raster() {
  const bitmap = { width: 2, height: 2, closed: 0, close() { this.closed += 1; } };
  const reply: PptxRenderReply = { kind: 'raster', slideId: 'slide', frameEpoch: 1, baseEpoch: null, content: 1, baseContent: null, factor: 1,
    width: 2, height: 2, rect: { x: 0, y: 0, width: 2, height: 2 }, bitmap: bitmap as unknown as ImageBitmap };
  return { reply, bitmap };
}

test('the private port sends only view, fonts, resync and dispose; font buffers are copied', () => {
  const h = host(); expect(h.client).toBeDefined();
  const source = Uint8Array.of(1, 2).buffer;
  h.client!.fonts([{ family: 'Host', source, descriptors: { weight: '600' } }]);
  h.client!.view(0.5, 2); h.client!.resync();
  const first = h.port.sent[0]!;
  expect(first.kind).toBe('fonts');
  if (first.kind !== 'fonts') throw new Error('missing faces');
  expect(first.faces[0]!.source).not.toBe(source);
  expect(first.faces[0]!.source).toEqual(source);
  h.client!.dispose();
  expect(h.port.sent.map((message) => message.kind)).toEqual(['fonts', 'view', 'resync', 'dispose']);
});

test('unclaimed, asynchronously retained and late rasters each close exactly once', () => {
  const h = host(), unclaimed = raster();
  h.port.onmessage?.({ data: unclaimed.reply }); expect(unclaimed.bitmap.closed).toBe(1);
  const held = raster(), holder = {};
  const off = h.client!.subscribe((reply) => { if (reply.kind === 'raster') pptxRasterOwnership.retain(reply.bitmap, holder); });
  h.port.onmessage?.({ data: held.reply }); expect(held.bitmap.closed).toBe(0);
  if (held.reply.kind !== 'raster') throw new Error('missing raster');
  pptxRasterOwnership.release(held.reply.bitmap, holder);
  expect(held.bitmap.closed).toBe(1); off();
  h.client!.dispose(); const late = raster(); h.port.onmessage?.({ data: late.reply });
  expect(late.bitmap.closed).toBe(1);
});

test.each(['messageerror', 'transport', 'unsupported'])('%s failure persists for new subscribers', (failure) => {
  const h = host(), received: PptxRenderReply[] = [];
  h.client!.subscribe((reply) => received.push(reply));
  if (failure === 'messageerror') h.port.onmessageerror?.();
  else if (failure === 'transport') for (const listener of h.errors) listener(new Error('crashed'));
  else h.port.onmessage?.({ data: { kind: 'unsupported' } });
  expect(h.client!.supported).toBe(false);
  h.client!.subscribe((reply) => received.push(reply));
  expect(received).toEqual([{ kind: 'unsupported' }, { kind: 'unsupported' }]);
});

test('port attachment failure leaves the main session available', () => {
  const h = host(true); expect(h.client).toBeUndefined(); expect(channels[0]!.port1.closed).toBe(true);
});

test('the internal kill switch skips channel creation and render port attachment', () => {
  Object.defineProperty(globalThis, '__betterofficePptxWorkerRender', {
    configurable: true, value: { options: { enabled: false } },
  });
  const posted: unknown[] = [];
  const transport: SessionTransport = { post: (message) => { posted.push(message); },
    listen: () => () => {}, onError: () => () => {}, close() {},
  };
  expect(attachPptxWorkerRender(transport)).toBeUndefined();
  expect(channels).toEqual([]); expect(posted).toEqual([]);
});

test.each([{}, { options: {} }, { options: { enabled: true } }])('worker rendering stays enabled with %j', (value) => {
  Object.defineProperty(globalThis, '__betterofficePptxWorkerRender', { configurable: true, value });
  const h = host();
  expect(h.client?.supported).toBe(true); expect(h.posted).toHaveLength(1);
});

test('replies attached to session messages reach listeners first and leave the message', () => {
  const listeners = new Set<(message: unknown) => void>();
  const transport: SessionTransport = { post() {}, onError: () => () => {}, close() {},
    listen(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; } };
  const client = attachPptxWorkerRender(transport)!; disposers.push(() => client.dispose());
  const seen: string[] = [];
  client.subscribe((reply) => seen.push(reply.kind));
  const order: string[] = [];
  transport.listen((message) => order.push(`${seen.length}:${'pptxRender' in (message as object)}`));
  const { reply, bitmap } = raster();
  const message = { protocol: 1, kind: 'reply', id: 4, pptxRender: reply };
  for (const listener of [...listeners]) listener(message);
  expect(seen).toEqual(['raster']); expect(order).toEqual(['1:false']); expect(bitmap.closed).toBe(1);
  client.dispose();
  const late = raster(); for (const listener of [...listeners]) listener({ protocol: 1, kind: 'reply', id: 5, pptxRender: late.reply });
  expect(late.bitmap.closed).toBe(0);
});

test('an unchanged view is posted once', () => {
  const h = host();
  h.client!.view(1, 2); h.client!.view(1, 2); h.client!.view(1, 1);
  expect(h.port.sent.filter((message) => message.kind === 'view')).toEqual([
    { kind: 'view', scale: 1, dpr: 2 }, { kind: 'view', scale: 1, dpr: 1 }]);
});

test('prerender replies and font notices leave the unresponsive-worker watchdog armed', () => {
  const timers: (() => void)[] = [];
  let cleared = 0;
  const set = spyOn(globalThis, 'setTimeout').mockImplementation(((callback: () => void) => {
    timers.push(callback); return timers.length;
  }) as unknown as typeof setTimeout);
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation((() => { cleared += 1; }) as typeof clearTimeout);
  try {
    const h = host();
    const replies: PptxRenderReply[] = [];
    h.client!.subscribe((reply) => { replies.push(reply); });
    h.client!.view(1, 2);
    expect(timers).toHaveLength(1);
    const bitmap = { closed: 0, close() { this.closed += 1; } };
    h.port.onmessage?.({ data: { kind: 'prerender', slideId: 'slide', token: 9, version: '1', fonts: 0, factor: 2,
      width: 2, height: 2, bitmap: bitmap as unknown as ImageBitmap } });
    h.port.onmessage?.({ data: { kind: 'font-generation', fonts: 1 } });
    expect(cleared).toBe(0);
    expect(bitmap.closed).toBe(1);
    timers[0]!();
    expect(h.client!.supported).toBe(false);
    expect(replies.at(-1)).toEqual({ kind: 'unsupported' });
  } finally { set.mockRestore(); clear.mockRestore(); }
});


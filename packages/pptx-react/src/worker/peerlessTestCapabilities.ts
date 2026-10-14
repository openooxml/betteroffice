import { afterEach, beforeEach } from 'bun:test';

let original: PropertyDescriptor | undefined;
beforeEach(() => {
  original = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
  Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, value: class {} });
});
afterEach(() => {
  if (original) Object.defineProperty(globalThis, 'OffscreenCanvas', original);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});

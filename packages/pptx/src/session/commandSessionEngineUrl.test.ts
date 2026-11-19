import { afterEach, describe, expect, test } from 'bun:test';
import { engineFetchUrl } from './commandSession';

const original = Object.getOwnPropertyDescriptor(globalThis, 'location');

afterEach(() => {
  if (original) Object.defineProperty(globalThis, 'location', original);
  else delete (globalThis as { location?: unknown }).location;
});

describe('engineFetchUrl', () => {
  test('resolves a served-path href, as Turbopack names assets, against the page', () => {
    Object.defineProperty(globalThis, 'location', {
      configurable: true,
      value: { href: 'http://localhost:3000/ko/room/1' },
    });
    const url = engineFetchUrl('/_next/static/media/pptx_wasm_bg.0abc.wasm');
    expect(url.href).toBe('http://localhost:3000/_next/static/media/pptx_wasm_bg.0abc.wasm');
    expect(url.protocol).toBe('http:');
  });

  test('keeps an absolute href, with or without a page location', () => {
    expect(engineFetchUrl('https://cdn.example/pptx_wasm_bg.wasm').href).toBe('https://cdn.example/pptx_wasm_bg.wasm');
    delete (globalThis as { location?: unknown }).location;
    expect(engineFetchUrl('file:///app/dist/generated/pptx_wasm_bg.wasm').href)
      .toBe('file:///app/dist/generated/pptx_wasm_bg.wasm');
  });
});

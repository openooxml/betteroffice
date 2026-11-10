import { expect, test } from 'bun:test';

type Releasable = { free(): void };

/** One engine class: how to make a live instance and a call that needs it alive. */
export interface DisposeCase<T extends Releasable = Releasable> {
  name: string;
  engineClass: { prototype: Releasable };
  make(): T;
  use(handle: T): unknown;
}

/**
 * A second `free()` or `[Symbol.dispose]()` throws on wasm (the glue has
 * already cleared the pointer) and does nothing on native.
 */
export const SECOND_RELEASE = { wasm: 'throws', native: 'no-op' } as const;

/** Every exported class of an engine module that owns engine memory. */
export function releasableClasses(engine: Record<string, unknown>): string[] {
  return Object.entries(engine)
    .filter(([, value]) => typeof value === 'function' && typeof (value as { prototype?: Releasable }).prototype?.free === 'function')
    .map(([name]) => name)
    .sort();
}

function release(handle: Releasable, how: 'free' | 'dispose') {
  if (how === 'free') handle.free();
  else (handle as Releasable & Disposable)[Symbol.dispose]();
}

/** Registers the shared release outcome table for one engine module. */
export function describeDisposeTable(engineKind: 'wasm' | 'native', engine: Record<string, unknown>, cases: DisposeCase[]) {
  test('every engine class with free() is in the table', () => {
    expect(cases.map((entry) => entry.name).sort()).toEqual(releasableClasses(engine));
  });

  for (const entry of cases) {
    test(`${entry.name}: [Symbol.dispose] is free, as the wasm glue assigns it`, () => {
      const prototype = entry.engineClass.prototype as Releasable & Record<symbol, unknown>;
      expect(typeof prototype[Symbol.dispose]).toBe('function');
      expect(prototype[Symbol.dispose]).toBe(prototype.free);
      expect(Object.getOwnPropertyDescriptor(prototype, Symbol.dispose)).toMatchObject({
        enumerable: true, writable: true, configurable: true,
      });
    });

    test(`${entry.name}: using releases at scope exit`, () => {
      let held: Releasable | undefined;
      {
        using handle = entry.make() as Releasable & Disposable;
        held = handle;
        expect(() => entry.use(handle)).not.toThrow();
      }
      expect(() => entry.use(held!)).toThrow(Error);
    });

    for (const how of ['free', 'dispose'] as const) {
      test(`${entry.name}: a call after ${how} throws an Error`, () => {
        const handle = entry.make();
        expect(() => entry.use(handle)).not.toThrow();
        release(handle, how);
        expect(() => entry.use(handle)).toThrow(Error);
      });
    }

    for (const [first, second] of [['free', 'free'], ['dispose', 'dispose'], ['free', 'dispose'], ['dispose', 'free']] as const) {
      test(`${entry.name}: ${second} after ${first} on ${engineKind}: ${SECOND_RELEASE[engineKind]}`, () => {
        const handle = entry.make();
        release(handle, first);
        if (SECOND_RELEASE[engineKind] === 'throws') expect(() => release(handle, second)).toThrow(Error);
        else expect(() => release(handle, second)).not.toThrow();
        expect(() => entry.use(handle)).toThrow(Error);
      });
    }
  }
}

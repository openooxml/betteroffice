import { expect, test } from 'bun:test';
import { stableStringify, stableStringifyWithout } from './yrsToDocument';

type Attrs = Record<string, unknown>;

function asObject(value: unknown): Attrs | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Attrs)
    : undefined;
}

function referenceStableStringify(value: unknown): string {
  if (value === null || value === undefined) return 'null';
  if (Array.isArray(value)) return `[${value.map(referenceStableStringify).join(',')}]`;
  const object = asObject(value);
  if (object) {
    return `{${Object.keys(object)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${referenceStableStringify(object[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

function randomValues(): () => unknown {
  let state = 0x51c0ffee;
  const random = (limit: number): number => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return (state >>> 0) % limit;
  };
  const strings = [
    '', 'plain', 'quote"', 'back\\slash', '\n\r\t\b\f\u0000\u001f',
    'caf\u00e9\u4e2d\ud83d\ude00', '\u2028\u2029', '\ud800', '\udfff',
  ];
  const keys = [
    '', 'simple', 'A1_$', '0', '10', '2', 'key.with.dots', 'key[0]', 'quote"',
    'back\\slash', '\n\r\t\u0000', 'trailing\n', '\u4e2d', '\ud800',
    'hyperlink', 'ins', 'del',
  ];
  const primitives: unknown[] = [
    null, undefined, NaN, Infinity, -Infinity, -0, 0, true, false,
    Symbol('value'), () => undefined,
  ];
  const value = (depth: number): unknown => {
    const kind = random(depth < 5 ? 6 : 3);
    if (kind === 0) return primitives[random(primitives.length)];
    if (kind === 1) return strings[random(strings.length)];
    if (kind === 2) return (random(200001) - 100000) / 100;
    if (kind === 3) {
      const array = new Array<unknown>(random(6));
      for (let index = 0; index < array.length; index += 1) {
        if (random(4) !== 0) array[index] = value(depth + 1);
      }
      return array;
    }
    const shuffled = [...keys];
    for (let index = shuffled.length - 1; index > 0; index -= 1) {
      const other = random(index + 1);
      [shuffled[index], shuffled[other]] = [shuffled[other]!, shuffled[index]!];
    }
    const object: Attrs = {};
    const count = random(6);
    for (let index = 0; index < count; index += 1) {
      object[shuffled[index]!] = value(depth + 1);
    }
    return object;
  };
  return () => value(0);
}

test('stableStringify matches the oracle for deterministic nested values', () => {
  const next = randomValues();
  for (let index = 0; index < 2000; index += 1) {
    const value = next();
    expect(stableStringify(value)).toBe(referenceStableStringify(value));
  }
});

test('stableStringify preserves sparse arrays and unsupported JSON values', () => {
  const inherited = new Array<unknown>(4);
  Object.setPrototypeOf(inherited, Object.assign(Object.create(Array.prototype), { 1: undefined }));
  const values = [
    undefined, null, NaN, Infinity, -Infinity, -0, Symbol('top'), () => undefined,
    new Array<unknown>(3), [undefined], [() => undefined, Symbol('nested')],
    [null, , undefined, , NaN, -0], inherited,
    { fn: () => undefined, symbol: Symbol('nested'), missing: undefined },
    { z: { b: 1, a: 2 }, a: ['\ud800', '\udc00', '\ud83d\ude00'] },
    { '': 0, 'trailing\n': 1, 'quote"': 2, 'back\\slash': 3, '\u0000': 4 },
    new Date(0), new Map([['a', 1]]), new Set([1]), /text/,
  ];
  for (const value of values) {
    expect(stableStringify(value)).toBe(referenceStableStringify(value));
  }
  expect(stableStringify(undefined)).toBe('null');
  expect(stableStringify([undefined])).toBe('[null]');
  expect(stableStringify(new Array<unknown>(3))).toBe('[,,]');
  expect(stableStringify([() => undefined, Symbol('nested')])).toBe('[,]');
  expect(stableStringify({ value: Symbol('nested') })).toBe('{"value":undefined}');
  expect(stableStringify(Symbol('top'))).toBeUndefined();
});

test('stableStringify preserves JSON errors', () => {
  const errorFrom = (stringify: (value: unknown) => string, value: unknown): Error | undefined => {
    try {
      stringify(value);
    } catch (error) {
      if (!(error instanceof Error)) throw error;
      return error;
    }
    return undefined;
  };
  for (const value of [1n, [1n], { value: 1n }]) {
    const expected = errorFrom(referenceStableStringify, value);
    const actual = errorFrom(stableStringify, value);
    expect(expected).toBeInstanceOf(TypeError);
    expect(actual?.constructor).toBe(expected?.constructor);
    expect(actual?.message).toBe(expected?.message);
  }
});

function referenceWithout(value: Attrs, omitted: ReadonlySet<string>): string {
  const copy = { ...value };
  for (const key of omitted) delete copy[key];
  return referenceStableStringify(copy);
}

test('stableStringifyWithout matches a copied object with omitted top-level keys', () => {
  const next = randomValues();
  const omissions = [
    new Set<string>(), new Set(['hyperlink', 'ins', 'del']),
    new Set(['simple', '', 'quote"', 'absent']),
  ];
  for (let index = 0; index < 1000; index += 1) {
    const value: Attrs = {
      z: next(), hyperlink: next(), ins: next(), del: next(),
      nested: { hyperlink: next(), ins: next(), del: next() },
      'quote"': next(), '': next(), simple: next(),
    };
    for (const omitted of omissions) {
      expect(stableStringifyWithout(value, omitted)).toBe(referenceWithout(value, omitted));
    }
  }
});

test('stableStringifyWithout only includes own enumerable string keys', () => {
  const value: Attrs = Object.create({ inherited: true });
  value.z = undefined;
  value.ins = { nested: true };
  value.a = { ins: 'kept', del: null, hyperlink: 'kept' };
  Object.defineProperty(value, 'hidden', { value: 1 });
  Object.defineProperty(value, Symbol('hidden'), { value: 2, enumerable: true });
  const omitted = new Set(['ins', 'del', 'hyperlink', 'inherited']);
  const descriptors = Object.getOwnPropertyDescriptors(value);
  expect(stableStringifyWithout(value, omitted)).toBe(referenceWithout(value, omitted));
  expect(Object.getOwnPropertyDescriptors(value)).toEqual(descriptors);
  expect(stableStringifyWithout(value, new Set(Object.keys(value)))).toBe('{}');
  expect(stableStringifyWithout({}, omitted)).toBe('{}');
});

test('serialization observes mutations between calls', () => {
  const value: Attrs = { a: { value: 1 }, hyperlink: 'link' };
  const omitted = new Set(['hyperlink']);
  expect(stableStringifyWithout(value, omitted)).toBe(referenceWithout(value, omitted));
  (value.a as Attrs).value = 2;
  omitted.add('a');
  expect(stableStringify(value)).toBe(referenceStableStringify(value));
  expect(stableStringifyWithout(value, omitted)).toBe(referenceWithout(value, omitted));
});

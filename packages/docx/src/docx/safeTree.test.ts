import { expect, test } from 'bun:test';
import { assertSafeTree } from './rustParseFacade';
import { assertSafeSaveTree } from './rustSaveFacade';

const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

function referenceAssertSafeSaveTree(value: unknown, path: string, active = new WeakSet<object>()): void {
  if (value === null || typeof value !== 'object') return;
  if (active.has(value)) throw new TypeError(`${path} contains a cycle`);
  active.add(value);
  if (Array.isArray(value)) {
    value.forEach((entry, index) => referenceAssertSafeSaveTree(entry, `${path}[${index}]`, active));
  } else {
    for (const key of Object.keys(value)) {
      if (FORBIDDEN_KEYS.has(key)) throw new TypeError(`${path} contains forbidden key ${key}`);
      referenceAssertSafeSaveTree((value as Record<string, unknown>)[key], `${path}.${key}`, active);
    }
  }
  active.delete(value);
}

function referenceAssertSafeTree(value: unknown, path: string, active = new WeakSet<object>()): void {
  if (value === null || typeof value !== 'object') return;
  if (active.has(value)) throw new TypeError(`${path} contains a cycle`);
  active.add(value);
  try {
    for (const key of Object.keys(value)) {
      if (FORBIDDEN_KEYS.has(key)) throw new TypeError(`${path} contains forbidden key ${key}`);
      referenceAssertSafeTree((value as Record<string, unknown>)[key], `${path}.${key}`, active);
    }
  } finally {
    active.delete(value);
  }
}

type Validator = (value: unknown, path: string, active?: WeakSet<object>) => void;

function outcome(validate: Validator, value: unknown, path: string, active?: WeakSet<object>): unknown {
  try {
    return validate(value, path, active);
  } catch (error) {
    if (!(error instanceof Error)) throw error;
    return { type: error.constructor, message: error.message };
  }
}

const VALIDATORS = [
  { name: 'save', validate: assertSafeSaveTree, reference: referenceAssertSafeSaveTree },
  { name: 'parse', validate: assertSafeTree, reference: referenceAssertSafeTree },
];

for (const { name, validate, reference } of VALIDATORS) {
  test(`${name} tree validation preserves forbidden-key errors and visit order`, () => {
    for (const key of FORBIDDEN_KEYS) {
      const values = [
        { [key]: true },
        { 'nested.key[2]': { '': { [key]: true } } },
        { 'nested.key[2]': [{ 'child[0].key': { [key]: true } }] },
        [[{ [key]: true }]],
        { first: { [key]: true }, [key]: false },
      ];
      for (const value of values) {
        for (const path of ['wire', 'save', 'custom.root[2]', '']) {
          expect(outcome(validate, value, path)).toEqual(outcome(reference, value, path));
        }
      }
    }
  });

  test(`${name} tree validation preserves cycle errors`, () => {
    const self: Record<string, unknown> = {};
    self['self.key[0]'] = self;
    const child: Record<string, unknown> = {};
    const nested = { 'outer.key[1]': [child] };
    child['back[2].key'] = nested;
    const array: unknown[] = [];
    array.push(array);
    const first: Record<string, unknown> = {};
    first.back = first;
    const beforeForbidden = { first, constructor: true };
    for (const value of [self, nested, array, beforeForbidden]) {
      for (const path of ['wire', 'save', 'custom.root[2]', '']) {
        expect(outcome(validate, value, path)).toEqual(outcome(reference, value, path));
      }
    }
  });

  test(`${name} tree validation preserves array and enumerable-key handling`, () => {
    const values: unknown[] = [];
    for (const key of FORBIDDEN_KEYS) {
      const array = [{ text: 'valid' }];
      Object.defineProperty(array, key, { value: true, enumerable: true });
      values.push(array);
    }
    const sparse = new Array<unknown>(4);
    sparse[2] = { text: 'valid' };
    values.push(sparse);
    const inherited = new Array<unknown>(3);
    Object.setPrototypeOf(inherited, Object.assign(Object.create(Array.prototype), {
      1: { prototype: 'forbidden' },
    }));
    values.push(inherited);
    const hidden = Object.create({ constructor: true });
    Object.defineProperty(hidden, 'prototype', { value: true });
    Object.defineProperty(hidden, Symbol('hidden'), { value: hidden, enumerable: true });
    values.push(hidden);
    for (const value of values) {
      expect(outcome(validate, value, 'root')).toEqual(outcome(reference, value, 'root'));
    }
  });

  test(`${name} tree validation preserves getter traversal and captured array length`, () => {
    const fixture = (): { value: unknown; visited: string[] } => {
      const visited: string[] = [];
      const array: unknown[] = [{}, {}];
      Object.defineProperty(array, 0, {
        enumerable: true,
        get() {
          visited.push('0');
          array.push({ prototype: true });
          return { child: 'valid' };
        },
      });
      Object.defineProperty(array, 1, {
        enumerable: true,
        get() {
          visited.push('1');
          return { child: 'valid' };
        },
      });
      return { value: { array }, visited };
    };
    const expected = fixture();
    const actual = fixture();
    expect(outcome(validate, actual.value, 'root'))
      .toEqual(outcome(reference, expected.value, 'root'));
    expect(actual.visited).toEqual(expected.visited);
  });

  test(`${name} tree validation preserves active-set state after an error`, () => {
    const child = { prototype: true };
    const root = { child };
    const expected = new WeakSet<object>();
    const actual = new WeakSet<object>();
    expect(outcome(validate, root, 'root', actual))
      .toEqual(outcome(reference, root, 'root', expected));
    for (const value of [root, child]) {
      expect(actual.has(value)).toBe(expected.has(value));
    }
    const preexisting = new WeakSet<object>([root]);
    expect(outcome(validate, root, 'custom', preexisting))
      .toEqual(outcome(reference, root, 'custom', new WeakSet<object>([root])));
    expect(preexisting.has(root)).toBe(true);
  });

  test(`${name} tree validation preserves getter errors and active-set cleanup`, () => {
    const child = Object.defineProperty({}, 'value', {
      enumerable: true,
      get() {
        throw new RangeError('getter failed');
      },
    });
    const root = { child };
    const expected = new WeakSet<object>();
    const actual = new WeakSet<object>();
    expect(outcome(validate, root, 'root', actual))
      .toEqual(outcome(reference, root, 'root', expected));
    for (const value of [root, child]) {
      expect(actual.has(value)).toBe(expected.has(value));
    }
  });

  test(`${name} tree validation accepts a large valid tree with shared children`, () => {
    const formatting = { bold: true, fontFamily: { ascii: 'Arial' } };
    const value = {
      document: {
        content: Array.from({ length: 298 }, (_, index) => ({
          type: 'paragraph',
          paraId: String(index),
          content: Array.from({ length: 8 }, () => ({
            type: 'run', formatting, content: [{ type: 'text', text: 'Document text' }],
          })),
        })),
      },
    };
    expect(outcome(reference, value, 'root')).toBeUndefined();
    expect(outcome(validate, value, 'root')).toBeUndefined();
    for (const primitive of [null, undefined, true, 1, 'text', Symbol('value'), () => undefined]) {
      expect(outcome(validate, primitive, 'root')).toBeUndefined();
    }
  });
}

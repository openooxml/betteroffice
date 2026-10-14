/** @internal Clones command data in runtimes with or without native structured cloning. */
export function cloneData<T>(value: T): T {
  const runtime = globalThis as typeof globalThis & { structuredClone?: <V>(input: V) => V };
  if (typeof runtime.structuredClone === 'function') return runtime.structuredClone(value);
  return copy(value, new Map()) as T;
}

function copy(value: unknown, seen: Map<object, unknown>): unknown {
  if (typeof value === 'function' || typeof value === 'symbol') {
    throw new TypeError('Unsupported command data');
  }
  if (value === null || typeof value !== 'object') return value;
  if (seen.has(value)) return seen.get(value);
  if ((typeof Blob !== 'undefined' && value instanceof Blob) ||
    (typeof File !== 'undefined' && value instanceof File)) {
    seen.set(value, value);
    return value;
  }
  if (ArrayBuffer.isView(value)) {
    const buffer = copy(value.buffer, seen) as ArrayBufferLike;
    const constructor = value.constructor as unknown as {
      new (buffer: ArrayBufferLike, byteOffset: number, length: number): ArrayBufferView;
      BYTES_PER_ELEMENT: number;
    };
    const result = value instanceof DataView
      ? new DataView(buffer, value.byteOffset, value.byteLength)
      : new constructor(buffer, value.byteOffset, value.byteLength / constructor.BYTES_PER_ELEMENT);
    seen.set(value, result);
    return result;
  }
  if (typeof SharedArrayBuffer !== 'undefined' && value instanceof SharedArrayBuffer) {
    seen.set(value, value);
    return value;
  }
  if (value instanceof ArrayBuffer) {
    const result = value.slice(0);
    seen.set(value, result);
    return result;
  }
  if (value instanceof Date) {
    const result = new Date(value.getTime());
    seen.set(value, result);
    return result;
  }
  if (value instanceof RegExp) {
    const result = new RegExp(value.source, value.flags);
    seen.set(value, result);
    return result;
  }
  if (value instanceof Number || value instanceof String || value instanceof Boolean ||
    Object.getPrototypeOf(value) === BigInt.prototype) {
    const result: object = Object(value.valueOf());
    seen.set(value, result);
    return result;
  }
  if (value instanceof Map) {
    const result = new Map<unknown, unknown>();
    seen.set(value, result);
    for (const [key, entry] of value) result.set(copy(key, seen), copy(entry, seen));
    return result;
  }
  if (value instanceof Set) {
    const result = new Set<unknown>();
    seen.set(value, result);
    for (const entry of value) result.add(copy(entry, seen));
    return result;
  }
  const result = Array.isArray(value) ? new Array<unknown>(value.length) : {};
  seen.set(value, result);
  for (const key of Object.keys(value)) {
    Object.defineProperty(result, key, {
      value: copy((value as Record<string, unknown>)[key], seen),
      enumerable: true, configurable: true, writable: true,
    });
  }
  return result;
}

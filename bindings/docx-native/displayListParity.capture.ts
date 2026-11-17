import { plugin } from 'bun';
import { createHash } from 'node:crypto';

/**
 * Preload for the display-list parity child: PNG export hands each page to the
 * engine's `renderPng`, which this replaces with one that records the page and
 * draws nothing. List it before the native engine preload so it wins that module.
 *
 * Each recording carries, for every face id the page or its chains name, the
 * sha256 of the bytes the session registered under that id, so a comparison
 * that renumbers ids still sees which face each id stands for.
 */
interface Face { bytes: string; view?: string }
type Method = (this: object, ...args: never[]) => unknown;

const stores = new WeakMap<object, Map<number, Face>>();
let probed: object | undefined;

const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

const canonicalNames: Record<string, string> = {
  'ｍｓ 明朝': 'ms mincho',
  'ｍｓ ｐ明朝': 'ms pmincho',
  'ｍｓ ゴシック': 'ms gothic',
  'ｍｓ ｐゴシック': 'ms pgothic',
};

function substituteName(family: string): string {
  const name = family.trim().toLowerCase();
  return canonicalNames[name] ?? name;
}

function storeOf(session: object): Map<number, Face> {
  let store = stores.get(session);
  if (!store) stores.set(session, store = new Map());
  return store;
}

function observe(EditSession: { prototype: Record<string, Method> }): void {
  const proto = EditSession.prototype;
  const register = proto.register_measure_font!;
  const substitute = proto.register_substitute_measure_font!;
  const clear = proto.clear_measure_fonts!;
  const outline = proto.outline_glyph_json!;
  proto.register_measure_font = function (this: object, bytes: Uint8Array) {
    const id = Reflect.apply(register, this, [bytes]) as number;
    storeOf(this).set(id, { bytes: digest(bytes) });
    return id;
  } as Method;
  proto.register_substitute_measure_font = function (this: object, base: number, family: string) {
    const id = Reflect.apply(substitute, this, [base, family]) as number;
    const face = storeOf(this).get(base);
    if (id !== base && face) storeOf(this).set(id, { bytes: face.bytes, view: substituteName(family) });
    return id;
  } as Method;
  proto.update_measure_font_outlines = function (this: object) {
    throw new Error('display-list parity: parity capture does not track outline updates');
  } as Method;
  proto.clear_measure_fonts = function (this: object) {
    Reflect.apply(clear, this, []);
    storeOf(this).clear();
  } as Method;
  proto.outline_glyph_json = function (this: object, fontId: number, glyphId: number) {
    probed = this;
    return Reflect.apply(outline, this, [fontId, glyphId]);
  } as Method;
}

function faceIds(fontChains: Record<string, number[]> | undefined, page: unknown): number[] {
  const ids = new Set(Object.values(fontChains ?? {}).flat());
  const walk = (value: unknown): void => {
    if (Array.isArray(value)) value.forEach(walk);
    else if (value !== null && typeof value === 'object') {
      for (const [key, field] of Object.entries(value)) {
        if (key === 'fontId' && typeof field === 'number') ids.add(field);
        else walk(field);
      }
    }
  };
  walk(page);
  return [...ids];
}

function record(engine: { fontChains?: Record<string, number[]>; page: unknown; outlineGlyph(fontId: number, glyphId: number): string }): string {
  const ids = faceIds(engine.fontChains, engine.page);
  const faces: Record<number, string> = {};
  if (ids.length > 0) {
    probed = undefined;
    try { engine.outlineGlyph(ids[0]!, 0); } catch { /* only the session the call reached matters */ }
    const store = probed && stores.get(probed);
    const kind = (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] === 'native' ? 'native' : 'wasm';
    for (const id of ids) {
      const face = store?.get(id);
      if (!face) throw new Error(`display-list parity: face id ${id} has no registered bytes on the ${kind} engine`);
      faces[id] = face.view === undefined ? face.bytes : `${face.bytes}/${face.view}`;
    }
  }
  return JSON.stringify({ fontChains: engine.fontChains, faces, page: engine.page });
}

const globals = globalThis as Record<symbol, unknown>;
globals[Symbol.for('betteroffice.displayListParity')] = [];
globals[Symbol.for('betteroffice.displayListParity.observe')] = observe;
globals[Symbol.for('betteroffice.displayListParity.record')] = record;

plugin({
  name: 'betteroffice-display-list-parity',
  setup(build) {
    build.onLoad({ filter: /[\\/]packages[\\/]docx[\\/]src[\\/]wasm[\\/]enginePng\.ts$/ }, () => ({
      loader: 'ts',
      contents: `
        import { EditSession } from './engineEdit';
        const key = Symbol.for('betteroffice.displayListParity');
        globalThis[Symbol.for('betteroffice.displayListParity.observe')](EditSession);
        const record = globalThis[Symbol.for('betteroffice.displayListParity.record')];
        export const engineKind = globalThis[Symbol.for('betteroffice.testEngine')] === 'native' ? 'native' : 'wasm';
        export async function renderPng(engine) {
          globalThis[key].push(record(engine));
          return new Uint8Array();
        }
        export default renderPng;
      `,
    }));
  },
});

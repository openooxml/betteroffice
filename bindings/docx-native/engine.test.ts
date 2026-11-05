import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import type { DecorationPrimitive, DisplayList, GlyphRunPrimitive } from '../../packages/docx/src/layout/render/displayList';
import initWasm, * as wasm from '../../packages/docx/src/wasm/generated/edit/docx_edit.js';
import initPreflight, { font_preflight as wasmFontPreflight } from '../../packages/docx/src/wasm/generated/preflight/docx_preflight.js';
import { expectPngRule, expectPngTracking, pngPixels } from '../raster-pixels';

type Engine = typeof wasm;
type Session = wasm.EditSession;
type ParagraphRead = {
  ok: boolean;
  version: string;
  paragraphs: { story: string; paraId: string; text: string }[];
};

const native = require('./index.js') as Engine;
const request = JSON.stringify({ story: 'body', view: 'accepted' });
let fixture: Uint8Array;
let fontBytes: Uint8Array;

beforeAll(async () => {
  await initWasm({
    module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)),
  });
  await initPreflight({
    module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/preflight/docx_preflight_bg.wasm', import.meta.url)),
  });
  fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
  const encoder = new TextEncoder();
  fixture = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
  });
});

function read(session: Session): ParagraphRead {
  const result = JSON.parse(session.read_paragraphs_json(request)) as ParagraphRead;
  expect(result.ok).toBe(true);
  expect(result.version).toMatch(/^[0-9a-f]{16}-\d+$/);
  return { ...result, version: '<version>' };
}

function paragraphId(session: Session): string {
  const paragraph = read(session).paragraphs[0];
  expect(paragraph).toBeDefined();
  return paragraph.paraId;
}

function expectPlainBytes(value: Uint8Array) {
  expect(value).toBeInstanceOf(Uint8Array);
  expect(Buffer.isBuffer(value)).toBe(false);
  expect(Object.getPrototypeOf(value)).toBe(Uint8Array.prototype);
  expect(value.byteOffset).toBe(0);
  expect(value.buffer.byteLength).toBe(value.byteLength);
}

function caughtValue(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected the engine call to throw');
}

function expectThrownParity(wasmCall: () => unknown, nativeCall: () => unknown): unknown {
  const expected = caughtValue(wasmCall);
  const actual = caughtValue(nativeCall);
  expect(typeof actual).toBe(typeof expected);
  expect(actual instanceof Error).toBe(expected instanceof Error);
  if (actual instanceof Error && expected instanceof Error) {
    expect(actual.name).toBe(expected.name);
    expect(actual.message).toBe(expected.message);
    expect((actual as Error & { code?: string }).code)
      .toBe((expected as Error & { code?: string }).code);
  } else {
    expect(actual).toBe(expected);
  }
  return actual;
}

describe('docx native engine parity', () => {
  test.each(['open_docx', 'seed_from_docx'] as const)('%s reads the same synthetic package', (method) => {
    const expected = new wasm.EditSession(7);
    const actual = new native.EditSession(7);
    try {
      if (method === 'open_docx') {
        expected.open_docx(fixture, true, 'opening');
        actual.open_docx(fixture, true, 'opening');
      } else {
        expected.seed_from_docx(fixture, 'opening');
        actual.seed_from_docx(fixture, 'opening');
      }
      expect(actual.client_id()).toBe(7);
      expect(read(actual)).toEqual(read(expected));
      expect(read(actual).paragraphs.map((paragraph) => paragraph.text)).toEqual(['Engine text']);
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('native insert_text supports undo and redo', () => {
    const expected = new wasm.EditSession(11);
    const session = new native.EditSession(11);
    try {
      expected.seed_from_docx(fixture, 'opening');
      session.seed_from_docx(fixture, 'opening');
      const before = read(session);
      session.track_undo();
      expect(session.can_undo()).toBe(false);
      expect(session.insert_text('body', paragraphId(session), 0, 'Native '))
        .toBe(expected.insert_text('body', paragraphId(expected), 0, 'Native '));
      const edited = read(session);
      expect(edited).toEqual(read(expected));
      expect(edited.paragraphs[0].text).toBe('Native Engine text');
      expect(session.can_undo()).toBe(true);
      expect(session.undo()).toBe(true);
      expect(read(session)).toEqual(before);
      expect(session.can_redo()).toBe(true);
      expect(session.redo()).toBe(true);
      expect(read(session)).toEqual(edited);
    } finally {
      expected.free();
      session.free();
    }
  });

  test('undo grouping reads the JavaScript clock, as wasm does', () => {
    let time = 1_000;
    const clock = spyOn(Date, 'now').mockImplementation(() => time);
    const texts = (engine: Engine) => {
      const session = new engine.EditSession(13);
      try {
        session.seed_from_docx(fixture, 'opening');
        session.track_undo();
        const steps: string[] = [];
        for (const [at, text] of [[1_000, 'A'], [1_100, 'B'], [2_000, 'C']] as const) {
          time = at;
          session.insert_text('body', paragraphId(session), 0, text);
        }
        steps.push(read(session).paragraphs[0].text);
        while (session.undo()) steps.push(read(session).paragraphs[0].text);
        return steps;
      } finally {
        session.free();
      }
    };
    try {
      const expected = texts(wasm);
      expect(expected).toEqual(['CBAEngine text', 'BAEngine text', 'Engine text']);
      expect(texts(native)).toEqual(expected);
    } finally {
      clock.mockRestore();
    }
  });

  test('native edits survive a throwing clock and a clock that reads the session', () => {
    const session = new native.EditSession(17);
    let reentrant: (() => unknown) | undefined;
    let depth = 0;
    const clock = spyOn(Date, 'now').mockImplementation(() => {
      if (!reentrant) throw new Error('clock unavailable');
      if (depth === 0) {
        depth += 1;
        try {
          reentrant();
        } finally {
          depth -= 1;
        }
      }
      return 5_000;
    });
    try {
      session.seed_from_docx(fixture, 'opening');
      session.track_undo();
      session.insert_text('body', paragraphId(session), 0, 'A');
      expect(read(session).paragraphs[0].text).toBe('AEngine text');
      reentrant = () => read(session);
      session.insert_text('body', paragraphId(session), 0, 'B');
      expect(read(session).paragraphs[0].text).toBe('BAEngine text');
      expect(session.undo()).toBe(true);
    } finally {
      clock.mockRestore();
      session.free();
    }
  });

  test('worker updates applied as local edits group by the JavaScript clock, as wasm does', () => {
    let time = 1_000;
    const clock = spyOn(Date, 'now').mockImplementation(() => time);
    const texts = (engine: Engine) => {
      const source = new wasm.EditSession(21);
      const target = new engine.EditSession(23);
      try {
        source.seed_from_docx(fixture, 'opening');
        target.load(source.encode_state());
        target.track_undo();
        for (const [at, text] of [[1_000, 'A'], [1_100, 'B'], [2_000, 'C']] as const) {
          source.insert_text('body', paragraphId(source), 0, text);
          const update = source.encode_diff(target.encode_state_vector());
          time = at;
          target.apply_local_update(update);
        }
        const steps = [read(target).paragraphs[0].text];
        while (target.undo()) steps.push(read(target).paragraphs[0].text);
        return steps;
      } finally {
        target.free();
        source.free();
      }
    };
    try {
      const expected = texts(wasm);
      expect(expected).toEqual(['CBAEngine text', 'BAEngine text', 'Engine text']);
      expect(texts(native)).toEqual(expected);
    } finally {
      clock.mockRestore();
    }
  });

  test('native calls that borrow typed arrays never run the JavaScript clock', () => {
    const session = new native.EditSession(19);
    const bytes = new Uint8Array(fontBytes);
    let reads = 0;
    const clock = spyOn(Date, 'now').mockImplementation(() => {
      reads += 1;
      bytes.buffer.transfer(0);
      return 1_000;
    });
    try {
      expect(session.register_measure_font(bytes)).toBe(0);
      expect(reads).toBe(0);
      expect(bytes.byteLength).toBe(fontBytes.byteLength);
      session.seed_from_docx(fixture, 'opening');
      const vector = new Uint8Array(new Uint8Array([0]));
      clock.mockImplementation(() => {
        reads += 1;
        vector.buffer.transfer(0);
        return 1_000;
      });
      expect(session.encode_diff_chunks(vector, 1 << 20).length).toBeGreaterThan(0);
      expect(reads).toBe(0);
      expect(vector.byteLength).toBe(1);
    } finally {
      clock.mockRestore();
      session.free();
    }
  });

  test('state crosses both boundaries and native diffs update a wasm peer', () => {
    const nativeSource = new native.EditSession(21);
    const wasmPeer = new wasm.EditSession(22);
    const wasmSource = new wasm.EditSession(23);
    const nativePeer = new native.EditSession(24);
    try {
      nativeSource.seed_from_docx(fixture, 'opening');
      const nativeState = nativeSource.encode_state();
      expectPlainBytes(nativeState);
      wasmPeer.load(nativeState);
      expect(read(wasmPeer)).toEqual(read(nativeSource));

      wasmSource.seed_from_docx(fixture, 'opening');
      nativePeer.load(wasmSource.encode_state());
      expect(read(nativePeer)).toEqual(read(wasmSource));

      nativeSource.insert_text('body', paragraphId(nativeSource), 0, 'Synced ');
      const diff = nativeSource.encode_diff(wasmPeer.encode_state_vector());
      expectPlainBytes(diff);
      expect(wasmPeer.apply_update(diff)).toBeUndefined();
      expect(read(wasmPeer)).toEqual(read(nativeSource));

      wasmSource.insert_text('body', paragraphId(wasmSource), 0, 'Remote ');
      nativePeer.apply_update(wasmSource.encode_diff(nativePeer.encode_state_vector()));
      expect(read(nativePeer)).toEqual(read(wasmSource));
    } finally {
      nativePeer.free();
      wasmSource.free();
      wasmPeer.free();
      nativeSource.free();
    }
  });

  test('chunk arrays contain plain bytes and reject invalid entries before applying', () => {
    const source = new native.EditSession(31);
    const peer = new wasm.EditSession(32);
    const target = new native.EditSession(33);
    try {
      source.seed_from_docx(fixture, 'opening');
      const chunks = source.encode_diff_chunks(peer.encode_state_vector(), 64);
      expect(Array.isArray(chunks)).toBe(true);
      expect(chunks.length).toBeGreaterThan(0);
      for (const chunk of chunks) expectPlainBytes(chunk);
      peer.apply_update_chunks(chunks);
      expect(read(peer)).toEqual(read(source));

      const fromWasm = peer.encode_diff_chunks(target.encode_state_vector(), 64);
      const before = target.encode_state();
      expect(expectThrownParity(
        () => peer.apply_update_chunks([chunks[0], new Uint32Array([1])]),
        () => target.apply_update_chunks([chunks[0], new Uint32Array([1])]),
      )).toBe('sync updates must be Uint8Arrays');
      expect(target.encode_state()).toEqual(before);
      target.apply_update_chunks(fromWasm);
      expect(read(target)).toEqual(read(peer));
    } finally {
      target.free();
      peer.free();
      source.free();
    }
  });

  test('queued update observation is synchronous and can be cleared', () => {
    for (const engine of [wasm, native]) {
      const session = new engine.EditSession(41);
      try {
        session.seed_from_docx(fixture, 'opening');
        expect(session.start_update_event_observation()).toBeUndefined();
        session.start_update_event_observation();
        session.insert_text('body', paragraphId(session), 0, 'Observed ');
        const update = session.drain_update_event();
        expectPlainBytes(update);
        expect(update.byteLength).toBeGreaterThan(1);
        expect(update[0]).toBe(0);
        session.clear_update_event_observation();
        expect(session.drain_update_event().byteLength).toBe(0);
        session.insert_text('body', paragraphId(session), 0, 'Unobserved ');
        expect(session.drain_update_event().byteLength).toBe(0);
      } finally {
        session.free();
      }
    }
  });

  test('resident fonts decode, shape and outline identically', () => {
    const expected = new wasm.EditSession(81);
    const actual = new native.EditSession(81);
    try {
      const font = expected.register_measure_font(fontBytes);
      expect(actual.register_measure_font(fontBytes)).toBe(font);
      expect(actual.register_substitute_measure_font(font, 'Calibri'))
        .toBe(expected.register_substitute_measure_font(font, 'Calibri'));
      const decoded = actual.decode_font(fontBytes);
      expectPlainBytes(decoded);
      expect(decoded).toEqual(expected.decode_font(fontBytes));
      expect(decoded).toEqual(fontBytes);
      expect(actual.take_base_latin_miss()).toBe(expected.take_base_latin_miss());
      const echo = actual.echo_glyph_set_json(font, 11, true, 'Ab!');
      expect(echo).toBe(expected.echo_glyph_set_json(font, 11, true, 'Ab!'));
      const glyph = JSON.parse(echo).glyphIds[0] as number;
      expect(actual.outline_glyph_json(font, glyph)).toBe(expected.outline_glyph_json(font, glyph));
      const invalid = new Uint8Array([1, 2, 3]);
      expectThrownParity(() => expected.decode_font(invalid), () => actual.decode_font(invalid));
      expectThrownParity(
        () => expected.register_measure_font(invalid),
        () => actual.register_measure_font(invalid),
      );
      expectThrownParity(
        () => expected.echo_glyph_set_json(font, 0, true, 'A'),
        () => actual.echo_glyph_set_json(font, 0, true, 'A'),
      );
      expected.clear_measure_fonts();
      actual.clear_measure_fonts();
      expect(actual.register_measure_font(fontBytes)).toBe(expected.register_measure_font(fontBytes));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('region layouts, retained metadata, frames and resident input match', () => {
    const expected = new wasm.EditSession(91);
    const actual = new native.EditSession(91);
    try {
      expected.open_docx(fixture, true, 'opening');
      actual.open_docx(fixture, true, 'opening');
      const font = expected.register_measure_font(fontBytes);
      expect(actual.register_measure_font(fontBytes)).toBe(font);
      const fontChains = { 'calibri|0|0': [font] };
      const input = JSON.stringify({
        bodyStory: 'body',
        regions: { sections: [{
          sectionId: 'main',
          properties: {
            pageWidth: 4320, pageHeight: 2880,
            marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
          },
        }] },
        measurement: {
          fontChains,
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          authoritativeShaping: true,
        },
        renderEnv: {},
      });
      expect(actual.layout_font_requirements_json(input)).toBe(expected.layout_font_requirements_json(input));
      const layout = actual.layout_document_with_regions_json(input);
      expect(layout).toBe(expected.layout_document_with_regions_json(input));
      expect(JSON.parse(layout).layout.pages.length).toBeGreaterThan(0);
      const expectedMeta = expected.layout_document_with_regions_retained_meta(input);
      const actualMeta = actual.layout_document_with_regions_retained_meta(input);
      try {
        expect(actualMeta).toBeInstanceOf(native.RetainedLayoutMeta);
        expect(actualMeta.page_count).toBe(expectedMeta.page_count);
        expect(actualMeta.partial).toBe(expectedMeta.partial);
        expect(actualMeta.provisional).toBe(expectedMeta.provisional);
        expect(actualMeta.notes_converged).toBe(expectedMeta.notes_converged);
        expect(actualMeta.page_sizes()).toBeInstanceOf(Float64Array);
        expect(actualMeta.page_sizes()).toEqual(expectedMeta.page_sizes());
        expect(actualMeta.layout_shell_json()).toBe(expectedMeta.layout_shell_json());
      } finally {
        actualMeta.free();
        expectedMeta.free();
      }
      expect(actual.retained_layout_json()).toBe(expected.retained_layout_json());
      expect(actual.retained_headers_footers_json()).toBe(expected.retained_headers_footers_json());
      expect(actual.retained_kernel_inputs_json()).toBe(expected.retained_kernel_inputs_json());
      const display = JSON.stringify({ fontChains });
      const frame = actual.build_display_list_frame(display, 0);
      expectPlainBytes(frame);
      expect(frame).toEqual(expected.build_display_list_frame(display, 0));
      expected.set_selection('body', paragraphId(expected), 2, paragraphId(expected), 2);
      actual.set_selection('body', paragraphId(actual), 2, paragraphId(actual), 2);
      const caret = actual.resident_caret_snapshot_json();
      expect(caret).toBe(expected.resident_caret_snapshot_json());
      const { frameEpoch, caretRect: rect } = JSON.parse(caret);
      expect(rect).not.toBeNull();
      const y = rect.y + rect.height / 2;
      expect(actual.display_hit_test_regions_json(rect.pageIndex, rect.x, y))
        .toBe(expected.display_hit_test_regions_json(rect.pageIndex, rect.x, y));
      expect(actual.display_range_rects_json(1, 4)).toBe(expected.display_range_rects_json(1, 4));
      expect(actual.display_range_rects_on_pages_json(1, 4, 0, 0))
        .toBe(expected.display_range_rects_on_pages_json(1, 4, 0, 0));
      expect(actual.display_range_rects_region_json('body', '', 1, 4))
        .toBe(expected.display_range_rects_region_json('body', '', 1, 4));
      const ranges = new Float64Array([1, 4, 2, 2]);
      const rects = actual.display_range_rects_batch('body', '', ranges);
      expect(rects).toBeInstanceOf(Float64Array);
      expect(rects).toEqual(expected.display_range_rects_batch('body', '', ranges));
      expect(actual.display_vertical_move_json(2, 'down', rect.x))
        .toBe(expected.display_vertical_move_json(2, 'down', rect.x));
      const edited = actual.apply_input('!', frameEpoch);
      expectPlainBytes(edited);
      expect(edited).toEqual(expected.apply_input('!', frameEpoch));
      expect(actual.resident_caret_snapshot_json()).toBe(expected.resident_caret_snapshot_json());
      expect(read(actual)).toEqual(read(expected));
      const nextEpoch = JSON.parse(actual.resident_caret_snapshot_json()).frameEpoch;
      const deleted = actual.apply_delete('backward', nextEpoch, 1);
      expectPlainBytes(deleted);
      expect(deleted).toEqual(expected.apply_delete('backward', nextEpoch, 1));
      expect(actual.resident_deleted_units()).toBe(expected.resident_deleted_units());
      expect(actual.resident_caret_snapshot_json()).toBe(expected.resident_caret_snapshot_json());
      const epoch = JSON.parse(actual.resident_caret_snapshot_json()).frameEpoch;
      expect(actual.build_display_pages_frame(new Uint32Array([0]), epoch))
        .toEqual(expected.build_display_pages_frame(new Uint32Array([0]), epoch));
      expectThrownParity(
        () => expected.build_display_list_frame(display, -1),
        () => actual.build_display_list_frame(display, -1),
      );
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('update callbacks deliver matching bytes synchronously and can be replaced and cleared', () => {
    const seed = new wasm.EditSession(101);
    const expected = new wasm.EditSession(102);
    const actual = new native.EditSession(102);
    const peer = new wasm.EditSession(103);
    const remote = new wasm.EditSession(104);
    const expectedUpdates: [Uint8Array, number][] = [];
    const actualUpdates: [Uint8Array, number][] = [];
    try {
      seed.seed_from_docx(fixture, 'opening');
      const state = seed.encode_state();
      for (const session of [expected, actual, peer, remote]) session.load(state);
      const observe = (updates: [Uint8Array, number][]) => (bytes: Uint8Array, isRemote: number) => {
        expectPlainBytes(bytes);
        expect([0, 1]).toContain(isRemote);
        updates.push([bytes, isRemote]);
      };
      expected.set_update_observer(observe(expectedUpdates));
      actual.set_update_observer(observe(actualUpdates));
      expected.insert_text('body', paragraphId(expected), 0, 'Local ');
      actual.insert_text('body', paragraphId(actual), 0, 'Local ');
      expect(actualUpdates.length).toBeGreaterThan(0);
      expect(actualUpdates).toEqual(expectedUpdates);
      expect(actualUpdates.every(([, origin]) => origin === 0)).toBe(true);
      const replacedCalls = actualUpdates.length;
      const replacedExpectedCalls = expectedUpdates.length;
      const replacementExpected: [Uint8Array, number][] = [];
      const replacementActual: [Uint8Array, number][] = [];
      expected.set_update_observer(observe(replacementExpected));
      actual.set_update_observer(observe(replacementActual));
      remote.insert_text('body', paragraphId(remote), 0, 'Remote ');
      const update = remote.encode_diff(seed.encode_state_vector());
      expected.apply_update(update);
      actual.apply_update(update);
      expect(actualUpdates.length).toBe(replacedCalls);
      expect(expectedUpdates.length).toBe(replacedExpectedCalls);
      expect(replacementActual).toEqual(replacementExpected);
      expect(replacementActual.some(([, origin]) => origin === 1)).toBe(true);
      for (const [bytes] of [...actualUpdates, ...replacementActual]) peer.apply_update(bytes);
      expect(read(peer)).toEqual(read(actual));
      expect(read(actual)).toEqual(read(expected));
      expected.clear_update_observer();
      actual.clear_update_observer();
      const clearedCalls = replacementActual.length;
      expected.insert_text('body', paragraphId(expected), 0, 'Cleared ');
      actual.insert_text('body', paragraphId(actual), 0, 'Cleared ');
      expect(replacementActual.length).toBe(clearedCalls);
      expect(replacementExpected.length).toBe(clearedCalls);
      let throwingCalls = 0;
      actual.set_update_observer(() => { throwingCalls += 1; throw new Error('observer'); });
      expected.insert_text('body', paragraphId(expected), 0, 'Throwing ');
      expect(() => actual.insert_text('body', paragraphId(actual), 0, 'Throwing ')).not.toThrow();
      expect(throwingCalls).toBeGreaterThan(0);
      expect(read(actual)).toEqual(read(expected));
      actual.clear_update_observer();
    } finally {
      remote.free();
      peer.free();
      actual.free();
      expected.free();
      seed.free();
    }
  });

  test('update callbacks can read versions and schedule state-vector reads', async () => {
    const seed = new wasm.EditSession(121);
    const expected = new wasm.EditSession(122);
    const actual = new native.EditSession(122);
    const remote = new wasm.EditSession(123);
    type Event = { update: Uint8Array; origin: number; version: string };
    const expectedEvents: Event[] = [];
    const actualEvents: Event[] = [];
    const expectedVectors: Uint8Array[] = [];
    const actualVectors: Uint8Array[] = [];
    try {
      seed.seed_from_docx(fixture, 'opening');
      const state = seed.encode_state();
      for (const session of [expected, actual, remote]) session.load(state);
      const expectedBefore = expected.version();
      const actualBefore = actual.version();
      const observe = (session: Session, events: Event[], vectors: Uint8Array[]) =>
        (update: Uint8Array, origin: number) => {
          events.push({ update, origin, version: session.version() });
          queueMicrotask(() => vectors.push(session.encode_state_vector()));
        };
      expected.set_update_observer(observe(expected, expectedEvents, expectedVectors));
      actual.set_update_observer(observe(actual, actualEvents, actualVectors));
      expected.insert_text('body', paragraphId(expected), 0, 'Local ');
      actual.insert_text('body', paragraphId(actual), 0, 'Local ');
      expect(expectedEvents).toHaveLength(1);
      expect(actualEvents).toHaveLength(1);
      expect(expectedEvents[0].version).toBe(expected.version());
      expect(actualEvents[0].version).toBe(actual.version());
      expect(expectedEvents[0].version).not.toBe(expectedBefore);
      expect(actualEvents[0].version).not.toBe(actualBefore);
      expect(actualEvents[0].update).toEqual(expectedEvents[0].update);
      expect(actualEvents[0].origin).toBe(0);
      expect(expectedEvents[0].origin).toBe(0);
      expect(actualEvents[0].version.slice(17)).toBe(expectedEvents[0].version.slice(17));
      await Promise.resolve();
      expect(actualVectors).toHaveLength(1);
      expect(actualVectors).toEqual(expectedVectors);
      expect(actualVectors[0]).toEqual(actual.encode_state_vector());
      expect(expectedVectors[0]).toEqual(expected.encode_state_vector());
      const expectedLocalVersion = expected.version();
      const actualLocalVersion = actual.version();
      remote.insert_text('body', paragraphId(remote), 0, 'Remote ');
      const update = remote.encode_diff(seed.encode_state_vector());
      expected.apply_update(update);
      actual.apply_update(update);
      expect(actualEvents.length).toBeGreaterThan(1);
      expect(actualEvents.length).toBe(expectedEvents.length);
      expect(expectedEvents.at(-1)?.version).toBe(expected.version());
      expect(actualEvents.at(-1)?.version).toBe(actual.version());
      expect(expected.version()).not.toBe(expectedLocalVersion);
      expect(actual.version()).not.toBe(actualLocalVersion);
      const normalize = (events: Event[], initial: string) => events.map((event) => {
        expect(event.version.slice(0, 16)).toBe(initial.slice(0, 16));
        return { ...event, version: event.version.slice(17) };
      });
      expect(actualEvents[0].update).toEqual(expectedEvents[0].update);
      expect(actualEvents.map(({ origin }) => origin)).toEqual(expectedEvents.map(({ origin }) => origin));
      const replay = (events: Event[]) => {
        const peer = new wasm.EditSession(124);
        try {
          peer.load(state);
          for (const { update } of events) peer.apply_update(update);
          return read(peer);
        } finally {
          peer.free();
        }
      };
      expect(replay(actualEvents)).toEqual(replay(expectedEvents));
      for (const event of actualEvents.slice(1)) expect(event.version).toBe(actual.version());
      expect(normalize(actualEvents.slice(-1), actualBefore))
        .toEqual(normalize(expectedEvents.slice(-1), expectedBefore));
      expect(actualEvents.slice(1).some((event) => event.origin === 1)).toBe(true);
      await Promise.resolve();
      expect(actualVectors.length).toBe(actualEvents.length);
      expect(actualVectors).toEqual(expectedVectors);
      expect(actualVectors.at(-1)).toEqual(actual.encode_state_vector());
      expect(expectedVectors.at(-1)).toEqual(expected.encode_state_vector());
      expect(read(actual)).toEqual(read(expected));
    } finally {
      remote.free();
      actual.free();
      expected.free();
      seed.free();
    }
  });

  test('native callbacks read committed session state synchronously', () => {
    const session = new native.EditSession(112);
    const remote = new wasm.EditSession(113);
    const snapshot = () => ({
      stateVector: session.encode_state_vector(),
      paragraphs: session.read_paragraphs_json(request),
      changedStories: session.stories_changed_since(0),
      version: session.version(),
      clientId: session.client_id(),
      directBatches: session.direct_batches_applied(),
    });
    const snapshots: ReturnType<typeof snapshot>[] = [];
    const drained: Uint8Array[] = [];
    try {
      session.seed_from_docx(fixture, 'opening');
      remote.load(session.encode_state());
      session.start_update_event_observation();
      session.set_update_observer(() => {
        snapshots.push(snapshot());
        drained.push(session.drain_update_event());
      });
      session.insert_text('body', paragraphId(session), 0, 'Local ');
      expect(snapshots).toEqual([snapshot()]);
      const localCalls = snapshots.length;
      remote.insert_text('body', paragraphId(remote), 0, 'Remote ');
      session.apply_update(remote.encode_diff(session.encode_state_vector()));
      expect(snapshots.length).toBeGreaterThan(localCalls);
      for (const observed of snapshots.slice(localCalls)) expect(observed).toEqual(snapshot());
      expect(drained).toHaveLength(snapshots.length);
      for (const update of drained) {
        expectPlainBytes(update);
        expect(update.byteLength).toBeGreaterThan(1);
      }
      expect(drained[0][0]).toBe(0);
      expect(drained.slice(localCalls).some((update) => update[0] === 1)).toBe(true);
      expect(session.drain_update_event()).toEqual(new Uint8Array());
    } finally {
      remote.free();
      session.free();
    }
  });

  test('native callbacks queue edits until the current callback returns', () => {
    const session = new native.EditSession(114);
    const peer = new wasm.EditSession(115);
    const updates: [Uint8Array, number][] = [];
    const calls: string[] = [];
    try {
      session.seed_from_docx(fixture, 'opening');
      peer.load(session.encode_state());
      const paraId = paragraphId(session);
      session.set_update_observer((update: Uint8Array, origin: number) => {
        updates.push([update, origin]);
        if (updates.length === 1) {
          calls.push('before edit');
          session.insert_text('body', paraId, 0, 'Nested ');
          calls.push('after edit');
        }
        calls.push(`callback ${updates.length}`);
      });
      session.insert_text('body', paraId, 0, 'Outer ');
      expect(calls).toEqual(['before edit', 'after edit', 'callback 1', 'callback 2']);
      expect(updates).toHaveLength(2);
      for (const [update, origin] of updates) {
        expectPlainBytes(update);
        expect(origin).toBe(0);
        peer.apply_update(update);
      }
      expect(read(session).paragraphs[0].text).toBe('Nested Outer Engine text');
      expect(read(peer)).toEqual(read(session));
    } finally {
      peer.free();
      session.free();
    }
  });

  test('clearing a native observer inside its callback stops queued delivery', () => {
    const session = new native.EditSession(116);
    let calls = 0;
    try {
      session.set_update_observer(() => {
        calls += 1;
        session.clear_update_observer();
      });
      session.load_json(JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'First' }] },
        { storyId: 'other', paragraphs: [{ text: 'Second' }] },
      ]));
      expect(calls).toBe(1);
      expect(session.story_ids()).toEqual(['body', 'other']);
      session.insert_text('body', paragraphId(session), 0, 'After ');
      expect(calls).toBe(1);
    } finally {
      session.free();
    }
  });

  test('freeing a native session inside its callback stops queued delivery', () => {
    const session = new native.EditSession(117);
    let calls = 0;
    try {
      session.set_update_observer(() => {
        calls += 1;
        session.free();
      });
      expect(() => session.load_json(JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'First' }] },
        { storyId: 'other', paragraphs: [{ text: 'Second' }] },
      ]))).not.toThrow();
      expect(calls).toBe(1);
      expect(() => session.version()).toThrow('EditSession was freed');
      expect(() => session.encode_state_vector()).toThrow('EditSession was freed');
      expect(() => session.read_paragraphs_json(request)).toThrow('EditSession was freed');
    } finally {
      session.free();
    }
  });

  test('a failing call preserves its own exception while delivering native callbacks', () => {
    const expected = new wasm.EditSession(118);
    const actual = new native.EditSession(118);
    const expectedUpdates: [Uint8Array, number][] = [];
    const actualUpdates: [Uint8Array, number][] = [];
    const vectors: Uint8Array[] = [];
    const replayed = (updates: [Uint8Array, number][]) => {
      const peer = new wasm.EditSession(119);
      try {
        for (const [update] of updates) peer.apply_update(update);
        return read(peer);
      } finally {
        peer.free();
      }
    };
    const origins = (updates: [Uint8Array, number][]) => updates.map(([, origin]) => origin);
    try {
      expected.set_update_observer((update: Uint8Array, origin: number) => {
        expectedUpdates.push([update, origin]);
        throw new Error('observer failure');
      });
      actual.set_update_observer((update: Uint8Array, origin: number) => {
        actualUpdates.push([update, origin]);
        vectors.push(actual.encode_state_vector());
        actual.set_undo_capture_mode('invalid');
      });
      const input = JSON.stringify([
        { storyId: 'body', paragraphs: [{ text: 'Committed' }] },
        {},
      ]);
      expect(expectThrownParity(() => expected.load_json(input), () => actual.load_json(input)))
        .toBe('a story entry requires a string "storyId"');
      expect(actualUpdates.length).toBeGreaterThan(0);
      expect(origins(actualUpdates)).toEqual(origins(expectedUpdates));
      expect(replayed(actualUpdates)).toEqual(replayed(expectedUpdates));
      expect(vectors).toHaveLength(actualUpdates.length);
      for (const vector of vectors) expect(vector).toEqual(actual.encode_state_vector());
      expect(read(actual)).toEqual(read(expected));
      const before = actualUpdates.length;
      expected.insert_text('body', paragraphId(expected), 0, 'After ');
      expect(() => actual.insert_text('body', paragraphId(actual), 0, 'After ')).not.toThrow();
      expect(actualUpdates.length).toBeGreaterThan(before);
      expect(origins(actualUpdates)).toEqual(origins(expectedUpdates));
      expect(replayed(actualUpdates)).toEqual(replayed(expectedUpdates));
      expect(read(actual)).toEqual(read(expected));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('free functions preserve ZIP shapes, parse JSON, exports and sanitization', () => {
    const expected = wasm.unzip_docx(fixture) as Record<string, Uint8Array>;
    const actual = native.unzip_docx(fixture) as Record<string, Uint8Array>;
    expect(Object.getPrototypeOf(actual)).toBe(Object.prototype);
    expect(Object.keys(actual)).toEqual(Object.keys(expected));
    for (const [name, bytes] of Object.entries(actual)) {
      expectPlainBytes(bytes);
      expect(bytes).toEqual(expected[name]);
    }
    const zipped = native.rezip_docx(actual);
    expectPlainBytes(zipped);
    expect(zipped).toEqual(wasm.rezip_docx(expected));
    expect(native.unzip_docx(zipped)).toEqual(expected);
    expect(native.parse_docx_s2(fixture)).toBe(wasm.parse_docx_s2(fixture));
    expect(native.export_docx_structured_json(fixture, '{"revisionView":"accepted"}'))
      .toBe(wasm.export_docx_structured_json(fixture, '{"revisionView":"accepted"}'));
    expect(native.paragraphPatchWriterVersion()).toBe(wasm.paragraphPatchWriterVersion());
    const sanitized = native.sanitizeOoxml(fixture, 'docx');
    expectPlainBytes(sanitized);
    expect(sanitized).toEqual(wasm.sanitizeOoxml(fixture, 'docx'));
    expectThrownParity(
      () => wasm.parse_docx_s2(new Uint8Array([1, 2, 3])),
      () => native.parse_docx_s2(new Uint8Array([1, 2, 3])),
    );
    for (const value of [null, undefined, 1, 'entries']) {
      expect(expectThrownParity(() => wasm.rezip_docx(value), () => native.rezip_docx(value)))
        .toBe('rezip_docx: expected an object');
    }
    const entries = Object.assign(Object.create({ ignored: new Uint8Array([9]) }), {
      data: [1, 2, 255],
      [Symbol('ignored')]: new Uint8Array([8]),
    });
    expect(native.unzip_docx(native.rezip_docx(entries)))
      .toEqual(wasm.unzip_docx(wasm.rezip_docx(entries)));
    expect(native.wasm_heap_counted()).toBe(false);
    native.set_wasm_heap_limit(1);
    expect(native.wasm_live_bytes()).toBe(0);
    expect(native.wasm_peak_bytes()).toBe(0);
    expect(native.wasm_failed_allocation_bytes()).toBe(0);
    expect(native.reset_wasm_peak_bytes()).toBeUndefined();
    expect(native.install_panic_hook()).toBeUndefined();
  });

  test('module measurement fonts share the free glyph API', () => {
    wasm.clear_measure_fonts();
    native.clear_measure_fonts();
    try {
      const font = wasm.register_measure_font(fontBytes);
      expect(native.register_measure_font(fontBytes)).toBe(font);
      expect(native.register_substitute_measure_font(font, 'Calibri'))
        .toBe(wasm.register_substitute_measure_font(font, 'Calibri'));
      expect(native.outline_glyph_json(font, 36)).toBe(wasm.outline_glyph_json(font, 36));
      expectThrownParity(
        () => wasm.register_measure_font(new Uint8Array([1])),
        () => native.register_measure_font(new Uint8Array([1])),
      );
    } finally {
      native.clear_measure_fonts();
      wasm.clear_measure_fonts();
    }
  });

  test.each([-1, 0.5, Number.NaN, Number.POSITIVE_INFINITY, 9_007_199_254_740_992])(
    'constructor rejects invalid client id %s with a string', (clientId) => {
      expect(expectThrownParity(
        () => new wasm.EditSession(clientId),
        () => new native.EditSession(clientId),
      )).toBe('client_id must be a non-negative safe integer');
    },
  );

  test('message strings and PeerMetadataError fields match wasm', () => {
    const expected = new wasm.EditSession(51);
    const actual = new native.EditSession(51);
    try {
      expect(expectThrownParity(
        () => expected.set_undo_capture_mode('invalid'),
        () => actual.set_undo_capture_mode('invalid'),
      )).toBe('undo capture mode must be auto or manual');
      const peerError = expectThrownParity(
        () => expected.encode_peer_metadata(),
        () => actual.encode_peer_metadata(),
      ) as Error & { code: string };
      expect(peerError.name).toBe('PeerMetadataError');
      expect(peerError.code).toBe('missing-source');
      expectThrownParity(() => expected.run_inheritance('{'), () => actual.run_inheritance('{'));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('optional returns use undefined', () => {
    const session = new native.EditSession(61);
    try {
      expect(session.materialize_docx()).toBeUndefined();
      expect(session.media_bytes('missing')).toBeUndefined();
      expect(session.media_type('missing')).toBeUndefined();
      expect(session.media_data_url('missing')).toBeUndefined();
      expect(session.clear_cell_selection()).toBeUndefined();
    } finally {
      session.free();
    }
  });

  test('prime_save_seed_states caches within the same budgets as wasm', () => {
    const expected = new wasm.EditSession(62);
    const actual = new native.EditSession(62);
    try {
      expected.seed_from_docx(fixture, 'opening');
      actual.seed_from_docx(fixture, 'opening');
      for (const [paragraphs, inflated] of [[0, 0], [1, 0], [1_000, 64 * 1024 * 1024]]) {
        expect(actual.prime_save_seed_states(paragraphs, inflated))
          .toBe(expected.prime_save_seed_states(paragraphs, inflated));
      }
      expect(read(actual)).toEqual(read(expected));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('prime_save_seed_states_slice answers each slice as wasm does', () => {
    const expected = new wasm.EditSession(63);
    const actual = new native.EditSession(63);
    try {
      expected.seed_from_docx(fixture, 'opening');
      actual.seed_from_docx(fixture, 'opening');
      for (const [paragraphs, inflated, bytes] of [[0, 0, 1], [1_000, 64 * 1024 * 1024, 1], [1_000, 64 * 1024 * 1024, 1]]) {
        expect(actual.prime_save_seed_states_slice(paragraphs, inflated, bytes))
          .toBe(expected.prime_save_seed_states_slice(paragraphs, inflated, bytes));
      }
      let answer = 1;
      for (let slice = 0; slice < 10_000 && answer === 1; slice += 1) {
        answer = expected.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64);
        expect(actual.prime_save_seed_states_slice(1_000, 64 * 1024 * 1024, 64)).toBe(answer);
      }
      expect(read(actual)).toEqual(read(expected));
    } finally {
      actual.free();
      expected.free();
    }
  });

  test('RunInheritance resolves the same attributes and formatting', () => {
    const expectedSession = new wasm.EditSession(71);
    const actualSession = new native.EditSession(71);
    const styles = JSON.stringify({
      docDefaults: { rPr: { fontSize: 22 } },
      styles: [
        { type: 'paragraph', styleId: 'Normal', default: true },
        { type: 'paragraph', styleId: 'Heading', rPr: { bold: true }, pPr: { spaceAfter: 80 } },
        { type: 'character', styleId: 'Emphasis', rPr: { italic: true } },
      ],
    });
    const expected = expectedSession.run_inheritance(styles);
    const actual = actualSession.run_inheritance(styles);
    try {
      expect(JSON.parse(actual.attributes())).toEqual(JSON.parse(expected.attributes()));
      expect(JSON.parse(actual.attributes('Heading', 'Emphasis')))
        .toEqual(JSON.parse(expected.attributes('Heading', 'Emphasis')));
      expect(JSON.parse(actual.attributes(null, null, '{"bold":true}')))
        .toEqual(JSON.parse(expected.attributes(null, null, '{"bold":true}')));
      expect(JSON.parse(actual.paragraph_formatting(undefined, 'null')))
        .toEqual(JSON.parse(expected.paragraph_formatting(undefined, 'null')));
      expect(JSON.parse(actual.paragraph_formatting('Heading', '{}')))
        .toEqual(JSON.parse(expected.paragraph_formatting('Heading', '{}')));
      const table = JSON.stringify({ type: 'table', rows: [{ cells: [{ type: 'tableCell', content: [] }] }] });
      expect(JSON.parse(actual.table_attributes(table))).toEqual(JSON.parse(expected.table_attributes(table)));
      expectThrownParity(() => expected.attributes(null, null, '{'), () => actual.attributes(null, null, '{'));
      expectThrownParity(() => expected.paragraph_formatting(null, '{'), () => actual.paragraph_formatting(null, '{'));
      expectThrownParity(() => expected.table_attributes('{'), () => actual.table_attributes('{'));
    } finally {
      actual.free();
      expected.free();
      actualSession.free();
      expectedSession.free();
    }
  });

  test('story_checksum of a million-character formatted run matches and stays within wasm memory', () => {
    const encoder = new TextEncoder();
    const text = 'lorem ipsum '.repeat(83_334).slice(0, 1_000_000);
    const rPr = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:b/><w:i/><w:color w:val="1F3864"/><w:sz w:val="22"/><w:u w:val="single"/><w:lang w:val="en-GB"/><w:highlight w:val="yellow"/></w:rPr>';
    const bytes = wasm.rezip_docx({
      '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
      '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
      'word/document.xml': encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
    });
    const expected = new wasm.EditSession(7);
    const actual = new native.EditSession(7);
    try {
      expected.open_docx(bytes, true);
      actual.open_docx(bytes, true);
      const checksum = expected.story_checksum('body');
      expect(checksum).toMatch(/^\d+$/);
      expect(actual.story_checksum('body')).toBe(checksum);
    } finally {
      actual.free();
      expected.free();
    }
  }, 120_000);

  test('decodeFontFace and font_preflight failures have the wasm error shape', () => {
    const nativePreflight = (native as unknown as { font_preflight: typeof wasmFontPreflight }).font_preflight;
    const invalid = new Uint8Array([1, 2, 3, 4]);
    for (const [codec, dict, maxOutput] of [['woff2', undefined, undefined], ['unknown', undefined, undefined], ['', undefined, undefined]] as const) {
      const actual = expectThrownParity(
        () => wasm.decodeFontFace(invalid, codec, dict, maxOutput),
        () => native.decodeFontFace(invalid, codec, dict, maxOutput),
      );
      expect(actual instanceof Error && 'code' in actual).toBe(false);
    }
    expectThrownParity(() => wasmFontPreflight(invalid), () => nativePreflight(invalid));
  });
});

test('native and wasm track and resolve character formatting with the same optional author arguments', () => {
  for (const engine of [wasm, native]) {
    for (const accept of [true, false]) {
      const session = new engine.EditSession(9303);
      try {
        session.open_docx(fixture, true, 'format-tracking');
        const id = paragraphId(session);
        session.toggle_mark('body', id, 0, id, 6, '{"type":"bold"}', 'Writer', '2026-10-01T12:00:00Z');
        session.format_range('body', id, 1, id, 5, '{"italic":true}', 'Writer', '2026-10-01T12:00:00Z');
        session.clear_formatting('body', id, 2, id, 4, 'Writer', '2026-10-01T12:00:00Z');
        const revisions = JSON.parse(session.list_revisions());
        expect(revisions).toEqual(expect.arrayContaining([expect.objectContaining({ kind: 'rPrChange', author: 'Writer' })]));
        for (const revisionId of new Set(revisions.map((revision: { revisionId: string }) => revision.revisionId))) {
          const target = JSON.stringify({ revisionId });
          if (accept) session.accept_change(target);
          else session.reject_change(target);
        }
        expect(JSON.parse(session.list_revisions())).toEqual([]);
        const context = JSON.parse(session.selection_context('body', id, 0, id, 1));
        expect(context.bold).toBe(accept);
        session.set_hyperlink('body', id, 0, id, 6, '{"href":"#target"}', 'Writer', '2026-10-01T12:00:00Z');
        expect(JSON.parse(session.list_revisions()).length).toBeGreaterThan(0);
      } finally {
        session.free();
      }
    }
  }
});

function pagePng(session: import('./index.js').EditSession, list: string, page: number, scale: number): Uint8Array {
  return session.renderPngTarget(list, `{"page":${page}}`, `{"scale":${scale}}`);
}

function pngDisplayList(
  session: import('./index.js').EditSession,
  families = ['Calibri'],
  font = session.register_measure_font(fontBytes),
): string {
  const fontChains = Object.fromEntries(families.flatMap((family) => [
    [`${family.toLowerCase()}|0|0`, [font]],
    [`${family.toLowerCase()}|1|0`, [font]],
  ]));
  const materialized = JSON.parse(session.materialize_docx()!) as {
    document: { package: { theme: { colorScheme: Record<string, string> } } };
  };
  const layout = session.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{
      sectionId: 'main',
      properties: {
        pageWidth: 4320, pageHeight: 2880,
        marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
      },
    }] },
    measurement: {
      fontChains,
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      authoritativeShaping: true,
    },
    renderEnv: { themeColors: materialized.document.package.theme.colorScheme },
  }));
  return session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }));
}

test('native docx rasterizes the compiled page at two scales', () => {
  const raw = require('./index.js') as typeof import('./index.js');
  const session = new raw.EditSession(401);
  try {
    session.open_docx(fixture, true);
    const list = pngDisplayList(session);
    const first = pagePng(session, list, 0, 1);
    const second = pagePng(session, list, 0, 2);
    for (const png of [first, second]) {
      expectPlainBytes(png);
      expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    }
    const one = Buffer.from(first);
    const two = Buffer.from(second);
    expect(two.readUInt32BE(16)).toBe(one.readUInt32BE(16) * 2);
    expect(two.readUInt32BE(20)).toBe(one.readUInt32BE(20) * 2);
    expect(pngPixels(first).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
    expect(pngPixels(second).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
  } finally {
    session.free();
  }
});

test('native docx PNG refuses invalid input, missing fonts and oversized surfaces', () => {
  const raw = require('./index.js') as typeof import('./index.js');
  const session = new raw.EditSession(402);
  try {
    session.open_docx(fixture, true);
    const list = pngDisplayList(session);
    for (const call of [
      () => pagePng(session, list, 9999, 1),
      () => pagePng(session, list, 0, 0),
      () => pagePng(session, list, 0, NaN),
      () => pagePng(session, list, 0, -1),
      () => pagePng(session, list, 0, Infinity),
      () => pagePng(session, '{', 0, 1),
      () => pagePng(session, '{}', 0, 1),
      () => pagePng(session, JSON.stringify({ pages: [{ pageIndex: 0, width: 0, height: 10, primitives: [] }] }), 0, 1),
    ]) {
      const error = caughtValue(call);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).message).toStartWith('png-invalid:');
      expect(error).not.toHaveProperty('code');
    }
    for (const call of [
      () => session.renderPngTarget(list, '{"page":0}', '{"scale":null}'),
      () => session.renderPngTarget(list, '{"page":0}', '{"scale":"2"}'),
      () => session.renderPngTarget(list, '{"page":0}', '{"maxWidth":10}'),
      () => session.renderPngTarget(list, '{"page":0,"extra":1}', '{}'),
      () => session.renderPngTarget(list, '{"slide":0}', '{}'),
      () => session.renderPngTarget(list, '{"page":-1}', '{}'),
      () => session.renderPngTarget(list, '{"page":0.5}', '{}'),
      () => session.renderPngTarget(list, '{}', '{}'),
      () => session.renderPngTarget(list, '{"page":0}', '['),
    ]) {
      const error = caughtValue(call);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).message).toStartWith('png-invalid:');
      expect(error).not.toHaveProperty('code');
    }
    expect(session.renderPngTarget(list, '{"page":0}', '{}')).toEqual(pagePng(session, list, 0, 1));
    for (const scale of [1e9, Number.MAX_VALUE]) {
      expect(() => pagePng(session, list, 0, scale)).toThrow('png-too-large:');
    }
    const area = JSON.stringify({ pages: [{ pageIndex: 0, width: 5000, height: 5000, primitives: [] }] });
    expect(() => pagePng(session, area, 0, 1)).toThrow('png-too-large:');
    const stale = JSON.parse(list);
    const glyph = stale.pages[0].primitives.find((primitive: { kind: string }) => primitive.kind === 'glyphRun');
    expect(glyph).toBeDefined();
    glyph.fontId = 0x7ffffffe;
    expect(() => pagePng(session, JSON.stringify(stale), 0, 1)).toThrow('png-font-unavailable:');
    const other = new raw.EditSession(404);
    try {
      other.open_docx(fixture, true);
      expect(() => pagePng(other, list, 0, 1)).toThrow('png-font-unavailable:');
      expect(pagePng(session, list, 0, 1).length).toBeGreaterThan(8);
    } finally {
      other.free();
    }
  } finally {
    session.free();
  }
});

test('native docx PNG paints theme accent text red', () => {
  const raw = require('./index.js') as typeof import('./index.js');
  const encoder = new TextEncoder();
  const themed = wasm.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'),
    'word/theme/theme1.xml': encoder.encode('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Red"><a:themeElements><a:clrScheme name="Red"><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme><a:fontScheme name="Default"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Default"/></a:themeElements></a:theme>'),
    'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:sz w:val="96"/><w:color w:themeColor="accent1"/></w:rPr><w:t>RED</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
  });
  const session = new raw.EditSession(403);
  try {
    session.open_docx(themed, true);
    const pixels = pngPixels(pagePng(session, pngDisplayList(session), 0, 1));
    let red = false;
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index] > 200 && pixels[index + 1] < 50 && pixels[index + 2] < 50 && pixels[index + 3] === 255) red = true;
    }
    expect(red).toBe(true);
  } finally {
    session.free();
  }
});

test('native docx PNG uses display list font chains for text primitives', () => {
  const raw = require('./index.js') as typeof import('./index.js');
  const session = new raw.EditSession(405);
  try {
    session.open_docx(fixture, true);
    const font = session.register_measure_font(fontBytes);
    const pages = [{ pageIndex: 0, width: 160, height: 60, primitives: [{
      kind: 'text', text: 'PNG', x: 5, baselineY: 30, width: 100,
      font: '24px Calibri', color: '#000000',
    }] }];
    const list = JSON.stringify({ pages, fontChains: { 'calibri|0|0': [font] } });
    expect(pngPixels(pagePng(session, list, 0, 1)).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
    expect(() => pagePng(session, JSON.stringify({ pages }), 0, 1)).toThrow('png-font-unavailable:');
  } finally {
    session.free();
  }
});

test('native docx PNG resolves media tokens in every page region and skips missing images', () => {
  const raw = require('./index.js') as typeof import('./index.js');
  const parts = wasm.unzip_docx(fixture) as Record<string, Uint8Array>;
  const blue = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYPj/HwADAgH/5ncLrgAAAABJRU5ErkJggg==';
  parts['word/media/image1.png'] = new Uint8Array(Buffer.from(blue, 'base64'));
  parts['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(parts['[Content_Types].xml'])
    .replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
  const session = new raw.EditSession(406);
  try {
    session.set_media_tokens(true);
    session.open_docx(wasm.rezip_docx(parts), true);
    expect(session.media_bytes('media:0')).toEqual(parts['word/media/image1.png']);
    const image = (x: number, relId = 'media:0') => ({ kind: 'image', relId, x, y: 10, w: 10, h: 10 });
    const list = JSON.stringify({ pages: [{
      pageIndex: 0, width: 160, height: 60,
      primitives: [image(10), image(110, 'media:999'), image(130, `data:image/png;base64,${blue}`)],
      header: { rId: 'header1', kind: 'header', y: 0, height: 30, primitives: [image(30)] },
      footer: { rId: 'footer1', kind: 'footer', y: 30, height: 30, primitives: [image(50)] },
      noteAreas: [{ primitives: [image(70)] }, { kind: 'endnote', primitives: [image(90)] }],
    }] });
    for (const scale of [1, 2]) {
      const png = pagePng(session, list, 0, scale);
      const width = Buffer.from(png).readUInt32BE(16);
      const pixels = pngPixels(png);
      const pixel = (x: number) => {
        const offset = (15 * scale * width + x * scale) * 4;
        return [...pixels.subarray(offset, offset + 4)];
      };
      for (const x of [15, 35, 55, 75, 95, 135]) expect(pixel(x)).toEqual([0, 0, 255, 255]);
      expect(pixel(115)).toEqual([255, 255, 255, 255]);
    }
  } finally {
    session.free();
  }
});

function docxTextRaster(family: string, properties = '') {
  const raw = require('./index.js') as typeof import('./index.js');
  const parts = wasm.unzip_docx(fixture) as Record<string, Uint8Array>;
  parts['word/document.xml'] = new TextEncoder().encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="24"/>${properties}</w:rPr><w:t>TTTT</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const session = new raw.EditSession(410);
  try {
    session.open_docx(wasm.rezip_docx(parts), true);
    const face = session.register_measure_font(fontBytes);
    const json = pngDisplayList(session, ['Calibri', 'Liberation Sans', 'Unregistered Raster Face'], face);
    const list = JSON.parse(json) as DisplayList;
    const runs = list.pages[0].primitives.filter((primitive): primitive is GlyphRunPrimitive => primitive.kind === 'glyphRun');
    expect(runs.length).toBeGreaterThan(0);
    for (const part of runs) {
      expect(part.fontId).toBe(face);
      expect(part.size).toBe(16);
    }
    const run = { ...runs[0], text: runs.map((part) => part.text).join(''), glyphs: runs.flatMap((part) => part.glyphs) };
    expect(run.text).toBe('TTTT');
    expect(run.glyphs.length).toBe(4);
    return { session, json, list, run };
  } catch (error) {
    session.free();
    throw error;
  }
}

test.each([1, 2])('native docx PNG preserves the layout fallback face at scale %s', (scale) => {
  const fallback = docxTextRaster('Unregistered Raster Face');
  const direct = docxTextRaster('Liberation Sans');
  try {
    expect(fallback.run.fontId).toBe(direct.run.fontId);
    expect(fallback.run.glyphs).toEqual(direct.run.glyphs);
    expect(fallback.run.size).toBe(direct.run.size);
    const actual = pagePng(fallback.session, fallback.json, 0, scale);
    const expected = pagePng(direct.session, direct.json, 0, scale);
    expect(pngPixels(actual)).toEqual(pngPixels(expected));
    expect(pngPixels(actual).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
  } finally {
    fallback.session.free();
    direct.session.free();
  }
});

test.each([1, 2])('native docx PNG preserves positioned letter spacing at scale %s', (scale) => {
  const spaced = docxTextRaster('Liberation Sans', '<w:spacing w:val="200"/>');
  const plain = docxTextRaster('Liberation Sans');
  try {
    const glyphs = (run: GlyphRunPrimitive) => run.glyphs.map((glyph) => {
      expect(glyph.advance).toBeDefined();
      return { x: glyph.x, advance: glyph.advance! };
    });
    const positioned = glyphs(spaced.run);
    const last = positioned[positioned.length - 1];
    const inked = [...positioned.slice(0, -1), { ...last, advance: last.advance - 200 / 15 }];
    expectPngTracking(
      pagePng(spaced.session, spaced.json, 0, scale),
      pagePng(plain.session, plain.json, 0, scale),
      inked, glyphs(plain.run), 200 / 15, scale,
    );
  } finally {
    spaced.session.free();
    plain.session.free();
  }
});

for (const [decoration, properties] of [
  ['underline', '<w:u w:val="single"/>'],
  ['strike', '<w:strike/>'],
] as const) {
  test.each([1, 2])(`native docx PNG paints display list ${decoration} at scale %s`, (scale) => {
    const decorated = docxTextRaster('Liberation Sans', properties);
    const plain = docxTextRaster('Liberation Sans');
    try {
      expect(decorated.run.glyphs).toEqual(plain.run.glyphs);
      const rules = decorated.list.pages[0].primitives.filter((primitive): primitive is DecorationPrimitive =>
        primitive.kind === 'decoration' && primitive.deco === decoration);
      expect(rules.length).toBeGreaterThan(0);
      expect(plain.list.pages[0].primitives.some((primitive) => primitive.kind === 'decoration')).toBe(false);
      rules.sort((left, right) => left.x - right.x);
      for (let index = 1; index < rules.length; index++) {
        expect(rules[index].y).toBe(rules[0].y);
        expect(rules[index].h).toBe(rules[0].h);
        expect(Math.abs(rules[index].x - rules[index - 1].x - rules[index - 1].w)).toBeLessThan(0.02);
      }
      const end = rules[rules.length - 1];
      const rule = { ...rules[0], w: end.x + end.w - rules[0].x };
      const first = decorated.run.glyphs[0];
      const last = decorated.run.glyphs[decorated.run.glyphs.length - 1];
      expect(Math.abs(rule.x - first.x)).toBeLessThan(0.02);
      expect(Math.abs(rule.w - (last.x + last.advance! - first.x))).toBeLessThan(0.02);
      expectPngRule(
        pagePng(decorated.session, decorated.json, 0, scale),
        pagePng(plain.session, plain.json, 0, scale),
        rule, scale,
      );
    } finally {
      decorated.session.free();
      plain.session.free();
    }
  });
}

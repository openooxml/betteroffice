import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import initWasm, * as wasm from '../../packages/docx/src/wasm/generated/edit/docx_edit.js';

type Engine = Pick<typeof wasm, 'EditSession'>;
type Session = wasm.EditSession;
type ParagraphRead = {
  ok: boolean;
  version: string;
  paragraphs: { story: string; paraId: string; text: string }[];
};

const native = require('./index.js') as Engine;
const request = JSON.stringify({ story: 'body', view: 'accepted' });
let fixture: Uint8Array;

beforeAll(async () => {
  await initWasm({
    module_or_path: readFileSync(new URL('../../packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)),
  });
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
});

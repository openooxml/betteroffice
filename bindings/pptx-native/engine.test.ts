import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import initWasm, * as wasm from '../../packages/pptx/src/wasm/generated/pptx_wasm.js';
import type { SlideDisplayList } from '../../packages/pptx/src/types';
import { expectSavedPartsEqual } from '../native-parity';
import { expectFontFace, fontCases, fontErrorCases } from '../font-fixtures';

type Engine = Pick<typeof wasm,
  | 'PptxDocument'
  | 'PptxRenderer'
  | 'parsePptxJson'
  | 'parsePptxJsonWithoutMedia'
  | 'exportPptxStructuredJson'
  | 'exportPptxMarkdownJson'
  | 'renderPptxMarkdownJson'
  | 'compileSlideJson'
  | 'rendererVersion'
  | 'decodeTiffPng'
  | 'decodeFontFace'
>;

const native = require('./index.js') as Engine & { engineVersion(): string };
const fixture = readFileSync(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url));
const bytes = () => new Uint8Array(fixture);
const normalized = (json: string): unknown => JSON.parse(json.replace(/[0-9a-f]{16}-\d+/g, '<version>'));

type Snapshot = {
  slides: { id: string; shapes: { textStories: { id: string }[] }[] }[];
};

let wasmMemory: WebAssembly.Memory;

beforeAll(async () => {
  ({ memory: wasmMemory } = await initWasm({
    module_or_path: readFileSync(new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)),
  }));
});

function storyId(doc: wasm.PptxDocument): string {
  const snapshot = JSON.parse(doc.snapshotJson()) as Snapshot;
  const story = snapshot.slides.flatMap((slide) => slide.shapes)
    .flatMap((shape) => shape.textStories)[0];
  expect(story).toBeDefined();
  return story.id;
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

test('linked font decoder returns known answers and matching native and wasm errors', () => {
  expect(typeof wasm.decodeFontFace).toBe('function');
  expect(typeof native.decodeFontFace).toBe('function');
  for (const entry of fontCases()) {
    const expected = wasm.decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput);
    const actual = native.decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput);
    expectFontFace(expected, entry);
    expectFontFace(actual, entry);
    expectPlainBytes(expected.font);
    expectPlainBytes(actual.font);
    expect(actual.font).toEqual(expected.font);
    expect(actual.payload).toEqual(expected.payload);
    expect('payload' in expected).toBe(entry.payload !== undefined);
    if (expected.payload && actual.payload) {
      expectPlainBytes(expected.payload);
      expectPlainBytes(actual.payload);
    }
  }
  for (const entry of fontErrorCases()) {
    const expected = caughtValue(() => wasm.decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput));
    const actual = caughtValue(() => native.decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput));
    expect(expected).toBeInstanceOf(Error);
    expect(actual).toBeInstanceOf(Error);
    expect(expected).toMatchObject({ name: 'Error', message: entry.message });
    expect(actual).toMatchObject({ name: 'Error', message: entry.message });
  }
});

function runScript(engine: Engine) {
  const doc = engine.PptxDocument.openCollaborative(bytes(), 7);
  const other = engine.PptxDocument.openCollaborative(bytes(), 8);
  const renderer = new engine.PptxRenderer();
  try {
    expect(doc.clientId).toBe(7);
    const snapshot = doc.snapshotJson();
    const metadata = doc.sessionMetadataJson();
    const content = doc.readContentJson('{}');
    const id = storyId(doc);
    const story = doc.storyJson(JSON.stringify({ storyId: id }));
    const withKeys = renderer.snapshotWithLayoutKeysJson(doc);
    const slides = (JSON.parse(snapshot) as Snapshot).slides;
    const layouts = slides.map((slide, index) => {
      const layout = renderer.layoutSlideJson(doc, index);
      const hit = renderer.hitTestJson(40.5, 40.25);
      const key = renderer.slideLayoutKey(doc, index);
      const active = renderer.setActiveSlide(doc, slide.id, key);
      expect(active).toBe(true);
      const slideHit = renderer.hitTestSlideJson(doc, slide.id, 40.5, 40.25);
      return { layout: normalized(layout), hit: normalized(hit), key, active, slideHit: normalized(slideHit) };
    });
    const anchor = doc.anchorCaretJson(JSON.stringify({ storyId: id, index: 0 }));
    const resolved = doc.resolveCaretAnchorJson(anchor);
    expect(JSON.parse(resolved)).toBe(0);
    doc.startUpdateObservation();
    const edited = doc.insertTextJson(JSON.stringify({ storyId: id, index: 0, text: 'Parity ' }));
    const movedAnchor = doc.resolveCaretAnchorJson(anchor);
    const drained = doc.drainUpdateEvent();
    expect(drained.byteLength).toBeGreaterThan(1);
    expect(drained[0]).toBe(0);
    expect(doc.drainUpdateEvent().byteLength).toBe(0);
    const stateVector = doc.encodeStateVector();
    const update = doc.encodeStateAsUpdate();
    const diff = doc.encodeDiff(new Uint8Array([0]));
    const applied = other.applyUpdateJson(update);
    expect(normalized(other.snapshotJson())).toEqual(normalized(doc.snapshotJson()));
    const structured = doc.exportStructuredJson('{}');
    const markdown = doc.exportMarkdownJson('{}');
    const saved = renderer.saveBytes(doc);
    const undone = doc.undoJson();
    const restoredAnchor = doc.resolveCaretAnchorJson(anchor);
    expect(JSON.parse(restoredAnchor)).toBe(0);
    const undoVector = doc.encodeStateVector();
    const undoUpdate = doc.encodeStateAsUpdate();
    return {
      json: {
        snapshot: normalized(snapshot), metadata: normalized(metadata), content: normalized(content),
        story: normalized(story), withKeys: normalized(withKeys), layouts,
        anchor: normalized(anchor), resolved: normalized(resolved), movedAnchor: normalized(movedAnchor),
        edited: normalized(edited), applied: normalized(applied),
        structured: normalized(structured), markdown: normalized(markdown), undone: normalized(undone),
        restoredAnchor: normalized(restoredAnchor),
      },
      drained, stateVector, update, diff, undoVector, undoUpdate, saved,
    };
  } finally {
    renderer.free();
    other.free();
    doc.free();
  }
}

function runOpens(engine: Engine) {
  const source = bytes();
  const baseline = engine.PptxDocument.openReplayBaseline(source, 7);
  const documents = [baseline];
  try {
    const update = baseline.encodeStateAsUpdate();
    const identity = baseline.peerHydrationJson('[]');
    const peer = engine.PptxDocument.openPeerDeckJson(source, identity);
    documents.push(peer);
    peer.registerPeerFontsJson('[]');
    peer.adoptPeerIdentity();
    expect(peer.documentVersion()).toBe(baseline.documentVersion());
    const updatedPeer = engine.PptxDocument.openPeerDeckJson(source, identity, update);
    documents.push(updatedPeer);
    updatedPeer.registerPeerFontsJson('[]');
    updatedPeer.adoptPeerIdentity();
    const joined = engine.PptxDocument.openCollaborativeFromUpdate(update, 8, source);
    documents.push(joined);
    const bare = engine.PptxDocument.openCollaborativeFromUpdate(update, 9);
    documents.push(bare);
    const replay = engine.PptxDocument.openReplayBaseline(source, 10, update);
    documents.push(replay);
    const defaultReplay = engine.PptxDocument.openReplayBaseline(source);
    documents.push(defaultReplay);
    const nullReplay = engine.PptxDocument.openReplayBaseline(source, null, null);
    documents.push(nullReplay);
    return documents.map((doc) => ({
      snapshot: normalized(doc.snapshotJson()), clientId: doc.clientId,
      mode: doc.undoCaptureMode(), version: normalized(JSON.stringify(doc.documentVersion())),
    }));
  } finally {
    for (const doc of documents) doc.free();
  }
}

describe('pptx native engine parity', () => {
  test('wasm and native run the same synchronous script', async () => {
    const expected = runScript(wasm);
    const actual = runScript(native);
    expect(actual.json).toEqual(expected.json);
    for (const key of ['drained', 'stateVector', 'update', 'diff', 'undoVector', 'undoUpdate'] as const) {
      expectPlainBytes(actual[key]);
      expect(actual[key]).toEqual(expected[key]);
    }
    expectPlainBytes(actual.saved);
    await expectSavedPartsEqual(actual.saved, expected.saved);
  });

  test('static opens preserve arguments, optional inputs and peer identity', () => {
    expect(runOpens(native)).toEqual(runOpens(wasm));
  });

  test('free functions return identical results', () => {
    const run = (engine: Engine) => {
      const parsed = engine.parsePptxJson(bytes());
      const metadata = engine.parsePptxJsonWithoutMedia(bytes());
      const structured = engine.exportPptxStructuredJson(bytes(), '{}');
      const markdown = engine.exportPptxMarkdownJson(bytes(), '{}');
      const rendered = engine.renderPptxMarkdownJson(JSON.stringify(JSON.parse(structured).content), '{}');
      const compiled = engine.compileSlideJson(JSON.stringify({ widthPx: 960, heightPx: 540, shapes: [] }));
      return [parsed, metadata, structured, markdown, rendered, compiled].map(normalized);
    };
    expect(run(native)).toEqual(run(wasm));
    const failure = expectThrownParity(
      () => wasm.decodeTiffPng(new Uint8Array([1, 2, 3])),
      () => native.decodeTiffPng(new Uint8Array([1, 2, 3])),
    );
    expect(typeof failure).toBe('string');
  });

  test('requested font faces match for the deck, each slide and an unknown slide', () => {
    const run = (engine: Engine) => {
      const doc = engine.PptxDocument.openCollaborative(bytes(), 7);
      const renderer = new engine.PptxRenderer();
      try {
        const slides = (JSON.parse(doc.snapshotJson()) as Snapshot).slides;
        const deck = renderer.requestedFontFacesJson(doc);
        expect(JSON.parse(deck).faces.length).toBeGreaterThan(0);
        const perSlide = slides.map((slide) => renderer.requestedFontFacesJson(doc, slide.id));
        return { deck, perSlide };
      } finally {
        renderer.free();
        doc.free();
      }
    };
    expect(run(native)).toEqual(run(wasm));
    const wasmDoc = wasm.PptxDocument.openCollaborative(bytes(), 7);
    const nativeDoc = native.PptxDocument.openCollaborative(bytes(), 7);
    const wasmRenderer = new wasm.PptxRenderer();
    const nativeRenderer = new native.PptxRenderer();
    try {
      expectThrownParity(
        () => wasmRenderer.requestedFontFacesJson(wasmDoc, 'missing-slide'),
        () => nativeRenderer.requestedFontFacesJson(nativeDoc, 'missing-slide'),
      );
    } finally {
      wasmRenderer.free();
      nativeRenderer.free();
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('requested script coverage fields match native and wasm for a synthetic deck', async () => {
    const zip = new JSZip();
    const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
    zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
    zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`);
    zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
    const faces = [
      { family: 'BIZ UD明朝', bold: true, italic: true, text: '日本語かな', scripts: ['cjk-jp'] },
      { family: 'Unknown Indic', bold: true, italic: false, text: 'नमस्ते', scripts: ['devanagari'] },
      { family: 'Calibri', bold: false, italic: false, text: 'Latin 日本語かな', scripts: ['cjk-jp'] },
    ];
    const shapes = faces.map((face, index) => `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="Text ${index}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="${index * 1200000}"/><a:ext cx="6000000" cy="1200000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="${Number(face.bold)}" i="${Number(face.italic)}"><a:latin typeface="${face.family}"/></a:rPr><a:t>${face.text}</a:t></a:r></a:p></p:txBody></p:sp>`).join('');
    zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`);
    const input = await zip.generateAsync({ type: 'uint8array' });
    const run = (engine: Engine) => {
      const doc = engine.PptxDocument.openCollaborative(input.slice(), 7);
      const renderer = new engine.PptxRenderer();
      try {
        const slide = (JSON.parse(doc.snapshotJson()) as Snapshot).slides[0];
        const deck = JSON.parse(renderer.requestedFontFacesJson(doc));
        expect(deck.faces).toEqual(faces.map(({ text: _text, ...face }) => face));
        expect(JSON.parse(renderer.requestedFontFacesJson(doc, slide.id))).toEqual(deck);
        return deck;
      } finally {
        renderer.free();
        doc.free();
      }
    };
    expect(run(native)).toEqual(run(wasm));
  });

  test('browser coverage fallback is opt-in and toggles equally in native and wasm', () => {
    const font = readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const run = (engine: Engine) => {
      const doc = engine.PptxDocument.openCollaborative(bytes(), 7);
      const renderer = new engine.PptxRenderer();
      try {
        const slideId = (JSON.parse(doc.snapshotJson()) as Snapshot).slides[0].id;
        const { shapeId } = JSON.parse(doc.addTextBoxJson(JSON.stringify({ slideId, draft: {
          name: 'Coverage', rect: { x: 100000, y: 100000, width: 3000000, height: 1000000 },
          text: 'A👩‍💻हB', style: { fontFamily: 'Arial', fontSizePt: 18, color: '#000000' },
        } }))) as { shapeId: string };
        const fontId = renderer.registerFont('Arial', false, false, new Uint8Array(font));
        const layout = () => {
          const list = JSON.parse(renderer.layoutSlideJson(doc, 0)) as SlideDisplayList;
          return list.primitives.filter((primitive) => primitive.kind === 'textBox' && primitive.shapeId === shapeId)
            .flatMap((box) => box.kind === 'textBox' ? box.lines.flatMap((line) => line.runs) : []);
        };
        const initial = layout();
        expect(initial.map((piece) => piece.text).join('')).toBe('A👩‍💻हB');
        expect(initial.every((piece) => piece.fontId === fontId)).toBe(true);
        const originalKey = renderer.slideLayoutKey(doc, 0);
        renderer.setBrowserCoverageFallback(false);
        expect(renderer.slideLayoutKey(doc, 0)).toBe(originalKey);
        renderer.setBrowserCoverageFallback(true);
        const enabledKey = renderer.slideLayoutKey(doc, 0);
        expect(enabledKey).not.toBe(originalKey);
        expect(renderer.setActiveSlide(doc, slideId, originalKey)).toBe(false);
        const enabled = layout();
        expect(enabled.some((piece) => piece.fontId === 0xffffffff)).toBe(true);
        expect(enabled.map((piece) => piece.text).join('')).toBe('A👩‍💻हB');
        renderer.setBrowserCoverageFallback(true);
        expect(renderer.slideLayoutKey(doc, 0)).toBe(enabledKey);
        renderer.setBrowserCoverageFallback(false);
        expect(renderer.slideLayoutKey(doc, 0)).not.toBe(enabledKey);
        expect(layout()).toEqual(initial);
        return [initial, enabled].map((pieces) => pieces.map(({ text, start, end, fontId }) => ({ text, start, end, fontId })));
      } finally {
        renderer.free();
        doc.free();
      }
    };
    expect(run(native)).toEqual(run(wasm));
  });

  test('native versions differ between sessions opened from the same bytes and client id', () => {
    const first = native.PptxDocument.openCollaborative(bytes(), 7);
    const second = native.PptxDocument.openCollaborative(bytes(), 7);
    try {
      expect(first.documentVersion()).toMatch(/^[0-9a-f]{16}-\d+$/);
      expect(second.documentVersion()).toMatch(/^[0-9a-f]{16}-\d+$/);
      expect(first.documentVersion()).not.toBe(second.documentVersion());
    } finally {
      first.free();
      second.free();
    }
  });

  test('message, decode and peer errors throw the same values', () => {
    const wasmDoc = wasm.PptxDocument.openCollaborative(bytes(), 7);
    const nativeDoc = native.PptxDocument.openCollaborative(bytes(), 7);
    try {
      expect(typeof expectThrownParity(
        () => wasmDoc.insertTextJson('{'),
        () => nativeDoc.insertTextJson('{'),
      )).toBe('string');
      const decode = expectThrownParity(
        () => wasmDoc.validateEditsJson('{'),
        () => nativeDoc.validateEditsJson('{'),
      );
      expect(decode).toBeInstanceOf(Error);
      expect((decode as Error).name).toBe('PptxEditDecodeError');
      const peer = expectThrownParity(
        () => wasm.PptxDocument.openReplayBaseline(bytes(), 0.5),
        () => native.PptxDocument.openReplayBaseline(bytes(), 0.5),
      );
      expect(peer).toBeInstanceOf(Error);
      expect((peer as Error).name).toBe('PresentationPeerError');
      expect((peer as Error & { code: string }).code).toBe('clientId');
      expect(typeof expectThrownParity(
        () => wasm.PptxDocument.openCollaborative(bytes(), 0),
        () => native.PptxDocument.openCollaborative(bytes(), 0),
      )).toBe('string');
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('one face under many names is held once and a face past the budget throws the same typed error', () => {
    const large = new Uint8Array(readFileSync(new URL('../../packages/fonts/assets/Carlito-BoldItalic.ttf', import.meta.url)));
    const small = readFileSync(new URL('../../crates/docx-edit/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const face = (index: number) => {
      const bytes = new Uint8Array(small.byteLength + 4);
      bytes.set(small);
      new DataView(bytes.buffer).setUint32(small.byteLength, index, true);
      return bytes;
    };
    const renderers = [new wasm.PptxRenderer(), new native.PptxRenderer()];
    try {
      const ids = renderers.map((renderer) => renderer.registerFont('Carlito', true, true, large));
      const before = wasmMemory.buffer.byteLength;
      for (let index = 0; index < 300; index++) {
        ids.push(...renderers.map((renderer) => renderer.registerFont(`Alias ${index}`, index % 2 === 1, false, large)));
      }
      expect(wasmMemory.buffer.byteLength - before).toBeLessThan(2 * large.byteLength);
      for (let index = 0; index < ids.length; index += 2) expect(ids[index + 1]).toBe(ids[index]);
      for (let index = 1; index < 256; index++) {
        for (const renderer of renderers) renderer.registerFont(`Face ${index}`, false, false, face(index));
      }
      const failure = expectThrownParity(
        () => renderers[0].registerFont('Overflow', false, false, face(256)),
        () => renderers[1].registerFont('Overflow', false, false, face(256)),
      );
      expect(failure).toBeInstanceOf(Error);
      expect((failure as Error).name).toBe('PptxResourceLimitError');
      expect((failure as Error).message).toBe('render resource limit exceeded: more than 256 font faces');
      expectThrownParity(
        () => renderers[0].registerFallbackFont('Overflow', false, false, face(256)),
        () => renderers[1].registerFallbackFont('Overflow', false, false, face(256)),
      );
      for (const renderer of renderers) renderer.registerFallbackFont('Overflow', false, false, face(7));
    } finally {
      for (const renderer of renderers) renderer.free();
    }
  });

  test('stale proposal acceptance throws the same string', () => {
    const docs = [wasm.PptxDocument.openCollaborative(bytes(), 7), native.PptxDocument.openCollaborative(bytes(), 7)];
    try {
      const proposals = docs.map((doc) => {
        const id = storyId(doc);
        const proposal = JSON.parse(doc.proposeJson(JSON.stringify({
          agentId: 'parity', edits: [{ type: 'replaceText', storyId: id, start: 0, end: 1, text: 'P' }],
        })));
        doc.insertTextJson(JSON.stringify({ storyId: id, index: 0, text: 'Changed ' }));
        return proposal.id as string;
      });
      const failure = expectThrownParity(
        () => docs[0].acceptProposalJson(JSON.stringify({ id: proposals[0] })),
        () => docs[1].acceptProposalJson(JSON.stringify({ id: proposals[1] })),
      );
      expect(typeof failure).toBe('string');
      expect(JSON.parse(failure as string).code).toBe('staleProposal');
    } finally {
      for (const doc of docs) doc.free();
    }
  });

  test('native byte results have transferable ArrayBuffers and independent copies', () => {
    const doc = native.PptxDocument.openCollaborative(bytes(), 7);
    const renderer = new native.PptxRenderer();
    try {
      doc.startUpdateObservation();
      doc.insertTextJson(JSON.stringify({ storyId: storyId(doc), index: 0, text: 'Bytes ' }));
      const drained = doc.drainUpdateEvent();
      const copied = drained.slice();
      const first = drained[0];
      copied[0] ^= 0xff;
      expect(drained[0]).toBe(first);
      const metadata = JSON.parse(native.parsePptxJsonWithoutMedia(bytes())) as {
        media: { partPath: string }[];
      };
      const media = metadata.media.map((part) => doc.mediaBytes(part.partPath));
      for (const value of [
        drained, doc.drainUpdateEvent(), doc.encodeStateVector(), doc.encodeStateAsUpdate(),
        doc.encodeDiff(new Uint8Array([0])), doc.saveBytes(), renderer.saveBytes(doc), ...media,
      ]) {
        expectPlainBytes(value);
        const expected = value.slice();
        const transferred = structuredClone(value, { transfer: [value.buffer as ArrayBuffer] });
        expect(value.byteLength).toBe(0);
        expect(transferred).toEqual(expected);
        expectPlainBytes(transferred);
      }
      expectPlainBytes(doc.encodeStateAsUpdate());
    } finally {
      renderer.free();
      doc.free();
    }
  });

  test('native free drops state and later calls throw Errors', () => {
    const doc = native.PptxDocument.openCollaborative(bytes(), 7);
    const renderer = new native.PptxRenderer();
    doc.free();
    for (const call of [
      () => doc.snapshotJson(), () => doc.clientId, () => doc.documentVersion(),
      () => doc.canUndo(), () => doc.encodeStateVector(), () => doc.insertTextJson('{}'),
      () => doc.clearUpdateObservation(), () => doc.saveBytes(),
      () => renderer.layoutSlideJson(doc, 0), () => renderer.saveBytes(doc),
      () => renderer.setActiveSlide(doc, 'slide', 'key'),
    ]) {
      const failure = caughtValue(call);
      expect(failure).toBeInstanceOf(Error);
      expect((failure as Error).message).toBe('PptxDocument was freed');
    }
    renderer.free();
    for (const call of [
      () => renderer.hitTestJson(0, 0), () => renderer.layoutSlideJson(doc, 0),
      () => renderer.saveBytes(doc), () => renderer.setActiveSlide(doc, 'slide', 'key'),
      () => renderer.registerFont('Arial', false, false, new Uint8Array()),
      () => renderer.setBrowserCoverageFallback(true),
    ]) {
      const failure = caughtValue(call);
      expect(failure).toBeInstanceOf(Error);
      expect((failure as Error).message).toBe('PptxRenderer was freed');
    }
    doc.free();
    renderer.free();
  });

  test('shared versions match wasm and the binding reports its package version', () => {
    expect(native.PptxDocument.version()).toBe(wasm.PptxDocument.version());
    expect(native.rendererVersion()).toBe(wasm.rendererVersion());
    expect(native.engineVersion()).toBe(require('./package.json').version);
  });
});

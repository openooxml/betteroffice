import { afterAll, afterEach, beforeAll, expect, it, setSystemTime, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { decodeS9Envelope } from '../docx/rustParseFacade';
import type { Document } from '../types/document';
import { createEditSession, preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { DirtyProjectionStories, mergeDocxHostMetadata, saveEditorDocument } from './editorSave';
import { createResidentEngineSession } from './residentEngineSession';
import { saveResidentDocument, type ResidentSaveRecord } from './residentSave';
import { sectionPackage } from './sectionTestPackage';
import { sessionInternals } from './sessionInternals';
import { structuralFixture, W, W14 } from './structuralSaveTestSupport';
import * as editWasm from './wasm/index';
import * as sessionFacade from './yrsSessionFacade';
import { wrapOpenedEditSession } from './yrsSessionFacade';
import { yrsToDocument } from './yrsToDocument';

const owned: Array<{ destroy(): void }> = [];
beforeAll(async () => {
  setSystemTime(new Date('2026-10-02T12:00:00Z'));
  await preloadEditWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
  )));
  await preloadOpcWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')
  )));
});
afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});
afterAll(() => setSystemTime());

function open(source: Uint8Array) {
  const raw = createEditSession(411);
  const hostJson = raw.open_docx(source, true);
  const opened = wrapOpenedEditSession(raw, 411, source, hostJson);
  owned.push(opened.session);
  return { raw, hostJson, ...opened };
}

function fixture(text = 'before table'): Uint8Array<ArrayBuffer> {
  const source = structuralFixture(`<w:document xmlns:w="${W}" xmlns:w14="${W14}">`);
  const parts = new Map(Object.entries(unzipContainer(source)));
  const xml = new TextDecoder().decode(parts.get('word/document.xml'));
  parts.set('word/document.xml', toBytes(xml.replace('before table', text)));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

for (const edit of ['text', 'section', 'no-section-change'] as const) {
  it(`materializes once on the first resident save and never on the second or third ${edit} save`, async () => {
    const source = edit === 'section' ? await sectionPackage() : fixture();
    const opened = open(source);
    const referenceSource = source.slice();
    const reference = open(referenceSource);
    reference.session.materializeDocx = () => decodeS9Envelope(
      reference.raw.materialize_docx()!, referenceSource.buffer as ArrayBuffer
    ).document;
    const materialize = spyOn(opened.raw, 'materialize_docx');
    const track = (session: typeof opened.session) => ({
      materializeDocx: spyOn(session, 'materializeDocx'),
      materializedSource: spyOn(sessionInternals(session), 'materializedSource'),
    });
    const consumers: ReturnType<typeof track>[] = [];
    const wrap = sessionFacade.wrapOpenedEditSession;
    const wrapping = spyOn(sessionFacade, 'wrapOpenedEditSession').mockImplementation((...args) => {
      const result = wrap(...args);
      consumers.push(track(result.session));
      return result;
    });
    const record: ResidentSaveRecord = { full: false };
    const dirty = new DirtyProjectionStories();
    let referenceBase: Document | undefined;
    try {
      for (let save = 0; save < 3; save += 1) {
        for (const session of [opened.session, reference.session]) {
          if (edit === 'section') {
            session.setSectionProperties('final', { margins: { left: 1800 + save * 120 } });
          } else if (edit === 'no-section-change') {
            const cell = { story: 'body', tableIndex: 0, row: 0, column: 0 };
            session.setCellShading({ anchor: cell, head: cell }, ['112233', '445566', '778899'][save]!);
          } else {
            const paraId = session.paragraphs('body')[0]!.paraId;
            session.insertText({ story: 'body', paraId, offset: 0 }, ['first', 'second', 'third'][save]!);
          }
        }
        dirty.add('body');
        const saved = await saveResidentDocument(
          opened.raw, 411, new Uint8Array(source.buffer), opened.hostJson,
          opened.host.document, [], record, dirty
        );
        const count = materialize.mock.calls.length;
        expect(consumers).toHaveLength(save + 1);
        const consumer = consumers[save]!;
        expect(consumer.materializedSource).toHaveBeenCalledTimes(1);
        if (save === 0) expect(consumer.materializeDocx).toHaveBeenCalledTimes(1);
        else expect(consumer.materializeDocx).not.toHaveBeenCalled();
        const base = referenceBase ?? reference.session.materializeDocx()!;
        const projected = yrsToDocument(
          reference.session, mergeDocxHostMetadata(base, reference.host.document),
          { storyIds: new Set(['body']) }
        );
        const expected = await saveEditorDocument(reference.session, projected, []);
        projected.originalBuffer = expected;
        referenceBase = projected;
        expect(new Uint8Array(saved)).toEqual(new Uint8Array(expected));
        if (save === 0) expect(count).toBeLessThanOrEqual(1);
        else expect(count).toBe(0);
        materialize.mockClear();
      }
    } finally {
      wrapping.mockRestore();
      for (const consumer of consumers) {
        consumer.materializeDocx.mockRestore();
        consumer.materializedSource.mockRestore();
      }
      materialize.mockRestore();
    }
  });
}

for (const edit of ['text', 'section', 'no-section-change'] as const) {
  for (const saves of [1, 3] as const) {
    it(`preserves the cached materialized source across ${saves} edited resident ${edit} saves`, async () => {
      const source = edit === 'section' ? await sectionPackage() : fixture();
      const { raw, hostJson, session, host } = open(source);
      const before = structuredClone(sessionInternals(session).materializedSource());
      const cached = sessionInternals(session).materializedSource();
      expect(cached).not.toBeNull();
      const materialize = spyOn(raw, 'materialize_docx');
      const record: ResidentSaveRecord = { full: false };
      const dirty = new DirtyProjectionStories();
      try {
        for (let save = 0; save < saves; save += 1) {
          if (edit === 'section') {
            session.setSectionProperties('final', { margins: { left: 1800 + save * 120 } });
          } else if (edit === 'no-section-change') {
            const cell = { story: 'body', tableIndex: 0, row: 0, column: 0 };
            session.setCellShading({ anchor: cell, head: cell }, ['112233', '445566', '778899'][save]!);
          } else {
            const paraId = session.paragraphs('body')[0]!.paraId;
            session.insertText({ story: 'body', paraId, offset: 0 }, ['first', 'second', 'third'][save]!);
          }
          dirty.add('body');
          await saveResidentDocument(
            raw, 411, new Uint8Array(source.buffer), hostJson, host.document, [], record, dirty
          );
          expect(sessionInternals(session).materializedSource()).toEqual(before);
          expect(sessionInternals(session).materializedSource()).toBe(cached);
          expect(materialize).not.toHaveBeenCalled();
        }
      } finally {
        materialize.mockRestore();
      }
    });
  }
}

it('isolates materialized source documents and invalidates them when the source changes', () => {
  const source = fixture();
  const opened = open(source);
  const materialize = spyOn(opened.raw, 'materialize_docx');
  const materializeDocx = spyOn(opened.session, 'materializeDocx');
  try {
    const internals = sessionInternals(opened.session);
    const separators = internals.sourceNoteSeparators();
    expect(separators).not.toBeNull();
    expect(internals.sourceNoteSeparators()).toEqual(separators);
    expect(materializeDocx).not.toHaveBeenCalled();
    const sourceDocument = internals.materializedSource()!;
    expect(internals.materializedSource()).toBe(sourceDocument);
    expect(sourceDocument.originalBuffer).toBe(source.buffer);
    const first = opened.session.materializeDocx()!;
    const expected = structuredClone(first);
    first.package.document.content.length = 0;
    first.package.media?.clear();
    const second = opened.session.materializeDocx()!;
    expect(second).toEqual(expected);
    expect(second).not.toBe(first);
    expect(second.package).not.toBe(first.package);
    expect(sourceDocument).toEqual(expected);
    expect(first.originalBuffer).toBe(source.buffer);
    expect(second.originalBuffer).toBe(first.originalBuffer);
    expect(materialize).toHaveBeenCalledTimes(1);
    const replacement = fixture('replacement text');
    for (const story of opened.session.storyIds()) opened.session.deleteStory(story);
    opened.session.openDocx(replacement, true);
    const replaced = opened.session.materializeDocx()!;
    expect(internals.materializedSource()).not.toBe(sourceDocument);
    expect(internals.materializedSource()).toEqual(replaced);
    expect(replaced).toEqual(decodeS9Envelope(
      opened.raw.materialize_docx()!, replacement.buffer as ArrayBuffer
    ).document);
    expect(replaced.package.document.content).not.toEqual(expected.package.document.content);
    expect(materialize).toHaveBeenCalledTimes(3);
  } finally {
    materializeDocx.mockRestore();
    materialize.mockRestore();
  }
});

for (const change of ['compareDocx', 'openDocx', 'resident raw open_docx'] as const) {
  it(`re-materializes after ${change} replaces the source between resident saves`, async () => {
    const source = await sectionPackage();
    const create = spyOn(editWasm, 'createEditSession');
    let resident: Awaited<ReturnType<typeof createResidentEngineSession>>;
    let raw: ReturnType<typeof createEditSession>;
    try {
      resident = await createResidentEngineSession(undefined, 411);
      raw = create.mock.results[0]!.value as ReturnType<typeof createEditSession>;
    } finally {
      create.mockRestore();
    }
    owned.push(resident);
    const hostJson = resident.openDocx(source, undefined, undefined, true);
    const opened = wrapOpenedEditSession(raw, 411, source, hostJson);
    const original = decodeS9Envelope(raw.materialize_docx()!, source.buffer as ArrayBuffer).document;
    const paraId = opened.session.paragraphs('body')[0]!.paraId;
    opened.session.insertText({ story: 'body', paraId, offset: 0 }, 'edited ');
    const record: ResidentSaveRecord = { full: false };
    const dirty = new DirtyProjectionStories();
    dirty.add('body');
    await saveResidentDocument(
      raw, 411, new Uint8Array(source.buffer), hostJson, opened.host.document, [], record, dirty
    );
    const parts = new Map(Object.entries(unzipContainer(source)));
    const xml = new TextDecoder().decode(parts.get('word/document.xml'));
    parts.set('word/document.xml', toBytes(xml.replace('First', 'Replacement')));
    const replacement = new Uint8Array(rezipPartsToArrayBuffer(parts));
    const second = wrapOpenedEditSession(raw, 411, source, hostJson);
    if (change === 'compareDocx') {
      for (const story of second.session.storyIds()) second.session.deleteStory(story);
      parts.set('word/document.xml', toBytes(xml.replace('First', 'Revised replacement')));
      const revised = new Uint8Array(rezipPartsToArrayBuffer(parts));
      const compared = JSON.parse(sessionInternals(second.session).compareDocx(
        replacement, revised, JSON.stringify({ author: 'QA', date: '2026-10-02T12:00:00Z' })
      ));
      expect(compared).toMatchObject({ ok: true, noop: false });
    } else if (change === 'openDocx') {
      second.session.openDocx(replacement, false);
    } else {
      resident.openDocx(replacement, undefined, undefined, false);
    }
    const materialize = spyOn(raw, 'materialize_docx');
    const track = (session: typeof opened.session) => spyOn(sessionInternals(session), 'materializedSource');
    const consumers: ReturnType<typeof track>[] = [];
    const wrap = sessionFacade.wrapOpenedEditSession;
    try {
      const saveSource = change === 'compareDocx' ? source : replacement;
      const reference = wrapOpenedEditSession(raw, 411, saveSource, hostJson);
      reference.session.materializeDocx = () => decodeS9Envelope(
        raw.materialize_docx()!, saveSource.buffer as ArrayBuffer
      ).document;
      expect(reference.session.materializeDocx()!.package.document.content).not.toEqual(
        original.package.document.content
      );
      const referenceRecord = { ...record };
      const projected = yrsToDocument(
        reference.session,
        mergeDocxHostMetadata(referenceRecord.base ?? reference.session.materializeDocx()!, opened.host.document),
        { storyIds: new Set(['body']) }
      );
      const expected = await saveEditorDocument(reference.session, projected, [], referenceRecord);
      const wrapping = spyOn(sessionFacade, 'wrapOpenedEditSession').mockImplementation((...args) => {
        const result = wrap(...args);
        consumers.push(track(result.session));
        return result;
      });
      try {
        dirty.add('body');
        const saved = await saveResidentDocument(
          raw, 411, new Uint8Array(saveSource.buffer), hostJson, opened.host.document, [], record, dirty
        );
        expect(materialize.mock.calls.length).toBeGreaterThanOrEqual(1);
        expect(consumers).toHaveLength(1);
        expect(consumers[0]!).toHaveBeenCalledTimes(1);
        expect(consumers[0]!.mock.results[0]!.value).toEqual(decodeS9Envelope(
          raw.materialize_docx()!, saveSource.buffer as ArrayBuffer
        ).document);
        expect(new Uint8Array(saved)).toEqual(new Uint8Array(expected));
      } finally {
        wrapping.mockRestore();
      }
    } finally {
      for (const consumer of consumers) consumer.mockRestore();
      materialize.mockRestore();
    }
  });
}

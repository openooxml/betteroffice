import { expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';

import type { LegacyParagraphV1, ParagraphPatchV1 } from '../docx/paragraphPatch';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import * as rustWriter from '../docx/rustSaveFacade';
import { preloadOpcWasm, unzipContainer } from '../docx/wasm';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import type { DocxParagraphSavePlan } from './paragraphIdentity';
import { captureSessionSave, saveYrsDocx, writeSessionSave } from './saveYrsDocx';
import { saveEditorDocument } from './editorSave';

for (const full of [false, true]) {
  const behavior = full
    ? "logs unspliced paragraphs under today's full editor save policy"
    : 'forwards paragraph patches on session saves';
  it(behavior, async () => {
    const paragraphPatches: ParagraphPatchV1[] = [{
      ordinal: 0,
      sessionKey: 'p1',
      seed: {
        units: [{ kind: 'text', text: 'abc', attrs: {} }], properties: {}, unitSha256: 'seed',
        inherit: { paragraph: {}, runs: [{}], paragraphMark: {} },
      },
      textEdits: [{
        from: 1,
        to: 2,
        insert: [{ text: 'X', attrs: {}, inherit: {} }],
        donor: { kind: 'at', at: 1 },
      }],
      runEdits: [],
      paraDeltas: [],
      paraInherit: {},
      markInherit: {},
      resultSha256: 'result',
    }];
    const legacyParagraphs: LegacyParagraphV1[] = [{
      ordinal: 1,
      code: 'UNSUPPORTED_STRUCTURE',
      detail: 'restyle: paragraph inheritance',
    }];
    const plan: DocxParagraphSavePlan = {
      assignments: [],
      patchedParts: [],
      splicedParts: [{
        part: 'word/header1.xml',
        sha256: 'part',
        paragraphs: [[0, 'p1'], [1, 'p2']],
        changed: [0, 1],
        paragraphPatches,
        legacyParagraphs,
      }],
    };
    const session = {
      paragraphSavePlan: () => plan,
      paragraphIdentities: () => ({ sessionId: 'session', paragraphs: [] }),
      writtenParagraphIds: () => ({}),
      recordSavedParagraphIds: () => [],
      storyIds: () => [],
      storySegments: () => [],
      listComments: () => [],
      version: () => 'v1',
    } as unknown as YrsSession;
    const document = { package: { document: { content: [] } } } as unknown as Document;
    const buffer = new ArrayBuffer(0);
    const writer = spyOn(rustWriter, 'writeDocumentWithRust').mockResolvedValue({
      buffer,
      determinism: { seed: 'seed', now: '2026-10-21T00:00:00Z' },
      warnings: [],
    });
    const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
    try {
      const capture = captureSessionSave(session);
      if (full) {
        await saveEditorDocument(session, { ...document, originalBuffer: buffer }, []);
        expect(writer.mock.calls[0]![5]!.patchedParts).toEqual([]);
        expect(writer.mock.calls[0]![5]!.splicedParts).toBeUndefined();
        expect(diagnostic).toHaveBeenCalledWith(
          'DOCX_MINIMAL_PATCH unspliced paragraphs',
          [{ part: 'word/header1.xml', ordinals: [0, 1] }]
        );
      } else {
        await writeSessionSave(session, document, capture, buffer, {}, () => false, true, () => true);
        const forwarded = writer.mock.calls[0]![5]!.splicedParts![0]!;
        expect(forwarded.paragraphs).toEqual([0, 1]);
        expect(forwarded.paragraphPatches).toBe(paragraphPatches);
        expect(forwarded.legacyParagraphs).toBe(legacyParagraphs);
        expect(diagnostic).not.toHaveBeenCalled();
      }
    } finally {
      diagnostic.mockRestore();
      writer.mockRestore();
    }
  });
}

function bodySource(body: string): Uint8Array {
  const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="doc" Type="${office}/officeDocument" Target="word/document.xml"/>` +
      '</Relationships>'
    )],
    ['word/document.xml', toBytes(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      body + '<w:sectPr/></w:body></w:document>'
    )],
  ])));
}

it('logs replacement counts per source-opened Yrs save alongside a patched neighbour', async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(
    new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)
  )));
  await preloadOpcWasm();
  const source = bodySource(
    '<w:p><w:r><w:t xml:space="preserve">Field </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>' +
    '<w:p><w:r><w:t>Neighbour</w:t></w:r></w:p>'
  );
  const session = await createYrsSession();
  const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
  try {
    session.seedFromDocx(source);
    const [field, neighbour] = session.paragraphs('body');
    for (const paragraph of [field!, neighbour!]) {
      session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
    }
    const part = session.paragraphSavePlan().splicedParts!
      .find(({ part }) => part === 'word/document.xml')!;
    expect(part.paragraphPatches?.map(({ ordinal }) => ordinal)).toEqual([1]);
    expect(part.legacyParagraphs).toEqual([expect.objectContaining({
      ordinal: 0,
      detail: expect.stringContaining('field:'),
    })]);
    for (let save = 0; save < 2; save += 1) {
      diagnostic.mockClear();
      const saved = await saveYrsDocx(session);
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      expect(xml).toContain('Edited Field ');
      expect(xml).toContain('Edited Neighbour');
      expect(diagnostic.mock.calls.filter(([warning]) =>
        typeof warning === 'string' && warning.startsWith('DOCX_MINIMAL_PATCH replace counts')
      )).toEqual([['DOCX_MINIMAL_PATCH replace counts {"field":1}']]);
    }
  } finally {
    diagnostic.mockRestore();
    session.destroy();
  }
});

it('keeps adjacent replacement patches through paired adoption and bootstrap and falls back for swapped or unpaired updates', async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(
    new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)
  )));
  await preloadOpcWasm();
  const source = bodySource(
    '<w:p><w:r w:rsidR="00000001"><w:rPr><w:noProof/></w:rPr><w:t>abc</w:t></w:r>' +
    '<w:r w:rsidR="00000002"><w:rPr><w:bCs/></w:rPr><w:t>def</w:t></w:r>' +
    '<w:r w:rsidR="00000003"><w:t>ghij</w:t></w:r></w:p>'
  );
  const sessions: YrsSession[] = [];
  const create = async (clientId: number) => {
    const session = await createYrsSession({ clientId });
    sessions.push(session);
    return session;
  };
  const fork = async (parent: YrsSession, clientId: number) => {
    const peer = await create(clientId);
    const state = parent.encodeState();
    const sidecar = parent.encodeEditProvenanceForTransfer(state);
    expect(sidecar.length).toBeGreaterThan(0);
    peer.bootstrapPeerWithProvenance(
      state, parent.encodePeerMetadata(), undefined, peer.openDocx(source, false), sidecar
    );
    return peer;
  };
  const plan = (session: YrsSession, useLocalProvenance = true) => {
    const savePlan = useLocalProvenance ? session.paragraphSavePlanWithLocalProvenance() : session.paragraphSavePlan();
    return savePlan.splicedParts!.find(({ part }) => part === 'word/document.xml')!;
  };
  try {
    const live = await create(41_001);
    live.seedFromDocx(source);
    const { paraId } = live.paragraphs('body')[0]!;
    const replace = (session: YrsSession, from: number, to: number, text: string) =>
      session.replaceRange({
        story: 'body', start: { paraId, offset: from }, end: { paraId, offset: to },
      }, text);
    replace(live, 9, 10, 'J');
    const sender = await fork(live, 41_002);
    const other = await fork(live, 41_002);
    const swapped = await fork(live, 41_003);
    const unpaired = await fork(live, 41_004);
    const base = sender.encodeStateVector();
    replace(sender, 3, 6, 'YZ');
    replace(sender, 0, 3, 'X');
    replace(other, 3, 6, 'Z');
    replace(other, 0, 3, 'XY');
    const update = sender.encodeStateAsUpdate(base);
    const otherUpdate = other.encodeStateAsUpdate(base);
    sender.encodeStateAsUpdate(sender.encodeStateVector());
    const sidecar = sender.encodeEditProvenanceForTransfer(update, base);
    const otherSidecar = other.encodeEditProvenanceForTransfer(otherUpdate, base);
    expect(update).not.toEqual(otherUpdate);
    expect(sidecar.length).toBeGreaterThan(0);
    expect(otherSidecar.length).toBeGreaterThan(0);

    const observed: Array<ReturnType<typeof plan>> = [];
    const unsubscribe = live.onUpdate((_update, origin) => {
      if (origin === 'local') observed.push(plan(live));
    });
    live.applyLocalUpdateWithProvenance(update, sidecar);
    unsubscribe();
    swapped.applyLocalUpdateWithProvenance(update, otherSidecar);
    unpaired.applyLocalUpdate(update);
    const exported = await fork(live, 41_005);
    for (const session of [live, exported, swapped, unpaired]) {
      expect(session.paragraphs('body')[0]!.text).toBe('XYZghiJ');
      expect(plan(session).changed).toEqual([0]);
    }
    expect(observed).toEqual([plan(live)]);
    for (const session of [live, exported]) {
      const part = plan(session);
      expect(part.legacyParagraphs).toEqual([]);
      expect(part.paragraphPatches).toHaveLength(1);
      expect(part.paragraphPatches![0]!.textEdits.map(({ from, to, donor, insert }) =>
        [from, to, donor, insert.map(run => run.text).join('')]
      )).toEqual([
        [0, 3, { kind: 'at', at: 0 }, 'X'],
        [3, 6, { kind: 'at', at: 3 }, 'YZ'],
        [9, 10, { kind: 'at', at: 9 }, 'J'],
      ]);
    }
    for (const session of [swapped, unpaired]) {
      expect(plan(session)).toEqual(plan(session, false));
    }
  } finally {
    for (const session of sessions.reverse()) session.destroy();
  }
});

for (const [kind, prefix] of [['footnote', 'fn'], ['endnote', 'en']] as const) {
  it.todo(`keeps untouched ${kind} paragraphs byte-identical after editing the opening paragraph (deferred: the opening note paragraph takes the counted Replace and the notes part is written whole, as on main)`, async () => {
    await preloadEditWasm(new Uint8Array(readFileSync(
      new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)
    )));
    await preloadOpcWasm();
    const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const part = `word/${kind}s.xml`;
    const opening = `<w:p><w:r><w:${kind}Ref/></w:r>` +
      '<w:r><w:t>Opening</w:t></w:r></w:p>';
    const untouched = '<w:p w:rsidR = \'00112233\'><w:r x:keep=\'A&#x26;B\'>' +
      '<w:rPr><w:noProof/></w:rPr><w:t>Same note tail</w:t></w:r></w:p>';
    const sibling = `<w:${kind} w:id='3'><w:p><w:r><w:${kind}Ref/></w:r>` +
      '<w:r x:keep="sibling"><w:rPr><w:noProof/></w:rPr><w:t>Sibling</w:t></w:r>' +
      `</w:p></w:${kind}>`;
    const notes = `<?xml version="1.0" encoding="UTF-8"?>\n` +
      `<w:${kind}s xmlns:w="${word}" xmlns:x='urn:note-metadata'>\n` +
      `<w:${kind} w:type='separator' w:id='-1'>` +
      `<w:p><w:r><w:separator/></w:r></w:p></w:${kind}>\n` +
      `<w:${kind} w:type='continuationSeparator' w:id='0'>` +
      `<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>\n` +
      `<w:${kind} w:id='2'>${opening}\n  ${untouched}</w:${kind}>\n` +
      `${sibling}\n</w:${kind}s>`;
    const source = new Uint8Array(rezipPartsToArrayBuffer(new Map([
      ['[Content_Types].xml', toBytes(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}s+xml"/>` +
        '</Types>'
      )],
      ['_rels/.rels', toBytes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="doc" Type="${office}/officeDocument" Target="word/document.xml"/>` +
        '</Relationships>'
      )],
      ['word/_rels/document.xml.rels', toBytes(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="notes" Type="${office}/${kind}s" Target="${kind}s.xml"/>` +
        '</Relationships>'
      )],
      ['word/document.xml', toBytes(
        `<w:document xmlns:w="${word}"><w:body><w:p>` +
        `<w:r><w:${kind}Reference w:id="2"/></w:r>` +
        `<w:r><w:${kind}Reference w:id="3"/></w:r>` +
        '</w:p><w:sectPr/></w:body></w:document>'
      )],
      [part, toBytes(notes)],
    ])));
    const session = await createYrsSession();
    try {
      session.seedFromDocx(source);
      const story = `${prefix}:2`;
      const [paragraph] = session.paragraphs(story);
      session.insertText({ story, paraId: paragraph!.paraId, offset: 0 }, 'Edited ');
      for (let save = 0; save < 2; save += 1) {
        const saved = await saveYrsDocx(session);
        const xml = new TextDecoder().decode(unzipContainer(saved.bytes)[part]);
        expect(xml).toContain(untouched);
        expect(xml).toContain(sibling);
        expect(xml).toBe(notes.replace('>Opening<', '>Edited Opening<'));
      }
    } finally {
      session.destroy();
    }
  });
}

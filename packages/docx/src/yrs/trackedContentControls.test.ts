import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { applyFrameDeltaOwned, decodeFrameDelta } from '../layout/render/frameDelta';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';
import { saveYrsDocx } from './saveYrsDocx';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const INFO = { id: 5, author: 'A', date: '2024-01-01T00:00:00Z' };
const PROPERTIES = { tag: 'tracked', alias: 'Tracked control', id: 42 };
const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
let clientId = 86500;

function fixture(tag: string, outerControl: boolean, mixed = false): Uint8Array {
  const deleted = tag === 'del' || tag === 'moveFrom';
  const text = outerControl ? 'Inserted' : 'Control';
  const run = (text: string, deleted = false) => `<w:r><w:${deleted ? 'delText' : 't'}>${text}</w:${deleted ? 'delText' : 't'}></w:r>`;
  const control = (content: string) => `<w:sdt><w:sdtPr><w:tag w:val="tracked"/><w:alias w:val="Tracked control"/><w:id w:val="42"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const tracked = (content: string) => `<w:${tag} w:id="5" w:author="A" w:date="${INFO.date}">${content}</w:${tag}>`;
  const content = outerControl
    ? control(`${mixed ? run('A') : ''}${tracked(run(text, deleted))}${mixed ? run('C') : ''}`)
    : tracked(control(run(text, deleted)));
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001">${content}<w:r><w:t>tail</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Other</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

type TextSite = { text: string; revision?: unknown; kind?: string; control?: unknown };

function sites(value: unknown, revision?: unknown, kind?: string, control?: unknown): TextSite[] {
  if (Array.isArray(value)) return value.flatMap((child) => sites(child, revision, kind, control));
  if (!value || typeof value !== 'object') return [];
  const node = value as Record<string, unknown>;
  if (['insertion', 'deletion', 'moveFrom', 'moveTo'].includes(String(node.type))) {
    revision = node.info;
    kind = String(node.type);
  }
  if (node.type === 'inlineSdt') control = node.properties;
  if (node.type === 'text') return [{ text: String(node.text), revision, kind, control }];
  return sites(node.content ?? node.structuredChildren ?? node.children ?? node.fieldResult, revision, kind, control);
}

function assertModel(document: Document, text: string, kind: string, mixed = false): void {
  const all = sites(document.package.document.content);
  expect(all.find((site) => site.text === text)).toMatchObject({
    text, revision: INFO, kind, control: PROPERTIES,
  });
  if (mixed) {
    expect(all.filter((site) => site.control).map((site) => site.text).join('')).toBe(`A${text}C`);
    expect(all.find((site) => site.text === 'A')?.revision).toBeUndefined();
    expect(all.find((site) => site.text === 'C')?.revision).toBeUndefined();
  }
}

function project(session: YrsSession): Document {
  const base = session.materializeDocx();
  if (!base) throw new Error('Opened document must materialize');
  return yrsToDocument(session, base);
}

describe('tracked changes and inline controls', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  for (const [tag, kind] of [['ins', 'insertion'], ['del', 'deletion'], ['moveFrom', 'moveFrom'], ['moveTo', 'moveTo']] as const) {
    for (const outerControl of [true, false]) {
      const nesting = outerControl ? 'revision inside control' : 'control inside revision';
      const text = outerControl ? 'Inserted' : 'Control';
      const bytes = () => fixture(tag, outerControl);

      it(`${tag}: ${nesting} parses and repacks`, async () => {
        const parsed = await parseDocx(bytes().buffer as ArrayBuffer, { preloadFonts: false });
        assertModel(parsed, text, kind);
        assertModel(await parseDocx(await repackDocx(parsed), { preloadFonts: false }), text, kind);
      });

      for (const seeder of ['Rust', 'TS'] as const) {
        it(`${tag}: ${nesting} opens with revision and control through ${seeder} seed`, async () => {
          const session = await createYrsSession({ clientId: ++clientId });
          try {
            session.openDocx(bytes(), seeder === 'Rust');
            if (seeder === 'TS') documentToYrs(session, await parseDocx(bytes().buffer as ArrayBuffer, { preloadFonts: false }));
            assertModel(project(session), text, kind);
            expect(session.paragraphs('body')[0]!.text).toContain(text);
            expect(session.listRevisions()).toEqual(expect.arrayContaining([
              expect.objectContaining({
                revisionId: '5', author: 'A', date: INFO.date, preview: text,
                kind: tag === 'ins' || tag === 'moveTo' ? 'insertion' : 'deletion',
              }),
            ]));
            expect(session.searchText(text)).toHaveLength(1);
          } finally {
            session.destroy();
          }
        });
      }

      for (const paraId of ['00000001', '00000002']) {
        it(`${tag}: ${nesting} survives save after edit in ${paraId}`, async () => {
          const session = await createYrsSession({ clientId: ++clientId });
          const reopened = await createYrsSession({ clientId: ++clientId });
          try {
            session.openDocx(bytes(), true);
            expect(session.applyEdits({
              expectVersion: session.version(),
              steps: [{ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId }, at: 'end', text: '!' }],
            }).ok).toBe(true);
            const saved = (await saveYrsDocx(session)).bytes;
            const parsed = await parseDocx(saved, { preloadFonts: false });
            assertModel(parsed, text, kind);
            expect(sites(parsed.package.document.content).map((site) => site.text).join(''))
              .toContain(paraId === '00000001' ? 'tail!' : 'Other!');
            reopened.openDocx(new Uint8Array(saved), true);
            assertModel(project(reopened), text, kind);
          } finally {
            reopened.destroy();
            session.destroy();
          }
        });
      }

      for (const decision of ['accept', 'reject'] as const) {
        it(`${tag}: ${nesting} ${decision}s text inside the control`, async () => {
          const session = await createYrsSession({ clientId: ++clientId });
          try {
            session.openDocx(bytes(), true);
            if (decision === 'accept') session.acceptChange({ revisionId: '5' });
            else session.rejectChange({ revisionId: '5' });
            const retained = (tag === 'ins' || tag === 'moveTo') === (decision === 'accept');
            const all = sites(project(session).package.document.content);
            expect(all.some((site) => site.text === text)).toBe(retained);
            if (retained) expect(all.find((site) => site.text === text)).toMatchObject({ control: PROPERTIES, revision: undefined });
            expect(session.listRevisions()).toHaveLength(0);
          } finally {
            session.destroy();
          }
        });
      }

      it(`${tag}: ${nesting} paints text with revision status`, async () => {
        const session = await createYrsSession({ clientId: ++clientId });
        try {
          session.openDocx(bytes(), true);
          session.registerFont(new Uint8Array(readFileSync(FONT)));
          session.layoutDocumentWithRegionsJson(JSON.stringify({
            bodyStory: 'body',
            regions: { sections: [{ sectionId: 'main', properties: {} }] },
            measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
            renderEnv: {},
          }));
          session.setSelection({ story: 'body', paraId: '00000001', offset: 0 });
          const frame = applyFrameDeltaOwned(null, decodeFrameDelta(session.buildDisplayListFrame('{}', 0)));
          const primitives = frame.displayList.pages.flatMap((page) => page.primitives)
            .filter((primitive) => primitive.kind === 'text' || primitive.kind === 'glyphRun');
          expect(primitives.map((primitive) => primitive.text).join('')).toContain(text);
          const revisionKind = tag === 'ins' || tag === 'moveTo' ? 'ins' : 'del';
          expect(primitives.filter((primitive) => primitive.revision?.kind === revisionKind)
            .map((primitive) => primitive.text).join('')).toContain(text);
        } finally {
          session.destroy();
        }
      });
    }
  }

  it('mixed sdtContent keeps surrounding runs and the inner revision through save', async () => {
    const session = await createYrsSession({ clientId: ++clientId });
    try {
      const bytes = fixture('ins', true, true);
      const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
      assertModel(parsed, 'Inserted', 'insertion', true);
      assertModel(await parseDocx(await repackDocx(parsed), { preloadFonts: false }), 'Inserted', 'insertion', true);
      session.openDocx(bytes, true);
      assertModel(project(session), 'Inserted', 'insertion', true);
      session.insertText({ story: 'body', paraId: '00000002', offset: 5 }, '!');
      assertModel(await parseDocx((await saveYrsDocx(session)).bytes, { preloadFonts: false }), 'Inserted', 'insertion', true);
    } finally {
      session.destroy();
    }
  });
});

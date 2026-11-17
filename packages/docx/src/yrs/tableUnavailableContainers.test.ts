import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession, saveYrsDocx, writeSessionSave, type YrsCellLoc, type YrsSession } from './index';
import { captureSessionSave } from './saveYrsDocx';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const table = (...cells: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cells.map(() => '<w:gridCol w:w="4680"/>').join('')}</w:tblGrid><w:tr>${cells.join('')}</w:tr></w:tbl>`;
const plain = (name: string) => table(cell(paragraph(`${name} one`)), cell(paragraph(`${name} two`)));
const control = `<w:sdt><w:sdtPr><w:alias w:val="Box"/></w:sdtPr><w:sdtContent>${plain('Control')}${paragraph('In control')}</w:sdtContent></w:sdt>`;

/** A table in the body, in a block content control inside a body table cell, in a header and in a footnote. */
function containers(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE_DOC}.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="${OFFICE_DOC}.wordprocessingml.footnotes+xml"/></Types>`));
  parts.set('_rels/.rels', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/_rels/document.xml.rels', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="top" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdF1" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`));
  parts.set('word/header1.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${W}>${plain('Header')}${paragraph('Header text')}</w:hdr>`));
  parts.set('word/footnotes.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${plain('Note')}${paragraph('Note text')}</w:footnote></w:footnotes>`));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body><w:p><w:r><w:t>Before</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>${table(cell(`${control}${paragraph('In cell')}`), cell(paragraph('Body two')))}${paragraph('After')}<w:sectPr><w:headerReference w:type="default" r:id="top"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const owned: YrsSession[] = [];

async function session(clientId: number): Promise<YrsSession> {
  const created = await createYrsSession({ clientId });
  owned.push(created);
  return created;
}

function exchange(a: YrsSession, b: YrsSession): void {
  const fromA = a.encodeState();
  const fromB = b.encodeState();
  b.applyUpdate(fromA);
  a.applyUpdate(fromB);
}

function tableStories(target: YrsSession): string[] {
  return target.storyIds().filter((story) =>
    target.storySegments(story).some((segment) => segment.kind === 'embed' && segment.embedKind === 'table'));
}

/** Crossing undos of two collaborators' changes to the first table of `story`. */
async function emptied(source: Uint8Array, story: string) {
  const a = await session(61);
  const b = await session(62);
  a.seedFromDocx(source);
  b.openDocx(source, false);
  b.applyUpdate(a.encodeState());
  const at: YrsCellLoc = { story, tableIndex: 0, row: 0, column: 0 };
  a.insertRow(at, 'below');
  exchange(a, b);
  b.setColumnWidth(at, 2000);
  expect(a.undo()).toBe(true);
  exchange(a, b);
  expect(b.undo()).toBe(true);
  exchange(a, b);
  const payload = a.tablePayload(story, 0) ?? {};
  expect(['tblPr', 'grid', 'rows'].filter((key) => key in payload)).toEqual([]);
  return { a, b };
}

async function refusal(call: () => Promise<unknown>): Promise<Error & { code?: string }> {
  try {
    await call();
  } catch (error) {
    return error as Error & { code?: string };
  }
  throw new Error('expected a refusal');
}

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  await preloadOpcWasm();
}, 60_000);

afterEach(() => {
  for (const created of owned.splice(0)) created.destroy();
});

describe('an unavailable table in any container', () => {
  it('the fixture holds a table in the body, a cell content control, a header and a footnote', async () => {
    const probe = await session(60);
    probe.seedFromDocx(containers());
    const stories = tableStories(probe);
    expect(stories).toContain('body');
    expect(stories.some((story) => story.startsWith('body:t0:') && story.includes(':sdt'))).toBe(true);
    expect(stories).toContain('hf:top');
    expect(stories.some((story) => story.startsWith('fn:'))).toBe(true);
    expect((await saveYrsDocx(probe, { updateModifiedDate: false })).bytes.byteLength).toBeGreaterThan(0);
  });

  it('is found and refuses the lowest writer, whatever story holds it', async () => {
    const source = containers();
    const probe = await session(60);
    probe.seedFromDocx(source);
    const stories = tableStories(probe);
    expect(stories.length).toBeGreaterThanOrEqual(4);
    for (const story of stories) {
      const { a, b } = await emptied(source, story);
      for (const target of [a, b]) {
        const before = target.encodeState();
        const base = target.materializeDocx()!;
        const writer = await refusal(() =>
          writeSessionSave(target, base, captureSessionSave(target), base.originalBuffer!, {}, () => false, false, () => false));
        expect(writer.code, story).toBe('table-unavailable');
        expect((await refusal(() => saveYrsDocx(target, { updateModifiedDate: false }))).code, story).toBe('table-unavailable');
        const elsewhere = new Set(['body', 'hf:top'].filter((root) => root !== story));
        expect((await refusal(async () => yrsToDocument(target, base, { storyIds: elsewhere }))).code, story)
          .toBe('table-unavailable');
        expect(target.encodeState()).toEqual(before);
      }
    }
  });

  it('a file table without w:tblPr or w:tblGrid seeds all three payload keys and stays available', async () => {
    const parts: PartsMap = new Map();
    parts.set('[Content_Types].xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`));
    parts.set('_rels/.rels', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`));
    parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${paragraph('Before')}<w:tbl><w:tr>${cell(paragraph('Bare one'))}${cell(paragraph('Bare two'))}</w:tr></w:tbl>${paragraph('After')}<w:sectPr/></w:body></w:document>`));
    const bare = await session(63);
    bare.seedFromDocx(new Uint8Array(rezipPartsToArrayBuffer(parts)));
    const payload = bare.tablePayload('body', 0)!;
    expect(['tblPr', 'grid', 'rows'].filter((key) => key in payload)).toEqual(['tblPr', 'grid', 'rows']);
    const saved = await saveYrsDocx(bare, { updateModifiedDate: false });
    expect(saved.bytes.byteLength).toBeGreaterThan(0);
    bare.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'below');
  });

  it('saves again once the unavailable table holding a cell content control is deleted', async () => {
    const source = containers();
    const { a } = await emptied(source, 'body');
    expect((await refusal(() => saveYrsDocx(a, { updateModifiedDate: false }))).code).toBe('table-unavailable');
    a.deleteTable({ story: 'body', tableIndex: 0 });
    const saved = await saveYrsDocx(a, { updateModifiedDate: false });
    expect(saved.bytes.byteLength).toBeGreaterThan(0);
  });
});

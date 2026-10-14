import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import {
  DirtyProjectionStories,
  hostSaveMetadata,
  mergeDocxHostMetadata,
  saveEditorDocument,
} from './editorSave';
import type { DocxParagraphText } from './edits';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, type YrsSession } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const TIMEOUT = 120_000;

const STYLES =
  `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  '<w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans" w:eastAsia="DejaVu Sans" w:cs="DejaVu Sans"/>' +
  '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>' +
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>' +
  '<w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>' +
  '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
  '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>';

const cell = (text: string) =>
  '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
  `<w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Sumary</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">The pilot started in September.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">Feedback so far is positive.</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/>' +
  '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
  `<w:tr>${cell('Rent')}${cell('1200')}</w:tr><w:tr>${cell('Travel')}${cell('300')}</w:tr></w:tbl>` +
  '<w:p><w:r><w:t xml:space="preserve">Collect the remaining forms.</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

type Root = 'plain' | 'w14';

function fixture(root: Root): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="${OFFICE}.wordprocessingml.styles+xml"/></Types>`
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rIdDoc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  ));
  parts.set('word/_rels/document.xml.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
  ));
  parts.set('word/styles.xml', toBytes(STYLES));
  const declarations = root === 'w14' ? ` xmlns:w14="${W14}"` : '';
  parts.set('word/document.xml', toBytes(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"${declarations}><w:body>${BODY}</w:body></w:document>`
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

type EditEngine = Pick<
  YrsSession,
  'applyEdits' | 'readParagraphs' | 'version' | 'replaceRange' | 'formatRange' | 'insertRow'
>;

type Scenario = 'heading' | 'unbold' | 'style' | 'bold' | 'row' | 'paragraphs';

const NEW_PARAGRAPH = 'The steering group meets on 20 October.';

function bodyParagraph(engine: EditEngine, prefix: string): DocxParagraphText {
  const read = engine.readParagraphs({ story: 'body', view: 'accepted' });
  if (!read.ok) throw new Error(`read refused: ${JSON.stringify(read)}`);
  const found = read.paragraphs.find((paragraph) => paragraph.text.startsWith(prefix));
  if (!found) throw new Error(`no body paragraph starts with ${prefix}`);
  return found;
}

/** Applies `scenario` and returns the stories it changed. */
function edit(engine: EditEngine, scenario: Scenario): string[] {
  const at = (paraId: string, start: number, end: number) => ({
    story: 'body', start: { paraId, offset: start }, end: { paraId, offset: end },
  });
  switch (scenario) {
    case 'heading': {
      const { paraId } = bodyParagraph(engine, 'Sumary');
      engine.replaceRange(at(paraId, 0, 6), 'Summary');
      return ['body'];
    }
    case 'unbold': {
      const { paraId } = bodyParagraph(engine, 'Sumary');
      engine.formatRange(at(paraId, 0, 3), { bold: false });
      return ['body'];
    }
    case 'bold': {
      const { paraId } = bodyParagraph(engine, 'The pilot');
      engine.formatRange(at(paraId, 4, 9), { bold: true });
      return ['body'];
    }
    case 'row': {
      const receipt = engine.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'below');
      return ['body', ...receipt.createdStoryIds, ...receipt.changedStoryIds];
    }
    case 'style':
    case 'paragraphs': {
      const { paraId } = bodyParagraph(engine, scenario === 'style' ? 'The pilot' : 'Feedback');
      const target = { story: 'body', paraId };
      const result = engine.applyEdits({
        expectVersion: engine.version(),
        steps: [scenario === 'style'
          ? { op: 'setParagraphStyle', target, styleId: 'Heading1' }
          : { op: 'insertParagraphs', target, at: 'end', paragraphs: [{ text: NEW_PARAGRAPH }] }],
      });
      if (!result.ok) throw new Error(`edit refused: ${JSON.stringify(result)}`);
      return ['body'];
    }
  }
}

const owned: Array<{ destroy(): void }> = [];
afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

let startWorker: (clientId?: number) => InProcessResidentWorker;
beforeAll(async () => {
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  await preloadOpcWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')))
  );
  startWorker = await residentWorkerFactory();
});

/** The agents' engine calls: seed, edit, fork from the source plus state, save the fork. */
async function agentsSave(bytes: Uint8Array, scenario: Scenario): Promise<Uint8Array> {
  const session = await createYrsSession();
  owned.push(session);
  session.seedFromDocx(bytes);
  edit(session, scenario);
  const draft = await createYrsSession();
  owned.push(draft);
  draft.openDocx(new Uint8Array(session.materializeDocx()!.originalBuffer!), false);
  draft.loadState(session.encodeState());
  return (await saveYrsDocx(draft)).bytes;
}

async function openWorker(bytes: Uint8Array) {
  const worker = startWorker(311);
  const client = new ResidentEngineWorkerClient(worker);
  owned.push(client);
  const { hostJson } = await client.open(bytes);
  return { client, resident: worker.sessions[0]!, host: decodeDocxHostJson(hostJson, bytes).document };
}

/** The worker-only editor: the resident session edits and saves inside the worker. */
async function workerSave(bytes: Uint8Array, scenario: Scenario): Promise<Uint8Array> {
  const { client, resident, host } = await openWorker(bytes);
  const { replaceRange, formatRange, insertRow } = resident;
  edit({ ...resident.proposalEngine, replaceRange, formatRange, insertRow }, scenario);
  const saved = await client.save({
    comments: host.package.document.comments ?? [],
    host: hostSaveMetadata(host),
  });
  return new Uint8Array(saved.bytes);
}

/** The main-thread engine save the hybrid peer and the in-thread editor share. */
async function mainThreadSave(bytes: Uint8Array, scenario: Scenario): Promise<Uint8Array> {
  const session = await createYrsSession({ clientId: 312 });
  owned.push(session);
  const host = session.openDocx(bytes.slice(), true).document;
  const dirty = new DirtyProjectionStories();
  for (const story of edit(session, scenario)) dirty.add(story);
  const base = session.materializeDocx() as Document;
  const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host), dirty.projectionOptions());
  return new Uint8Array(await saveEditorDocument(session, projected, host.package.document.comments ?? []));
}

const SAVES = {
  'agents engine': agentsSave,
  'worker-only': workerSave,
  'main-thread (hybrid and in-thread)': mainThreadSave,
} as const;

function documentXml(bytes: Uint8Array): string {
  return new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
}

function paragraphs(xml: string): string[] {
  return xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
}

function rootAttributes(xml: string): Map<string, string> {
  const tag = xml.match(/<w:document\b[^>]*>/)?.[0] ?? '';
  return new Map([...tag.matchAll(/\s([\w:]+)="([^"]*)"/g)].map((match) => [match[1]!, match[2]!]));
}

/** The root keeps its declarations and gains at most what written paragraph IDs need. */
function expectRootKept(xml: string, sourceXml: string): void {
  const source = rootAttributes(sourceXml);
  const written = rootAttributes(xml);
  for (const [name, value] of source) expect(written.get(name)).toBe(value);
  const added = [...written.keys()].filter((name) => !source.has(name));
  expect(added.filter((name) => !['xmlns:w14', 'xmlns:mc', 'mc:Ignorable'].includes(name))).toEqual([]);
  expect(['', 'w14']).toContain(written.get('mc:Ignorable') ?? '');
}

const TOUCHED: Record<Scenario, string | null> = {
  heading: 'Sumary',
  unbold: 'Sumary',
  style: 'The pilot',
  bold: 'The pilot',
  row: null,
  paragraphs: null,
};

const text = (paragraph: string) => paragraph.replace(/<[^>]+>/g, '');

describe.each(Object.keys(SAVES) as (keyof typeof SAVES)[])('%s save keeps inherited formatting inherited', (path) => {
  describe.each(['plain', 'w14'] as Root[])('with a %s document root', (root) => {
    const source = fixture(root);
    const sourceXml = documentXml(source);

    async function save(scenario: Scenario) {
      const xml = documentXml(await SAVES[path](source, scenario));
      const touched = TOUCHED[scenario];
      const kept = paragraphs(xml);
      for (const paragraph of paragraphs(sourceXml)) {
        if (touched !== null && text(paragraph).startsWith(touched)) continue;
        expect(kept).toContain(paragraph);
      }
      expectRootKept(xml, sourceXml);
      return { xml, kept };
    }

    it('writes a heading text edit without run properties', async () => {
      const { kept } = await save('heading');
      const heading = kept.find((paragraph) => text(paragraph) === 'Summary');
      expect(heading).toMatch(/^<w:p\b[^>]*><w:pPr><w:pStyle w:val="Heading1"\/><\/w:pPr><w:r><w:t\b[^>]*>Summary<\/w:t><\/w:r><\/w:p>$/);
    }, TIMEOUT);

    it('writes un-bolding part of a heading as an explicit bold override', async () => {
      const { kept } = await save('unbold');
      const heading = kept.find((paragraph) => text(paragraph) === 'Sumary')!;
      const runs = heading.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
      expect(runs).toHaveLength(2);
      expect(runs[0]).toMatch(/^<w:r><w:rPr><w:b w:val="(?:0|false)"\/>(?:<w:bCs w:val="(?:0|false)"\/>)?<\/w:rPr><w:t\b[^>]*>Sum<\/w:t><\/w:r>$/);
      expect(runs[1]).toMatch(/^<w:r><w:t\b[^>]*>ary<\/w:t><\/w:r>$/);
    }, TIMEOUT);

    it('writes a style change as the style reference only', async () => {
      const { kept } = await save('style');
      const styled = kept.find((paragraph) => text(paragraph).startsWith('The pilot'));
      expect(styled).toMatch(/^<w:p\b[^>]*><w:pPr><w:pStyle w:val="Heading1"\/><\/w:pPr><w:r><w:t\b[^>]*>The pilot started in September\.<\/w:t><\/w:r><\/w:p>$/);
    }, TIMEOUT);

    it('writes bold on a phrase as bold only', async () => {
      const { kept } = await save('bold');
      const formatted = kept.find((paragraph) => text(paragraph).startsWith('The pilot'))!;
      expect(formatted).not.toContain('<w:rFonts');
      expect(formatted).not.toContain('<w:sz ');
      const properties = formatted.match(/<w:rPr>[\s\S]*?<\/w:rPr>/g) ?? [];
      expect(properties).toHaveLength(1);
      expect(properties[0]).toMatch(/^<w:rPr><w:b\/>(?:<w:bCs\/>)?<\/w:rPr>$/);
    }, TIMEOUT);

    it('splices a table row insert without rewriting other paragraphs', async () => {
      const { xml, kept } = await save('row');
      expect(kept).toHaveLength(paragraphs(sourceXml).length + 2);
      for (const inherited of ['<w:rFonts', '<w:spacing', '<w:tblBorders', '<w:tblCellMar', '<w:tcBorders', '<w:tcMar']) {
        expect(xml).not.toContain(inherited);
      }
    }, TIMEOUT);

    it('splices inserted paragraphs without rewriting other paragraphs', async () => {
      const { kept } = await save('paragraphs');
      const inserted = kept.find((paragraph) => text(paragraph) === NEW_PARAGRAPH);
      expect(inserted).toBeDefined();
      expect(inserted).not.toContain('<w:rPr>');
      expect(inserted).not.toContain('<w:spacing');
    }, TIMEOUT);
  });
});

type TableEditEngine = EditEngine & Pick<YrsSession, 'storyIds'> & {
  setCellBorders(
    range: Parameters<YrsSession['setCellBorders']>[0],
    borders: Parameters<YrsSession['setCellBorders']>[1]
  ): { changedStoryIds: string[] };
  setCellTextFormat(
    range: Parameters<YrsSession['setCellTextFormat']>[0],
    patch: Parameters<YrsSession['setCellTextFormat']>[1]
  ): { changedStoryIds: string[] };
};
type TableEdit = (engine: TableEditEngine) => string[];

async function saveTableEdit(path: keyof typeof SAVES, source: Uint8Array, apply: TableEdit) {
  if (path === 'worker-only') {
    const { client, resident, host } = await openWorker(source);
    const patchCell = (patch: Record<string, unknown>) => {
      let index = 0;
      for (const segment of resident.storySegments('body')) {
        if (segment.kind === 'embed' && segment.embedKind === 'table') {
          const rows = segment.payload.rows as Array<{ cells: Array<{ tcPr: Record<string, unknown> }> }>;
          Object.assign(rows[0]!.cells[0]!.tcPr, patch);
          resident.applyRawOps('body', [{ op: 'setEmbedAttr', index, key: 'rows', value: rows }]);
          return { changedStoryIds: ['body'] };
        }
        index += segment.kind === 'text' ? segment.text.length : 1;
      }
      throw new Error('no resident table to patch');
    };
    apply({
      ...resident.proposalEngine,
      storyIds: resident.geometryReader.storyIds,
      replaceRange: resident.replaceRange,
      formatRange: resident.formatRange,
      insertRow: resident.insertRow,
      setCellBorders: (_range, borders) => {
        const table = resident.tablePayload('body', 0)!;
        const rows = table.rows as Array<{ cells: Array<{ tcPr: Record<string, unknown> }> }>;
        return patchCell({ borders: { ...(rows[0]!.cells[0]!.tcPr.borders as object), ...borders } });
      },
      setCellTextFormat: (_range, patch) => patchCell({ ...patch }),
    });
    const saved = await client.save({
      comments: host.package.document.comments ?? [],
      host: hostSaveMetadata(host),
    });
    return documentXml(new Uint8Array(saved.bytes));
  }
  const session = await createYrsSession();
  owned.push(session);
  if (path === 'agents engine') {
    session.seedFromDocx(source);
    apply(session);
    const draft = await createYrsSession();
    owned.push(draft);
    draft.openDocx(new Uint8Array(session.materializeDocx()!.originalBuffer!), false);
    draft.loadState(session.encodeState());
    return documentXml((await saveYrsDocx(draft)).bytes);
  }
  const host = session.openDocx(source.slice(), true).document;
  const dirty = new DirtyProjectionStories();
  for (const story of apply(session)) dirty.add(story);
  const base = session.materializeDocx() as Document;
  const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host), dirty.projectionOptions());
  return documentXml(new Uint8Array(
    await saveEditorDocument(session, projected, host.package.document.comments ?? [])
  ));
}

function savedTable(xml: string): string {
  const table = xml.match(/<w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>/)?.[0];
  expect(table).toBeDefined();
  return table!;
}

function expectTableInheritance(table: string): void {
  for (const property of ['tblBorders', 'tblCellMar', 'tcBorders', 'tcMar', 'spacing', 'pStyle', 'jc', 'rFonts']) {
    expect(table).not.toContain(`<w:${property}`);
  }
  expect(table).not.toContain('<w:szCs');
}

function s2Fixture(root: Root): Uint8Array {
  const parts: PartsMap = new Map(Object.entries(unzipContainer(fixture(root))));
  const xml = new TextDecoder().decode(parts.get('word/document.xml'));
  parts.set('word/document.xml', toBytes(xml.replace('<w:sectPr>', `${savedTable(xml)}<w:sectPr>`)));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

describe('cell and table formatting stays inherited on save', () => {
  describe.each(Object.keys(SAVES) as (keyof typeof SAVES)[])('%s', (path) => {
    describe.each(['plain', 'w14'] as Root[])('with a %s document root', (root) => {
      const source = s2Fixture(root);
      const at = { story: 'body', tableIndex: 0, row: 0, column: 0 };
      const range = { anchor: at, head: at };

      it('saves a cell text edit without direct table-style paragraph or cell properties', async () => {
        const xml = await saveTableEdit(path, source, (engine) => {
          const story = engine.storyIds().find((id) => id === 'body:t0:r0c0')!;
          const read = engine.readParagraphs({ story, view: 'accepted' });
          if (!read.ok) throw new Error(`cell read refused: ${JSON.stringify(read)}`);
          const paraId = read.paragraphs[0]!.paraId;
          engine.replaceRange({
            story, start: { paraId, offset: 0 }, end: { paraId, offset: 4 },
          }, 'Lease');
          return [story];
        });
        const table = savedTable(xml);
        expect(text(table)).toContain('Lease');
        expectTableInheritance(table);
        expect(Array.from(table.match(/<w:sz\b[^>]*\/>/g) ?? [])).toEqual(Array(4).fill('<w:sz w:val="20"/>'));
      }, TIMEOUT);

      it('keeps untouched table properties inherited when inserting body paragraphs', async () => {
        const xml = await saveTableEdit(path, source, (engine) => {
          edit(engine, 'paragraphs');
          const receipt = engine.insertRow({ ...at, tableIndex: 1 }, 'below');
          return ['body', ...receipt.createdStoryIds, ...receipt.changedStoryIds];
        });
        expect(text(xml)).toContain(NEW_PARAGRAPH);
        const table = savedTable(xml);
        expectTableInheritance(table);
        expect(Array.from(table.match(/<w:sz\b[^>]*\/>/g) ?? [])).toEqual(Array(4).fill('<w:sz w:val="20"/>'));
      }, TIMEOUT);

      it('writes a changed cell border', async () => {
        const xml = await saveTableEdit(path, source, (engine) => {
          const receipt = engine.setCellBorders(range, {
            top: { style: 'double', size: 12, color: { rgb: 'FF0000' } },
          });
          return ['body', ...receipt.changedStoryIds];
        });
        const properties = savedTable(xml).match(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/)?.[0];
        expect(properties).toContain('<w:tcBorders>');
        expect(properties).toMatch(/<w:top\b[^>]*w:val="double"[^>]*w:sz="12"[^>]*w:color="FF0000"/);
      }, TIMEOUT);

      it('writes a changed cell margin', async () => {
        const xml = await saveTableEdit(path, source, (engine) => {
          const receipt = engine.setCellTextFormat(range, { margins: { left: 240, right: 108 } });
          return ['body', ...receipt.changedStoryIds];
        });
        const properties = savedTable(xml).match(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/)?.[0];
        expect(properties).toContain('<w:tcMar>');
        expect(properties).toMatch(/<w:left\b[^>]*w:w="240"[^>]*w:type="dxa"/);
      }, TIMEOUT);

      it('inserts cells without adding paragraph styles, alignment, fonts or sizes', async () => {
        const xml = await saveTableEdit(path, source, (engine) => edit(engine, 'row'));
        const table = savedTable(xml);
        expectTableInheritance(table);
        const empty = paragraphs(table).filter((paragraph) => text(paragraph) === '');
        expect(empty).toHaveLength(2);
        for (const paragraph of empty) expect(paragraph).not.toContain('<w:rPr');
        expect(Array.from(table.match(/<w:sz\b[^>]*\/>/g) ?? [])).toEqual(Array(4).fill('<w:sz w:val="20"/>'));
      }, TIMEOUT);
    });
  });
});

const TOGGLE_CASES = [
  { name: 'paragraph basedOn b+b', paragraph: true, character: false, table: false, chain: 'paragraph', bold: true },
  { name: 'character basedOn b+b', paragraph: false, character: true, table: false, chain: 'character', bold: true },
  { name: 'paragraph b + character b', paragraph: true, character: true, table: false, chain: '', bold: false },
  { name: 'table b + paragraph b', paragraph: true, character: false, table: true, chain: '', bold: false },
  { name: 'table b + character b', paragraph: false, character: true, table: true, chain: '', bold: false },
  { name: 'table b alone', paragraph: false, character: false, table: true, chain: '', bold: true },
] as const;

function runFixture(styles: string, paragraph: string, inTable = false): Uint8Array {
  const parts: PartsMap = new Map(Object.entries(unzipContainer(fixture('w14'))));
  parts.set('word/styles.xml', toBytes(
    `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>${styles}</w:styles>`
  ));
  const body = inTable
    ? '<w:tbl><w:tblPr><w:tblStyle w:val="BoldTable"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${paragraph}</w:tc></w:tr></w:tbl>`
    : paragraph;
  parts.set('word/document.xml', toBytes(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`
  ));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function editFirstRun(engine: TableEditEngine, story: string, bold: boolean): string[] {
  const read = engine.readParagraphs({ story, view: 'accepted' });
  if (!read.ok) throw new Error(`read refused: ${JSON.stringify(read)}`);
  const paragraph = read.paragraphs[0]!;
  const range = {
    story, start: { paraId: paragraph.paraId, offset: 0 },
    end: { paraId: paragraph.paraId, offset: paragraph.text.length },
  };
  if (bold) engine.formatRange(range, { bold: true });
  else engine.replaceRange(range, 'جديد');
  return [story];
}

describe('toggle inheritance and independent complex-script saves', () => {
  describe.each(Object.keys(SAVES) as (keyof typeof SAVES)[])('%s', (path) => {
    it('saves all six style combinations with and without an explicit bold request', async () => {
      for (const shape of TOGGLE_CASES) {
        for (const request of [false, true]) {
          const style = (type: string, id: string) =>
            `<w:style w:type="${type}" w:styleId="${id}">` +
            (shape.chain === type ? '<w:basedOn w:val="Base"/>' : '') +
            '<w:rPr><w:b/></w:rPr></w:style>';
          const styles =
            (shape.chain ? `<w:style w:type="${shape.chain}" w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style>` : '') +
            (shape.paragraph ? style('paragraph', 'BoldParagraph') : '') +
            (shape.character ? style('character', 'BoldCharacter') : '') +
            (shape.table ? style('table', 'BoldTable') : '');
          const paragraph = '<w:p>' +
            (shape.paragraph ? '<w:pPr><w:pStyle w:val="BoldParagraph"/></w:pPr>' : '') +
            '<w:r><w:rPr>' + (shape.character ? '<w:rStyle w:val="BoldCharacter"/>' : '') +
            (request ? '<w:b w:val="0"/>' : '') + '</w:rPr><w:t>قديم</w:t></w:r></w:p>';
          const story = shape.table ? 'body:t0:r0c0' : 'body';
          const xml = await saveTableEdit(path, runFixture(styles, paragraph, shape.table),
            (engine) => editFirstRun(engine, story, request));
          const properties = xml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
          const direct = properties.match(/<w:b\b([^>]*)\/>/);
          const effective = direct ? !/w:val="(?:0|false|off)"/.test(direct[1]!) : shape.bold;
          expect({ shape: shape.name, request, effective }).toEqual({
            shape: shape.name, request, effective: request || shape.bold,
          });
          if (request && !shape.bold) expect(properties).toMatch(/<w:b\b(?:\s[^>]*)?\/>/);
          if (!request) expect(properties).not.toMatch(/<w:b\b/);
        }
      }
    }, TIMEOUT);

    it.each(['b', 'i'] as const)('keeps direct %sCs while replacing Arabic under its non-CS style', async (property) => {
      const styles = '<w:style w:type="paragraph" w:styleId="Script">' +
        `<w:rPr><w:${property}/></w:rPr></w:style>`;
      for (const boundary of [false, true]) {
        const paragraph = '<w:p><w:pPr><w:pStyle w:val="Script"/></w:pPr>' +
          (boundary ? `<w:r><w:t>تمهيد</w:t></w:r><w:r><w:rPr><w:${property}Cs/></w:rPr><w:tab/></w:r>` : '') +
          `<w:r><w:rPr><w:${property}Cs/></w:rPr><w:t>قديم</w:t></w:r></w:p>`;
        const xml = await saveTableEdit(path, runFixture(styles, paragraph), (engine) => {
          const read = engine.readParagraphs({ story: 'body', view: 'accepted' });
          if (!read.ok) throw new Error(`read refused: ${JSON.stringify(read)}`);
          const paraId = read.paragraphs[0]!.paraId;
          const start = boundary ? 6 : 0;
          const range = { story: 'body', start: { paraId, offset: start }, end: { paraId, offset: start + 4 } };
          engine.replaceRange(range, 'جديد');
          engine.formatRange(range, property === 'b' ? { bold: false } : { italic: false });
          return ['body'];
        });
        expect(text(xml)).toBe(boundary ? 'تمهيدجديد' : 'جديد');
        const runs = Array.from(xml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? []);
        expect(runs.find((run) => run.includes('جديد'))).toContain(`<w:${property}Cs/>`);
        if (boundary) expect(runs.find((run) => run.includes('<w:tab'))).toContain(`<w:${property}Cs/>`);
      }
    }, TIMEOUT);
  });
});

describe('source run formatting after disjoint paragraph edits', () => {
  describe.each(Object.keys(SAVES) as (keyof typeof SAVES)[])('%s', (path) => {
    it.each([
      { name: 'Arabic bold', property: 'bCs', runs: ['اااا', 'بببب', 'جججج'], replacements: ['أ', 'چ'] },
      { name: 'Latin bold', property: 'bCs', runs: ['aaaa', 'bbbb', 'cccc'], replacements: ['A', 'C'] },
      { name: 'Arabic italic', property: 'iCs', runs: ['اااا', 'بببب', 'جججج'], replacements: ['أ', 'چ'] },
    ])('keeps each source run’s direct CS values for $name', async ({ property, runs, replacements }) => {
      const styles = '<w:style w:type="paragraph" w:styleId="P"><w:rPr><w:b/></w:rPr></w:style>';
      const paragraph = '<w:p><w:pPr><w:pStyle w:val="P"/></w:pPr>' +
        runs.map((value, index) =>
          `<w:r><w:rPr><w:${property}${index === 1 ? '' : ' w:val="0"'}/></w:rPr><w:t>${value}</w:t></w:r>`
        ).join('') + '</w:p>';
      const xml = await saveTableEdit(path, runFixture(styles, paragraph), (engine) => {
        const { paraId } = bodyParagraph(engine, runs[0]!);
        for (const [index, offset] of [0, 11].entries()) {
          engine.replaceRange({
            story: 'body', start: { paraId, offset }, end: { paraId, offset: offset + 1 },
          }, replacements[index]!);
        }
        return ['body'];
      });
      const savedRuns = Array.from(xml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g) ?? []);
      const values = savedRuns.flatMap((run) => {
        const direct = run.match(new RegExp(`<w:${property}\\b([^>]*)/>`));
        const value = direct ? !/w:val="(?:0|false|off)"/.test(direct[1]!) : undefined;
        return Array.from(text(run), () => value);
      });
      expect(values.slice(4, 8)).toEqual([true, true, true, true]);
      expect(values).toEqual([false, false, false, false, true, true, true, true, false, false, false, false]);
      expect(text(xml)).toBe(replacements[0]! + runs[0]!.slice(1) + runs[1]! + runs[2]!.slice(0, 3) + replacements[1]!);
      for (const run of savedRuns) {
        expect(run).not.toMatch(property === 'bCs' ? /<w:iCs\b/ : /<w:bCs\b/);
      }
    });
  });
});

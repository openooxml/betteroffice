import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { preloadEditWasm } from './wasm/index';
import { createYrsSession } from './index';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentSourceWorkerClient as ResidentEngineWorkerClient } from './residentSourceWorkerClient';
import { ResidentEngineWorkerClient as DefaultWorkerClient } from './residentEngineWorkerClient';
import type { DocxCommandPayloads, ResidentCommandInitRequest } from './residentCommandProtocol';
import { applyFrameDelta, decodeFrameDelta } from '../layout/render/frameDelta';

const TIMEOUT = 20_000;
const PNG = 'data:image/png;base64,iVBORwECAwQ=';
const LAYOUT = JSON.stringify({
  bodyStory: 'body', regions: { sections: [{ properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } },
  renderEnv: {},
});
let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
  startWorker = await residentWorkerFactory();
});

function source(paragraphs = 120, pageTotal = false): Uint8Array {
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${r}/image" Target="media/picture.png"/></Relationships>`)],
    ['word/media/picture.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4])],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${r}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>` +
      Array.from({ length: paragraphs }, (_, i) => `<w:p w14:paraId="${(i + 1).toString(16).padStart(8, '0')}"><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Paragraph ${i + 1}</w:t></w:r>` +
        (pageTotal && i === 0 ? `<w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>${paragraphs}</w:t></w:r></w:fldSimple>` : '') + '</w:p>').join('') +
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr/></w:body></w:document>')],
  ])));
}

function init(): ResidentCommandInitRequest {
  return {
    session: 'source-worker', incarnation: 1, epoch: 1, readOnly: false,
    fonts: [new Uint8Array(readFileSync(resolve(process.cwd(), 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')))],
    fontRegistrationSlots: true, layoutInput: LAYOUT, layoutWithRegions: true,
    layoutRevision: 1, extras: JSON.stringify({ fontChains: { 'liberation sans|0|0': [0] } }), displayWindow: [0, 1],
  };
}

function firstPage(buffers: readonly (ArrayBuffer | Uint8Array)[]) {
  let frame: ReturnType<typeof applyFrameDelta> | null = null;
  for (const bytes of buffers) frame = applyFrameDelta(frame, decodeFrameDelta(bytes));
  const page = frame!.pages.find(({ pageIndex }) => pageIndex === 0)!.page;
  return { width: page.width, height: page.height, runs: page.primitives.flatMap(primitive =>
    primitive.kind === 'glyphRun' ? [{ text: primitive.text, glyphs: primitive.glyphs }] : primitive.kind === 'text' ? [{ text: primitive.text }] : []) };
}

async function parsed(bytes: Uint8Array) {
  const parser = await createYrsSession();
  try {
    parser.openDocx(bytes, true);
    return parser.materializeDocx()!;
  } finally {
    parser.destroy();
  }
}

for (const input of ['bytes', 'document'] as const) {
  test(`${input} previews hand off to editing and save without another parse`, async () => {
    const bytes = source();
    const document = input === 'document' ? await parsed(bytes) : undefined;
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    const settings = init();
    try {
      const opened = await client.openSource({
        ...(document ? { document } : { bytes }), mediaTokens: true,
        preview: { ...settings, pages: 1 },
      });
      expect(bytes.byteLength).toBeGreaterThan(0);
      expect(opened.document.package.document.content).toEqual([]);
      expect(opened.preview?.documentPreview).toBe(true);
      expect(opened.preview?.layoutProvisional).toBe(true);
      const previewFrame = applyFrameDelta(null, decodeFrameDelta(opened.preview!.frame));
      expect(previewFrame.pages.filter(page => !page.page.unbuilt)).toHaveLength(1);
      expect(previewFrame.pages.length).toBeLessThan(121);
      const engine = worker.sessions[0]!;
      expect(engine.paragraphs('body').length).toBe(121);
      const parse = spyOn(engine, 'openDocx').mockImplementation(() => { throw new Error('second parse'); });
      const prefix = spyOn(engine, 'layoutDocumentWithRegionsPrefixRetainedJson');
      const initial = await client.commandInit(settings);
      expect(initial.frame.payload.layoutProvisional).not.toBe(true);
      expect(parse).not.toHaveBeenCalled();
      expect(prefix).not.toHaveBeenCalled();
      expect(worker.sessions).toHaveLength(1);
      const paraId = engine.paragraphs('body')[0]!.paraId;
      const ack = await client.edit({
        session: initial.session, incarnation: initial.incarnation, seq: 1,
        basis: { epoch: initial.epoch, version: initial.version }, kind: 'insertText',
        payload: { text: 'Edited ', selection: { anchor: { story: 'body', paraId, offset: 0 }, focus: { story: 'body', paraId, offset: 0 } } },
      });
      expect(ack.outcomes[0]?.status).toBe('applied');
      expect(engine.paragraphs('body')[0]!.text).toBe('Edited Paragraph 1');
      const saved = await client.save({ comments: [] });
      const packageParts = unzipContainer(new Uint8Array(saved.bytes));
      expect(new TextDecoder().decode(packageParts['word/document.xml'])).toContain('Edited ');
      expect(packageParts['word/media/picture.png']).toEqual(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]));
      if (document) expect(JSON.stringify(document.package.document.content)).not.toContain('Edited ');
    } finally {
      client.destroy();
    }
  }, TIMEOUT);

  test(`${input} media tokens resolve in the worker and do not leak to the host`, async () => {
    const bytes = source(1);
    const document = input === 'document' ? await parsed(bytes) : undefined;
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    try {
      const opened = await client.openSource({ ...(document ? { document } : { bytes }), mediaTokens: true });
      const engine = worker.sessions[0]!;
      const images = engine.storySegments('body').filter(segment => segment.kind === 'embed' && segment.embedKind === 'image');
      expect(images).toHaveLength(1);
      const image = images[0]!;
      if (image.kind !== 'embed') throw new Error('Missing image');
      const token = (image.payload as { src: string }).src;
      expect(token).toMatch(/^media:\d+$/);
      expect(engine.mediaDataUrl(token)).toBe(PNG);
      expect(JSON.stringify(opened.document)).not.toContain('media:');
    } finally {
      client.destroy();
    }
  }, TIMEOUT);
}

test('prewarm loads the worker without creating a document session', async () => {
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.prewarm();
    expect(worker.requests).toContain('warm');
    expect(worker.sessions).toHaveLength(0);
    await client.openSource({ bytes: source(1), preview: true });
    expect(worker.sessions).toHaveLength(1);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

for (const section of ['first', '0']) {
  test(`earlier section ${section} properties survive save and reopen`, async () => {
    const parts = unzipContainer(source(1));
    parts['word/document.xml'] = toBytes('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
    const document = await parsed(new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(parts)))));
    if (section === 'first') document.package.document.sections![0]!.id = section;
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    try {
      await client.openSource({ document });
      const initial = await client.commandInit(init());
      const properties = await client.documentRead({ kind: 'properties', target: { kind: 'section', section } });
      expect(properties.value).toMatchObject({ properties: { marginLeft: 1440, pageWidth: 12240 } });
      const changed = await client.edit({ session: initial.session, incarnation: initial.incarnation, seq: 1,
        basis: { epoch: initial.epoch, version: initial.version }, kind: 'setSectionProperties',
        payload: { section, patch: { marginLeft: 2880, pageWidth: 10000 }, expectVersion: properties.version } });
      expect(changed.outcomes[0]?.status).toBe('applied');
      expect((await client.documentRead({ kind: 'properties', target: { kind: 'section', section } })).value)
        .toMatchObject({ properties: { marginLeft: 2880, pageWidth: 10000 } });
      const reopened = await parsed(new Uint8Array((await client.save({ comments: [] })).bytes));
      expect(reopened.package.document.sections![0]!.properties).toMatchObject({ marginLeft: 2880, pageWidth: 10000 });
      expect(reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 1440, pageWidth: 12240 });
      expect((await client.edit({ session: initial.session, incarnation: initial.incarnation, seq: 2,
        basis: { epoch: initial.epoch, version: 1 }, kind: 'undo', payload: {} })).outcomes[0]?.status).toBe('applied');
      expect((await client.documentRead({ kind: 'properties', target: { kind: 'section', section } })).value)
        .toMatchObject({ properties: { marginLeft: 1440, pageWidth: 12240 } });
      expect(worker.sessions[0]!.paragraphs('body')[0]!.properties.sectPr).toMatchObject({ marginLeft: 1440, pageWidth: 12240 });
      const undone = await parsed(new Uint8Array((await client.save({ comments: [] })).bytes));
      expect(undone.package.document.sections![0]!.properties).toMatchObject({ marginLeft: 1440, pageWidth: 12240 });
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('a checkpoint reopens a source in a fresh worker that resumes commands at its version and saves the same package', async () => {
  const bytes = source(1);
  const workers = [startWorker(), startWorker()];
  const [first, second] = workers.map(worker => new ResidentEngineWorkerClient(worker));
  try {
    await first!.openSource({ bytes });
    const initial = await first!.commandInit(init());
    const paraId = workers[0]!.sessions[0]!.paragraphs('body')[0]!.paraId;
    const insert = (client: ResidentEngineWorkerClient, incarnation: number, seq: number, offset: number, text: string) => client.edit({
      session: initial.session, incarnation, seq, basis: { epoch: initial.epoch, version: seq - 1 }, kind: 'insertText',
      payload: { text, selection: { anchor: { story: 'body', paraId, offset }, focus: { story: 'body', paraId, offset } } },
    });
    expect((await insert(first!, 1, 1, 0, 'One ')).outcomes[0]?.status).toBe('applied');
    const checkpoint = await first!.checkpoint();
    expect(checkpoint).toMatchObject({ epoch: initial.epoch, version: 1, throughSeq: 1 });
    expect((await insert(first!, 1, 2, 4, 'Two ')).outcomes[0]?.status).toBe('applied');
    const expected = unzipContainer(new Uint8Array((await first!.save({ comments: [] })).bytes));

    await second!.openSource({ bytes, state: checkpoint.state });
    const resumed = await second!.commandInit({ ...init(), incarnation: 2, version: checkpoint.version, throughSeq: checkpoint.throughSeq, recoveryState: checkpoint.recoveryState });
    expect(resumed).toMatchObject({ incarnation: 2, version: 1, throughSeq: 1 });
    expect(workers[1]!.sessions[0]!.paragraphs('body').map(paragraph => [paragraph.paraId, paragraph.text]))
      .toEqual([[paraId, 'One Paragraph 1'], expect.anything()]);
    expect((await insert(second!, 2, 2, 4, 'Two ')).outcomes[0]).toMatchObject({ seq: 2, status: 'applied' });
    expect(workers[1]!.sessions[0]!.paragraphs('body')[0]!.text).toBe('One Two Paragraph 1');
    expect(unzipContainer(new Uint8Array((await second!.save({ comments: [] })).bytes))).toEqual(expected);
  } finally {
    first!.destroy();
    second!.destroy();
  }
}, TIMEOUT);

test('a source state needs whole source bytes', async () => {
  const client = new ResidentEngineWorkerClient(startWorker());
  try {
    for (const invalid of [{ document: await parsed(source(1)) }, { bytes: source(1), preview: true }, { bytes: source(1), cut: { blocks: 1 } }]) {
      await expect(client.openSource({ ...invalid, state: Uint8Array.of(0) } as never)).rejects.toThrow('Invalid DOCX source state');
    }
  } finally { client.destroy(); }
}, TIMEOUT);

test('source opening adopts a warmed client and retains its worker and memory subscription', async () => {
  const worker = startWorker();
  const warmed = new DefaultWorkerClient(worker);
  await warmed.warm();
  const client = ResidentEngineWorkerClient.adopt(warmed);
  const samples: unknown[] = [];
  client.onMemory(value => samples.push(value));
  try {
    expect(client as DefaultWorkerClient).toBe(warmed);
    expect(ResidentEngineWorkerClient.adopt(client)).toBe(client);
    await client.openSource({ bytes: source(1), preview: true });
    expect(worker.sessions).toHaveLength(1);
    expect(worker.requests.filter(type => type === 'warm')).toHaveLength(1);
    expect(samples.length).toBeGreaterThan(0);
    expect((await client.commandInit(init())).frame).toBeDefined();
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('source preview pixels resolve media inside the worker before editing starts', async () => {
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    const opened = await client.openSource({ bytes: source(1), mediaTokens: true, preview: { ...init(), pages: 1 } });
    const frame = applyFrameDelta(null, decodeFrameDelta(opened.preview!.frame));
    const page = frame.pages.find(page => page.page.primitives.some(primitive => primitive.kind === 'image'))!;
    expect(page).toBeDefined();
    await client.attachCanvases([{ pageId: String(page.pageId), canvas: { width: 1, height: 1 } as OffscreenCanvas }],
      [String(page.pageId)], 1, 1, { color: '#000', width: 2 });
    expect(worker.imageBlobs).toHaveLength(1);
    expect(worker.imageBlobs[0]!.type).toBe('image/png');
    expect(new Uint8Array(await worker.imageBlobs[0]!.arrayBuffer())).toEqual(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]));
    expect(client.isReady()).toBe(false);
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('invalid or failed sources leave the client available for a new opening', async () => {
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await expect(client.openSource({} as never)).rejects.toThrow('exactly one');
    await expect(client.openSource({ bytes: new Uint8Array([1, 2, 3]) })).rejects.toThrow();
    await client.openSource({ bytes: source(1), preview: false });
    await expect(client.openSource({ bytes: source(1) })).rejects.toThrow('already holds');
    expect(client.isReady()).toBe(false);
    expect((await client.commandInit(init())).frame).toBeDefined();
  } finally {
    client.destroy();
  }
}, TIMEOUT);

test('foreign and invalid packages reject opening and allow a valid source afterward', async () => {
  const missing = new Map(Object.entries(unzipContainer(source(1))));
  missing.delete('word/document.xml');
  const malformed = new Map(Object.entries(unzipContainer(source(1))));
  malformed.set('word/document.xml', toBytes('<w:document><w:body></w:document>'));
  const unsupported = new Map(Object.entries(unzipContainer(source(1))));
  unsupported.set('word/document.xml', toBytes('<w:document xmlns:w="urn:unsupported-ooxml:main"><w:body/></w:document>'));
  const spreadsheet = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')],
    ['xl/workbook.xml', toBytes('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>')],
  ]);
  const archive = new Map([['readme.txt', toBytes('Archive contents')]]);
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    for (const parts of [missing, malformed, unsupported, spreadsheet, archive]) {
      const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
      await expect(client.openSource({ bytes })).rejects.toThrow(/namespace|unsupported|invalid|format|document|root/i);
      expect(client.isReady()).toBe(false);
    }
    await client.openSource({ bytes: source(1), preview: false });
    expect((await client.commandInit(init())).frame).toBeDefined();
  } finally {
    client.destroy();
  }
}, TIMEOUT);

for (const input of ['bytes', 'document'] as const) {
  test(`viewport-first ${input} opening previews its first page with the command fonts before commands start`, async () => {
    const bytes = source();
    const document = input === 'document' ? await parsed(bytes) : undefined;
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    try {
      const opened = await client.openSource(document ? { document } : { bytes });
      expect(opened.preview).toBeUndefined();
      const { fonts, fontRegistrationSlots, layoutInput, extras } = init();
      const preview = await client.previewSource({ fonts, fontRegistrationSlots, layoutInput, extras });
      expect(preview.documentPreview).toBe(true);
      const pages = applyFrameDelta(null, decodeFrameDelta(preview.frame)).pages.filter(({ page }) => !page.unbuilt);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.page.primitives.some(primitive => primitive.kind === 'glyphRun' || primitive.kind === 'text')).toBe(true);
      const initial = await client.commandInit({ ...init(), ...{ provisionalPages: 1 } });
      expect(initial.frame.payload.layoutProvisional).toBe(true);
      await expect(client.previewSource({ layoutInput })).rejects.toThrow('before commands');
    } finally { client.destroy(); }
  });
}

for (const pageTotal of [false, true]) {
  test(`a source preview paints the first page the command session paints${pageTotal ? ' with a saved page total' : ''}`, async () => {
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    try {
      await client.openSource({ bytes: source(120, pageTotal) });
      const { fonts, fontRegistrationSlots, layoutInput, extras } = init();
      const preview = await client.previewSource({ fonts, fontRegistrationSlots, extras,
        layoutInput: JSON.stringify({ ...JSON.parse(layoutInput), cachedPageTotals: true }) });
      const initial = await client.commandInit({ ...init(), ...{ provisionalPages: 1 } });
      expect(initial.frame.payload.layoutProvisional === true).toBe(!pageTotal);
      const previewed = firstPage([...(preview.pageFrames ?? []), preview.frame]);
      expect(previewed).toEqual(firstPage([...(initial.frame.payload.pageFrames ?? []), initial.frame.payload.frame]));
      expect(previewed.runs.map(({ text }) => text).join('')).toBe(pageTotal ? 'Paragraph 1 of 120' : 'Paragraph 1');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

for (const pageTotal of [false, true]) {
  test(`a source cut previews the first page the whole source's command session paints${pageTotal ? ' with a saved page total' : ''}`, async () => {
    const worker = startWorker();
    const client = new ResidentEngineWorkerClient(worker);
    try {
      const bytes = source(400, pageTotal);
      expect(await client.openSourceCut({ bytes, cut: { blocks: 20 } })).toBeDefined();
      const cutParagraphs = worker.sessions.at(-1)!.geometryReader.paragraphs('body').length;
      await expect(client.commandInit(init())).rejects.toThrow('whole opened document');
      const { fonts, fontRegistrationSlots, layoutInput, extras } = init();
      const preview = await client.previewSource({ fonts, fontRegistrationSlots, extras,
        layoutInput: JSON.stringify({ ...JSON.parse(layoutInput), cachedPageTotals: true }) });
      await expect(client.openSource({ bytes, cut: { blocks: 20 } })).rejects.toThrow('openSourceCut');
      await client.openSource({ bytes });
      expect(cutParagraphs).toBeGreaterThan(0);
      expect(cutParagraphs).toBeLessThan(worker.sessions.at(-1)!.geometryReader.paragraphs('body').length);
      const initial = await client.commandInit({ ...init(), ...{ provisionalPages: 1 } });
      const previewed = firstPage([...(preview.pageFrames ?? []), preview.frame]);
      expect(previewed).toEqual(firstPage([...(initial.frame.payload.pageFrames ?? []), initial.frame.payload.frame]));
      expect(previewed.runs.map(({ text }) => text).join('')).toBe(pageTotal ? 'Paragraph 1 of 400' : 'Paragraph 1');
      await expect(client.openSourceCut({ bytes, cut: { blocks: 20 } })).rejects.toThrow('already holds');
    } finally { client.destroy(); }
  }, TIMEOUT);
}

test('a whole source that fails to open leaves the worker on the cut it was to replace', async () => {
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    const bytes = source(400);
    expect(await client.openSourceCut({ bytes, cut: { blocks: 20 } })).toBeDefined();
    await expect(client.openSource({ bytes: new Uint8Array([1, 2, 3]) })).rejects.toThrow();
    const { fonts, fontRegistrationSlots, layoutInput, extras } = init();
    const preview = await client.previewSource({ fonts, fontRegistrationSlots, layoutInput, extras });
    expect(firstPage([...(preview.pageFrames ?? []), preview.frame]).runs.map(({ text }) => text).join('')).toBe('Paragraph 1');
    await client.openSource({ bytes });
    expect((await client.commandInit({ ...init(), ...{ provisionalPages: 1 } })).frame).toBeDefined();
  } finally { client.destroy(); }
}, TIMEOUT);

test('a whole source that fails after its cut failed leaves the client free for a new cut', async () => {
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    const cut = client.openSourceCut({ bytes: new Uint8Array([1, 2, 3]), cut: { blocks: 20 } }).catch(() => null);
    await expect(client.openSource({ bytes: new Uint8Array([1, 2, 3]) })).rejects.toThrow();
    await cut;
    expect(await client.openSourceCut({ bytes: source(400), cut: { blocks: 20 } })).toBeDefined();
  } finally { client.destroy(); }
}, TIMEOUT);

for (const dependent of [false, true]) {
  test(`checkpoint recovery restores caret and stored marks${dependent ? ' across the predecessor boundary' : ''}`, async () => {
    const bytes = source(1);
    const workers = [startWorker(), startWorker()];
    const clients = workers.map(worker => new ResidentEngineWorkerClient(worker));
    const [first, second] = clients;
    try {
      await first!.openSource({ bytes });
      const initial = await first!.commandInit(init());
      const edit = (seq: number, version: number, kind: string, payload: object) => ({
        session: initial.session, incarnation: 1, seq, basis: { epoch: 1, version }, kind, payload,
      });
      expect((await first!.edit(edit(1, 0, 'insertText', { text: 'kept' }))).outcomes[0]?.status).toBe('applied');
      expect((await first!.edit(edit(2, 1, 'formatSelection', { delta: { bold: true } }))).outcomes[0]?.status).toBe('noop');
      const checkpoint = await first!.checkpoint();
      const suffix = { ...edit(3, 1, 'insertText', { text: ' later' }), ...(dependent ? { afterSeq: 1 } : {}) };
      expect((await first!.edit(suffix)).outcomes[0]?.status).toBe('applied');
      const expected = workers[0]!.sessions[0]!.paragraphSegments('body', [0]);
      await second!.openSource({ bytes, state: checkpoint.state });
      const reopened = await second!.commandInit({ ...init(), incarnation: 2, version: checkpoint.version, throughSeq: checkpoint.throughSeq, recoveryState: checkpoint.recoveryState, replayAppliedSeqs: [3] });
      expect(reopened.selection?.focus.offset).toBe(4);
      expect(reopened.editState.storedMarks).toMatchObject({ delta: { bold: true }, clear: false });
      expect((await second!.edit({ ...suffix, incarnation: 2, replay: true })).outcomes[0]?.status).toBe('applied');
      expect(workers[1]!.sessions[0]!.paragraphs('body')[0]!.text).toBe('kept laterParagraph 1');
      expect(workers[1]!.sessions[0]!.paragraphSegments('body', [0])).toEqual(expected);
    } finally { clients.forEach(client => client!.destroy()); }
  }, TIMEOUT);
}

test('previously applied typing replays after the editor becomes read-only', async () => {
  const worker = startWorker();
  const client = new ResidentEngineWorkerClient(worker);
  try {
    await client.openSource({ bytes: source(1) });
    await client.commandInit({ ...init(), incarnation: 2, readOnly: true, replayAppliedSeqs: [1] });
    const ack = await client.edit({ session: 'source-worker', incarnation: 2, seq: 1, replay: true,
      basis: { epoch: 1, version: 0 }, kind: 'insertText', payload: { text: 'kept' } });
    expect(ack.outcomes[0]?.status).toBe('applied');
    expect(worker.sessions[0]!.paragraphs('body')[0]!.text).toBe('keptParagraph 1');
    expect(ack.editState?.readOnly).toBe(true);
    expect((await client.edit({ session: 'source-worker', incarnation: 2, seq: 2, replay: true,
      basis: { epoch: 1, version: 1 }, kind: 'insertText', payload: { text: 'blocked' } })).outcomes[0])
      .toMatchObject({ status: 'refused', reason: 'read-only' });
  } finally { client.destroy(); }
}, TIMEOUT);

test('checkpoint recovery restores proposal records and preview decisions', async () => {
  const bytes = source(1);
  const workers = [startWorker(), startWorker()];
  const clients = workers.map(worker => new ResidentEngineWorkerClient(worker));
  try {
    await clients[0]!.openSource({ bytes });
    await clients[0]!.commandInit(init());
    const identities = await clients[0]!.documentRead({ kind: 'paragraphIdentities' });
    const proposed = await clients[0]!.edit({ session: 'source-worker', incarnation: 1, seq: 1,
      basis: { epoch: 1, version: 0 }, kind: 'propose', payload: { expectVersion: identities.version,
        proposals: [{ id: 'retained', paragraph: identities.value.paragraphs[0]!.session!,
          op: 'insertText', at: 'end', text: 'Proposal', suggest: { author: 'Writer', date: '2026-01-01' } }] } });
    expect(proposed.outcomes[0]?.status).toBe('applied');
    const review = await clients[0]!.documentRead({ kind: 'reviewSnapshot' });
    expect((await clients[0]!.edit({ session: 'source-worker', incarnation: 1, seq: 2,
      basis: { epoch: 1, version: proposed.version }, kind: 'setProposalStates', payload: {
        expectVersion: review.version, expectPreviewVersion: review.value!.proposals.previewVersion,
        changes: [{ id: 'retained', state: 'accepted' }] } })).outcomes[0]?.status).toBe('applied');
    const checkpoint = await clients[0]!.checkpoint();
    const before = (await clients[0]!.documentRead({ kind: 'reviewSnapshot' })).value!;
    await clients[1]!.openSource({ bytes, state: checkpoint.state });
    await clients[1]!.commandInit({ ...init(), incarnation: 2, version: checkpoint.version, throughSeq: checkpoint.throughSeq, recoveryState: checkpoint.recoveryState });
    const after = (await clients[1]!.documentRead({ kind: 'reviewSnapshot' })).value!;
    expect(after.proposals.proposals).toEqual(before.proposals.proposals);
    expect(after.proposals.previewVersion).toBe(before.proposals.previewVersion);
    expect(after.revisions).toEqual(before.revisions);
    const changed = await clients[1]!.edit({ session: 'source-worker', incarnation: 2, seq: 3,
      basis: { epoch: 1, version: checkpoint.version }, kind: 'setProposalStates', payload: {
        expectVersion: after.version, expectPreviewVersion: after.proposals.previewVersion,
        changes: [{ id: 'retained', state: 'rejected' }] } });
    expect(changed.outcomes[0]?.status).toBe('applied');
    expect((await clients[1]!.documentRead({ kind: 'reviewSnapshot' })).value!.proposals.proposals[0]?.state).toBe('rejected');
  } finally { clients.forEach(client => client.destroy()); }
}, TIMEOUT);


test('checkpoint recovery retains resolved reservations and table gridlines', async () => {
  const bytes = source(1);
  const workers = [startWorker(), startWorker()];
  const clients = workers.map(worker => new ResidentEngineWorkerClient(worker));
  try {
    await clients[0]!.openSource({ bytes });
    const settings = { ...init(), layoutInput: JSON.stringify({ ...JSON.parse(LAYOUT), renderEnv: { showTableGridlines: true } }) };
    await clients[0]!.commandInit(settings);
    const identities = await clients[0]!.documentRead({ kind: 'paragraphIdentities' });
    const environment: DocxCommandPayloads['setRenderEnvironment'] = { themeColors: {}, showHiddenText: false, showTableGridlines: true, reservations: {
      version: identities.version, key: 'retained', sets: [{ owner: 'layout', items: [{ id: 'space', kind: 'block',
        anchor: { kind: 'paragraph', paragraph: identities.value.paragraphs[0]!.session! }, side: 'before', height: 50 }] }],
    } };
    const changed = await clients[0]!.edit({ session: 'source-worker', incarnation: 1, seq: 1,
      basis: { epoch: 1, version: 0 }, kind: 'setRenderEnvironment', payload: environment });
    expect(changed.outcomes[0]?.status).toBe('noop');
    const checkpoint = await clients[0]!.checkpoint();
    const before = (await clients[0]!.requestFullFrame()).frame!;
    await clients[1]!.openSource({ bytes, state: checkpoint.state });
    const layout = spyOn(workers[1]!.sessions[0]!, 'layoutDocumentWithRegionsAfterEdit');
    const resumed = await clients[1]!.commandInit({ ...settings, incarnation: 2, version: checkpoint.version,
      throughSeq: checkpoint.throughSeq, recoveryState: checkpoint.recoveryState });
    const restored = JSON.parse(layout.mock.calls.at(-1)![0]).renderEnv;
    expect(restored.showTableGridlines).toBe(true);
    expect(Object.values(restored.reservations ?? {})).toMatchObject([{ owner: 'layout', id: 'space', height: 50 }]);
    expect(resumed.frame.payload.reservations).toEqual(before.payload.reservations);
    expect(firstPage([resumed.frame.payload.frame])).toEqual(firstPage([before.payload.frame]));
    layout.mockRestore();
  } finally { clients.forEach(client => client.destroy()); }
}, TIMEOUT);

for (const remove of [false, true]) {
  test(`checkpoint recovery retains source comment ${remove ? 'deletion' : 'metadata'} and revision refs`, async () => {
    const parts = new Map(Object.entries(unzipContainer(source(1))));
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    parts.set('[Content_Types].xml', toBytes(new TextDecoder().decode(parts.get('[Content_Types].xml'))
      .replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')));
    parts.set('word/_rels/document.xml.rels', toBytes(new TextDecoder().decode(parts.get('word/_rels/document.xml.rels'))
      .replace('</Relationships>', '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')));
    parts.set('word/document.xml', toBytes(`<w:document xmlns:w="${w}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:commentRangeStart w:id="0"/><w:r><w:t>Text</w:t></w:r><w:ins w:id="7" w:author="Writer" w:date="2026-01-01"><w:r><w:t>Revision</w:t></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr/></w:body></w:document>`));
    parts.set('word/comments.xml', toBytes(`<w:comments xmlns:w="${w}"><w:comment w:id="0" w:author="Writer"><w:p><w:r><w:t>Original</w:t></w:r></w:p></w:comment></w:comments>`));
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
    const clients = [startWorker(), startWorker()].map(worker => new ResidentEngineWorkerClient(worker));
    try {
      await clients[0]!.openSource({ bytes });
      await clients[0]!.commandInit(init());
      const review = (await clients[0]!.documentRead({ kind: 'reviewSnapshot' })).value!;
      expect(review.comments).toHaveLength(1);
      expect(review.revisions).toHaveLength(1);
      const id = review.comments[0]!.id;
      const changed = await clients[0]!.edit({ session: 'source-worker', incarnation: 1, seq: 1,
        basis: { epoch: 1, version: 0 }, kind: remove ? 'removeComment' : 'updateComment',
        payload: remove ? { id, expectVersion: review.version } : { id, expectVersion: review.version, patch: { author: 'Editor', body: 'Changed', done: true } } });
      expect(changed.outcomes[0]?.status).toBe('applied');
      const checkpoint = await clients[0]!.checkpoint();
      const before = (await clients[0]!.documentRead({ kind: 'reviewSnapshot' })).value!;
      await clients[1]!.openSource({ bytes, state: checkpoint.state });
      await clients[1]!.commandInit({ ...init(), incarnation: 2, version: checkpoint.version,
        throughSeq: checkpoint.throughSeq, recoveryState: checkpoint.recoveryState });
      const after = (await clients[1]!.documentRead({ kind: 'reviewSnapshot' })).value!;
      expect(after.comments).toEqual(before.comments);
      expect(after.uiComments).toEqual(before.uiComments);
      expect(after.commentIds).toEqual(before.commentIds);
      expect(after.revisions).toEqual(before.revisions);
      expect(unzipContainer(new Uint8Array((await clients[1]!.save({ comments: [] })).bytes)))
        .toEqual(unzipContainer(new Uint8Array((await clients[0]!.save({ comments: [] })).bytes)));
    } finally { clients.forEach(client => client.destroy()); }
  }, TIMEOUT);
}


test('recovery keeps reservations from a replayed render-settings update', async () => {
  const bytes = source(1);
  const clients = [startWorker(), startWorker()].map(worker => new ResidentEngineWorkerClient(worker));
  try {
    await clients[0]!.openSource({ bytes });
    await clients[0]!.commandInit(init());
    const identities = await clients[0]!.documentRead({ kind: 'paragraphIdentities' });
    const environment: DocxCommandPayloads['setRenderEnvironment'] = { themeColors: {}, showHiddenText: false, reservations: {
      version: identities.version, key: 'retained', sets: [{ owner: 'layout', items: [{ id: 'space', kind: 'block',
        anchor: { kind: 'paragraph', paragraph: identities.value.paragraphs[0]!.session! }, side: 'before', height: 50 }] }],
    } };
    await clients[0]!.edit({ session: 'source-worker', incarnation: 1, seq: 1,
      basis: { epoch: 1, version: 0 }, kind: 'setRenderEnvironment', payload: environment });
    const checkpoint = await clients[0]!.checkpoint();
    const update = { kind: 'setRenderEnvironment', payload: { ...environment, showHiddenText: true } } as const;
    expect((await clients[0]!.edit({ session: 'source-worker', incarnation: 1, seq: 2,
      basis: { epoch: 1, version: checkpoint.version }, ...update })).outcomes[0]?.status).toBe('noop');
    await clients[1]!.openSource({ bytes, state: checkpoint.state });
    await clients[1]!.commandInit({ ...init(), incarnation: 2, version: checkpoint.version,
      throughSeq: checkpoint.throughSeq, recoveryState: checkpoint.recoveryState });
    const replayed = await clients[1]!.edit({ session: 'source-worker', incarnation: 2, seq: 2, replay: true,
      basis: { epoch: 1, version: checkpoint.version }, ...update });
    expect(replayed.outcomes[0]?.status).toBe('noop');
    const recovered = (await clients[1]!.checkpoint()).recoveryState?.layoutReservations;
    expect(Object.values(recovered?.values ?? {})).toMatchObject([{ owner: 'layout', id: 'space', height: 50 }]);
  } finally { clients.forEach(client => client.destroy()); }
}, TIMEOUT);

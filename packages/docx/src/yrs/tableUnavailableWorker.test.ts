import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadOpcWasm } from '../wasm/opc';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { DocxSourceWorkerClient } from './residentSourceWorkerClient';
import { createYrsSession, type YrsCellLoc, type YrsSession } from './index';
import { closeCommandEngines, openCommandEngine } from './residentCommandTestSupport';

let start: Awaited<ReturnType<typeof residentWorkerFactory>>;
beforeAll(async () => {
  await openCommandEngine();
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'))));
  start = await residentWorkerFactory({ transferBuffers: true });
}, 30_000);
afterEach(closeCommandEngines);

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;

function documentBytes(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph('Before')}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr>${cell('First')}${cell('Second')}</w:tr></w:tbl>${paragraph('After')}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const at = (row: number, column: number): YrsCellLoc => ({ story: 'body', tableIndex: 0, row, column });

function exchange(a: YrsSession, b: YrsSession): void {
  const fromA = a.encodeState();
  const fromB = b.encodeState();
  b.applyUpdate(fromA);
  a.applyUpdate(fromB);
}

async function open(clientId: number, bytes: Uint8Array, initialUpdate?: Uint8Array) {
  const worker = start(clientId);
  const client = new DocxSourceWorkerClient(worker);
  await client.openSource({ bytes, collaboration: { clientId, initialUpdate } });
  const engine = worker.sessions[0]!;
  const at = { story: 'body', paraId: engine.paragraphs('body')[0]!.paraId, offset: 0 };
  await client.commandInit({ session: `replica-${clientId}`, incarnation: 1, epoch: 1,
    readOnly: false, selection: { anchor: at, focus: at },
    fonts: [new Uint8Array(readFileSync('crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))],
    layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [0] } }, renderEnv: {} }),
    layoutWithRegions: true, layoutRevision: 1, extras: '{}',
  });
  return { client, engine };
}

test('a worker that joins a document whose table crossing undos emptied warns once with a stable code', async () => {
  const bytes = documentBytes();
  const seeded = await open(9811, bytes);
  const peers: YrsSession[] = [];
  try {
    const peer = async (clientId: number) => {
      const created = await createYrsSession({ clientId });
      peers.push(created);
      created.openDocx(bytes, false);
      created.applyUpdate(seeded.engine.encodeState());
      return created;
    };
    const a = await peer(9812);
    const b = await peer(9813);
    a.insertRow(at(0, 0), 'below');
    exchange(a, b);
    b.setColumnWidth(at(0, 0), 2000);
    expect(a.undo()).toBe(true);
    exchange(a, b);
    expect(b.undo()).toBe(true);
    exchange(a, b);
    const payload = a.tablePayload('body', 0) ?? {};
    expect(['tblPr', 'grid', 'rows'].filter((key) => key in payload)).toEqual([]);

    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const joined = await open(9814, bytes, a.encodeState());
      expect(joined.engine.paragraphs('body').map(({ text }) => text)).toEqual(['Before', 'After']);
      const warnings = warn.mock.calls.filter(([message]) =>
        typeof message === 'string' && message.startsWith('table-unavailable:'));
      expect(warnings).toHaveLength(1);
      expect(warnings[0]![1]).toEqual([{ story: 'body', tableIndex: 0 }]);
      const before = joined.engine.encodeState();
      await expect(joined.client.save({ comments: [] })).rejects.toThrow(
        'a table is unavailable after a concurrent undo');
      expect(joined.engine.encodeState()).toEqual(before);
      expect(warn.mock.calls.filter(([message]) =>
        typeof message === 'string' && message.startsWith('table-unavailable:'))).toHaveLength(1);
    } finally {
      warn.mockRestore();
    }
  } finally {
    for (const created of peers) created.destroy();
  }
}, 60_000);

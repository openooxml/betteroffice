import { beforeAll, expect, setSystemTime, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadOpcWasm, unzipContainer } from './wasm';
import { writeDocxS13WithWarningsWire } from './parseWasm';
import { parseDocx } from './parser';
import { repackDocx, repackDocxWithWarnings } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { serializeDocumentSaveModel } from './modelSave';
import { EDIT_PARAGRAPH_ID, oneSavePackage } from '../yrs/__fixtures__/oneSavePackage';
import { createYrsSession } from '../yrs/index';
import { yrsToDocument } from '../yrs/yrsToDocument';

const ROOT = resolve(import.meta.dir, '../../../..');
const CORE = 'docProps/core.xml';
const CODE = 'DOCX_MODEL_PARTS_REWRITTEN';
const NOW = '2030-01-02T03:04:05.006Z';
const WITHOUT_STAMP = { updateModifiedDate: false };
const parse = (bytes: Uint8Array) => parseDocx(bytes, { preloadFonts: false });

beforeAll(() => preloadOpcWasm());

const fixturePaths = [
  ...[
    'crates/docx-edit/tests/fixtures',
    'crates/betteroffice-docx/tests/corpus/fixtures',
    'packages/docx/src/yrs/__fixtures__',
  ].flatMap((dir) => readdirSync(resolve(ROOT, dir), { recursive: true, encoding: 'utf8' })
    .filter((name) => name.endsWith('.docx'))
    .map((name) => `${dir}/${name}`)),
  'apps/demo/public/betteroffice-demo.docx',
  'crates/ooxml-opc/tests/fixtures/betteroffice-demo.docx',
  'crates/ooxml-redact/tests/fixtures/redaction-integrity.docx',
  'packages/docx-react/src/components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx',
];

async function expectUnchanged(source: Uint8Array) {
  const document = await parse(source);
  const result = await repackDocxWithWarnings(document, WITHOUT_STAMP);
  expect(new Uint8Array(result.buffer)).toEqual(Uint8Array.from(source));
  expect(result.warnings).toEqual(document.warnings ?? []);
}

for (const path of fixturePaths.sort()) {
  test(`unchanged model saves byte for byte: ${path}`, async () => {
    await expectUnchanged(new Uint8Array(readFileSync(resolve(ROOT, path))));
  }, 30_000);
}

test('unchanged synthetic model saves byte for byte with stamping disabled', async () => {
  const { bytes } = oneSavePackage();
  await expectUnchanged(bytes);
  expect(new Uint8Array(await repackDocx(await parse(bytes), WITHOUT_STAMP)))
    .toEqual(Uint8Array.from(bytes));
});

test('saving an unchanged model still applies the writer’s relationship ID updates', async () => {
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const href = 'https://example.test/';
  const source = rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      `<Relationships xmlns="${relationships}"><Relationship Id="doc" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${word}" xmlns:r="${office}"><w:body><w:p>` +
      '<w:hyperlink r:id="second"><w:r><w:t>Link</w:t></w:r></w:hyperlink>' +
      '</w:p></w:body></w:document>'
    )],
    ['word/_rels/document.xml.rels', toBytes(
      `<Relationships xmlns="${relationships}">` +
      `<Relationship Id="first" Type="${office}/hyperlink" Target="${href}" TargetMode="External"/>` +
      `<Relationship Id="second" Type="${office}/hyperlink" Target="${href}" TargetMode="External"/>` +
      '</Relationships>'
    )],
  ]));
  const document = await parseDocx(source, { preloadFonts: false });
  const paragraph = document.package.document.content[0];
  const link = paragraph?.type === 'paragraph' ? paragraph.content[0] : undefined;
  if (link?.type !== 'hyperlink') throw new Error('Missing the external hyperlink');

  expect(link.rId).toBe('second');
  const saved = await repackDocx(document, WITHOUT_STAMP);
  expect(new Uint8Array(saved)).toEqual(new Uint8Array(source));
  expect(link.rId).toBe('first');
});

test('the default stamp changes only the core properties of an unchanged model', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  setSystemTime(new Date(NOW));
  try {
    const saved = await repackDocxWithWarnings(document, { modifiedBy: 'Fixture author' });
    expect(saved.warnings).toEqual(document.warnings ?? []);
    const before = unzipContainer(bytes);
    const after = unzipContainer(new Uint8Array(saved.buffer));
    const expectedCore = new TextDecoder().decode(before[CORE])
      .replace('2024-01-02T03:04:05Z</dcterms:modified>', `${NOW}</dcterms:modified>`)
      .replace('</cp:coreProperties>',
        '<cp:lastModifiedBy>Fixture author</cp:lastModifiedBy></cp:coreProperties>');
    expect(after).toEqual({ ...before, [CORE]: new TextEncoder().encode(expectedCore) });
    expect(after[CORE]).not.toEqual(before[CORE]);
  } finally {
    setSystemTime();
  }
});

test('a changed model uses the plan-less writer and reports the rewritten parts', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  const paragraph = document.package.document.content.find(
    (block) => block.type === 'paragraph' && block.paraId === EDIT_PARAGRAPH_ID
  );
  if (paragraph?.type !== 'paragraph' || paragraph.content[0]?.type !== 'run' ||
      paragraph.content[0].content[0]?.type !== 'text') {
    throw new Error('Missing the editable paragraph');
  }
  paragraph.content[0].content[0].text = 'Edited model';

  setSystemTime(new Date(NOW));
  try {
    const expected = writeDocxS13WithWarningsWire(JSON.stringify({
      determinism: { seed: createHash('sha256').update(bytes).digest('hex'), now: NOW },
      ...JSON.parse(serializeDocumentSaveModel(document)),
      options: WITHOUT_STAMP,
    }), bytes);
    const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
    expect(new Uint8Array(saved.buffer)).toEqual(Uint8Array.from(expected.bytes));
    const before = unzipContainer(bytes);
    const after = unzipContainer(expected.bytes);
    const rewritten = Object.entries(after)
      .filter(([path, content]) => !before[path] || !Buffer.from(content).equals(Buffer.from(before[path])))
      .map(([path]) => path)
      .sort();
    expect(rewritten).toContain('word/document.xml');
    expect(saved.warnings).toEqual([
      ...(document.warnings ?? []),
      ...expected.warnings,
      `${CODE} ${JSON.stringify(rewritten)} Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`,
    ]);
  } finally {
    setSystemTime();
  }
});

test('a changed Yrs projection is saved as a caller model', async () => {
  const { bytes } = oneSavePackage();
  const host = await parse(bytes);
  const session = await createYrsSession({ clientId: 51203 });
  try {
    session.seedFromDocx(bytes);
    const projected = yrsToDocument(session, host);
    projected.package.document.content.push({
      type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Projected edit' }] }],
    });
    const saved = await repackDocxWithWarnings(projected, WITHOUT_STAMP);
    expect(saved.warnings.some((warning) => warning.startsWith(CODE))).toBe(true);
    const document = unzipContainer(new Uint8Array(saved.buffer))['word/document.xml'];
    expect(new TextDecoder().decode(document)).toContain('Projected edit');
  } finally {
    session.destroy();
  }
});

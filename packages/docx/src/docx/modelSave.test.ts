import { beforeAll, expect, setSystemTime, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadOpcWasm, unzipContainer } from './wasm';
import { parseDocx } from './parser';
import { repackDocx, repackDocxWithWarnings } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { modelPartsRewrittenWarning } from './modelSave';
import { writeDocumentWithRust } from './rustSaveFacade';
import { EDIT_PARAGRAPH_ID, oneSavePackage } from '../yrs/__fixtures__/oneSavePackage';
import { createYrsSession } from '../yrs/index';
import { yrsToDocument } from '../yrs/yrsToDocument';
import type { Document } from '../types/document';

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
  const plain = await writeDocumentWithRust(await parse(source), document.originalBuffer!, WITHOUT_STAMP);
  const result = await repackDocxWithWarnings(document, WITHOUT_STAMP);
  expect(new Uint8Array(result.buffer)).toEqual(Uint8Array.from(source));
  expect(result.warnings).toEqual(plain.warnings);
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

test('the default stamp returns the plain writer output for an unchanged model', async () => {
  const { bytes } = oneSavePackage();
  const parts = new Map(Object.entries(unzipContainer(bytes)));
  const opaque = toBytes('opaque member');
  parts.set('__proto__', opaque);
  const source = rezipPartsToArrayBuffer(parts);
  const document = await parseDocx(source, { preloadFonts: false });
  setSystemTime(new Date(NOW));
  try {
    const plain = await writeDocumentWithRust(await parse(new Uint8Array(source)), source);
    const saved = await repackDocxWithWarnings(document);
    expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(plain.buffer));
    const after = unzipContainer(new Uint8Array(saved.buffer));
    expect(new TextDecoder().decode(after[CORE])).toContain(`${NOW}</dcterms:modified>`);
    expect(Object.getPrototypeOf(after)).toEqual(opaque);
    expect(saved.warnings.slice(0, -1)).toEqual(plain.warnings);
    expect(saved.warnings.at(-1)).toContain(`"${CORE}"`);
  } finally {
    setSystemTime();
  }
});

test('modifiedBy uses the plain writer with timestamp stamping disabled', async () => {
  const { bytes } = oneSavePackage();
  const parts = new Map(Object.entries(unzipContainer(bytes)));
  const core = new TextDecoder().decode(parts.get(CORE)).replace(
    '</cp:coreProperties>', '<cp:lastModifiedBy >Before</cp:lastModifiedBy></cp:coreProperties>'
  );
  parts.set(CORE, toBytes(core));
  const source = rezipPartsToArrayBuffer(parts);
  const writerOptions = { ...WITHOUT_STAMP, modifiedBy: 'Alice' };
  const plain = await writeDocumentWithRust(await parse(new Uint8Array(source)), source, writerOptions);
  let reads = 0;
  const saved = await repackDocxWithWarnings(await parse(new Uint8Array(source)), {
    ...WITHOUT_STAMP,
    get modifiedBy() {
      if (++reads > 2) throw new Error('Unexpected modifiedBy read');
      return 'Alice';
    },
  });
  expect(reads).toBe(2);
  expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(plain.buffer));
  const writtenCore = new TextDecoder().decode(unzipContainer(new Uint8Array(saved.buffer))[CORE]);
  expect(writtenCore).toContain('<cp:lastModifiedBy>Alice</cp:lastModifiedBy>');
  expect(writtenCore).toContain('2024-01-02T03:04:05Z</dcterms:modified>');
});

test('an unchanged model keeps writer warnings when returning the source bytes', async () => {
  const { bytes } = oneSavePackage();
  const parts = new Map(Object.entries(unzipContainer(bytes)));
  const path = 'word/document.xml';
  const xml = new TextDecoder().decode(parts.get(path)).replace(
    '</w:r>',
    '<w:object xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<o:OLEObject r:id="rIdOle"/></w:object></w:r>'
  );
  parts.set(path, toBytes(xml));
  const source = rezipPartsToArrayBuffer(parts);
  const plain = await writeDocumentWithRust(await parse(new Uint8Array(source)), source, WITHOUT_STAMP);
  const saved = await repackDocxWithWarnings(await parse(new Uint8Array(source)), WITHOUT_STAMP);
  expect(plain.warnings.join('\n')).toContain('opaque drawing references relationship "rIdOle"');
  expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(source));
  expect(saved.warnings).toEqual(plain.warnings);
});

function editableRun(document: Document) {
  const paragraph = document.package.document.content.find(
    (block) => block.type === 'paragraph' && block.paraId === EDIT_PARAGRAPH_ID
  );
  const run = paragraph?.type === 'paragraph' ? paragraph.content[0] : undefined;
  if (run?.type !== 'run' || run.content[0]?.type !== 'text') {
    throw new Error('Missing the editable paragraph');
  }
  return { run, textNode: run.content[0] };
}

test('model matching uses the value the writer serializes from a text proxy', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  const { run, textNode } = editableRun(document);
  run.content[0] = new Proxy(textNode, {
    get(target, key, receiver) {
      if (key === 'text') return 'Edited by proxy';
      return Reflect.get(target, key, receiver);
    },
  });
  const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
  const xml = unzipContainer(new Uint8Array(saved.buffer))['word/document.xml'];
  expect(new TextDecoder().decode(xml)).toContain('Edited by proxy');
  expect(saved.warnings.some((warning) => warning.startsWith(CODE))).toBe(true);
});

test('a changed model uses the plan-less writer and reports the rewritten parts', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  const plainDocument = await parse(bytes);
  editableRun(document).textNode.text = 'Edited model';
  editableRun(plainDocument).textNode.text = 'Edited model';

  setSystemTime(new Date(NOW));
  try {
    const plain = await writeDocumentWithRust(plainDocument, plainDocument.originalBuffer!);
    const saved = await repackDocxWithWarnings(document);
    expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(plain.buffer));
    const before = unzipContainer(bytes);
    const after = unzipContainer(new Uint8Array(plain.buffer));
    const rewritten = [...new Set([...Object.keys(before), ...Object.keys(after)])]
      .filter((path) => !before[path] || !after[path] || !Buffer.from(before[path]).equals(Buffer.from(after[path])))
      .sort();
    expect(rewritten).toContain('word/document.xml');
    expect(rewritten).toContain(CORE);
    expect(saved.warnings).toEqual([
      ...plain.warnings,
      `${CODE} ${JSON.stringify(rewritten)} Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`,
    ]);
  } finally {
    setSystemTime();
  }
});

test('the rewritten parts warning includes additions and removals', () => {
  const bytes = toBytes('member');
  const original = new Map([['removed.xml', bytes], ['unchanged.xml', bytes]]);
  const written = new Map([['added.xml', bytes], ['unchanged.xml', bytes.slice()]]);
  expect(modelPartsRewrittenWarning(original, written)).toBe(
    `${CODE} ["added.xml","removed.xml"] Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`
  );
  expect(modelPartsRewrittenWarning(original, original)).toBeUndefined();
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

import { beforeAll, expect, setSystemTime, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadOpcWasm, unzipContainer } from './wasm';
import { parseDocx } from './parser';
import { createEmptyDocx, repackDocx, repackDocxWithWarnings } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { modelPartsRewrittenWarning, readModelSaveDirectory } from './modelSave';
import { writeDocumentWithRust } from './rustSaveFacade';
import { EDIT_PARAGRAPH_ID, oneSavePackage } from '../yrs/__fixtures__/oneSavePackage';
import { createYrsSession } from '../yrs/index';
import { yrsToDocument } from '../yrs/yrsToDocument';
import type { Document } from '../types/document';

const ROOT = resolve(import.meta.dir, '../../../..');
const CORE = 'docProps/core.xml';
const CODE = 'DOCX_MODEL_PARTS_REWRITTEN';
const UNAVAILABLE_WARNING = `${CODE} (part names unavailable) Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`;
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

test('unchanged model saves byte for byte or uses writer output for repository documents', async () => {
  let reusedSource = 0;
  for (const path of fixturePaths.sort()) {
    const source = readFileSync(resolve(ROOT, path));
    const document = await parse(source);
    const plain = await writeDocumentWithRust(await parse(source), document.originalBuffer!, WITHOUT_STAMP);
    const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
    const bytes = Buffer.from(saved.buffer);

    if (bytes.equals(source)) {
      expect(saved.warnings, path).toEqual(plain.warnings);
      expect(saved.warnings.some((warning) => warning.startsWith(CODE)), path).toBe(false);
      reusedSource++;
    } else {
      expect(bytes.equals(Buffer.from(plain.buffer)), `${path} differs from the plain writer`).toBe(true);
      const warning = modelPartsRewrittenWarning(document.originalBuffer!, plain.buffer);
      if (!warning?.startsWith(`${CODE} [`)) {
        throw new Error(`${path} has no named rewritten parts`);
      }
      expect(saved.warnings, path).toEqual([...plain.warnings, warning]);
      console.log(`${path}: ${warning}`);
    }
  }
  expect(reusedSource, 'repository documents that reused source bytes').toBeGreaterThan(0);
  expect(reusedSource, `source reuse across ${fixturePaths.length} repository documents`)
    .toBeGreaterThanOrEqual(Math.ceil(fixturePaths.length * 0.9));
}, 30_000 * fixturePaths.length);

test('unchanged synthetic model saves byte for byte with stamping disabled', async () => {
  const { bytes } = oneSavePackage();
  await expectUnchanged(bytes);
  expect(new Uint8Array(await repackDocx(await parse(bytes), WITHOUT_STAMP)))
    .toEqual(Uint8Array.from(bytes));
});

test('an unchanged save reads the source once and copies its bytes', async () => {
  class SourceBuffer extends ArrayBuffer {
    override slice(): ArrayBuffer {
      return new ArrayBuffer(0);
    }
  }

  const { bytes } = oneSavePackage();
  const source = new SourceBuffer(bytes.byteLength);
  new Uint8Array(source).set(bytes);
  const document = await parseDocx(source, { preloadFonts: false });
  let reads = 0;
  Object.defineProperty(document, 'originalBuffer', {
    get() {
      reads++;
      return source;
    },
  });
  const saved = await repackDocx(document, WITHOUT_STAMP);
  expect(reads).toBe(1);
  expect(saved).not.toBe(source);
  expect(new Uint8Array(saved)).toEqual(Uint8Array.from(bytes));
});

const OFFICE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HYPERLINK_TARGET = 'https://example.test/';
const DOCUMENT_RELS = 'word/_rels/document.xml.rels';

function bodyHyperlinkPackage(id: string, linkRelationships: string): ArrayBuffer {
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
  return rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      `<Relationships xmlns="${relationships}"><Relationship Id="doc" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/></Relationships>`
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${word}" xmlns:r="${OFFICE_RELATIONSHIPS}"><w:body><w:p>` +
      `<w:hyperlink r:id="${id}"><w:r><w:t>Link</w:t></w:r></w:hyperlink>` +
      '</w:p></w:body></w:document>'
    )],
    [DOCUMENT_RELS, toBytes(
      `<Relationships xmlns="${relationships}">${linkRelationships}</Relationships>`
    )],
  ]));
}

function bodyHyperlink(document: Document) {
  const paragraph = document.package.document.content[0];
  const link = paragraph?.type === 'paragraph' ? paragraph.content[0] : undefined;
  if (link?.type !== 'hyperlink') throw new Error('Missing the external hyperlink');
  return link;
}

test('saving an unchanged model still applies the writer’s relationship ID updates', async () => {
  const source = bodyHyperlinkPackage('second',
    `<Relationship Id="first" Type="${OFFICE_RELATIONSHIPS}/hyperlink" Target="${HYPERLINK_TARGET}" TargetMode="External"/>` +
    `<Relationship Id="second" Type="${OFFICE_RELATIONSHIPS}/hyperlink" Target="${HYPERLINK_TARGET}" TargetMode="External"/>`
  );
  const document = await parseDocx(source, { preloadFonts: false });
  const link = bodyHyperlink(document);

  expect(link.rId).toBe('second');
  const saved = await repackDocx(document, WITHOUT_STAMP);
  expect(new Uint8Array(saved)).toEqual(new Uint8Array(source));
  expect(link.rId).toBe('first');
});

test('an unchanged hyperlink with single-quoted relationship attributes saves the writer output', async () => {
  const source = bodyHyperlinkPackage('rId5',
    `<Relationship Id='rId5' Type='${OFFICE_RELATIONSHIPS}/hyperlink' Target='${HYPERLINK_TARGET}' TargetMode='External'/>`
  );
  const document = await parseDocx(source, { preloadFonts: false });
  const link = bodyHyperlink(document);
  expect(link.rId).toBe('rId5');
  expect(link.href).toBe(HYPERLINK_TARGET);

  setSystemTime(new Date(NOW));
  try {
    const plain = await writeDocumentWithRust(await parse(new Uint8Array(source)), source, WITHOUT_STAMP);
    const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
    expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(plain.buffer));
    expect(link.rId).toBe('rId1');
    const written = await parse(new Uint8Array(saved.buffer));
    expect(written.package.relationships?.get(link.rId!)?.target).toBe(HYPERLINK_TARGET);
    const warning = modelPartsRewrittenWarning(source, plain.buffer);
    expect(warning).toContain(`"${DOCUMENT_RELS}"`);
    expect(saved.warnings).toEqual([...plain.warnings, warning!]);
  } finally {
    setSystemTime();
  }
});

test('the default stamp returns the plain writer output with opaque ZIP members', async () => {
  const { bytes } = oneSavePackage();
  const parts = new Map(Object.entries(unzipContainer(bytes)));
  const opaque = toBytes('opaque member');
  parts.set('__proto__', opaque);
  parts.set('length', toBytes('another opaque member'));
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
    const warning = modelPartsRewrittenWarning(source, plain.buffer);
    expect(warning).toContain(`"${CORE}"`);
    expect(warning).not.toContain('"__proto__"');
    expect(warning).not.toContain('"length"');
    expect(saved.warnings).toEqual([...plain.warnings, warning!]);
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
      reads++;
      return reads === 1 ? 'Before' : 'Alice';
    },
  });
  expect(reads).toBe(2);
  expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(plain.buffer));
  const writtenCore = new TextDecoder().decode(unzipContainer(new Uint8Array(saved.buffer))[CORE]);
  expect(writtenCore).toContain('<cp:lastModifiedBy>Alice</cp:lastModifiedBy>');
  expect(writtenCore).toContain('2024-01-02T03:04:05Z</dcterms:modified>');
});

async function sourceWithModifier(modifiedBy: string): Promise<ArrayBuffer> {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  const { buffer } = await writeDocumentWithRust(document, document.originalBuffer!, {
    ...WITHOUT_STAMP,
    modifiedBy,
  });
  return buffer;
}

test('a proxy modifiedBy value reaches the writer and changes the saved core', async () => {
  const source = await sourceWithModifier('Before');
  const options = new Proxy({ ...WITHOUT_STAMP, modifiedBy: 'Before' }, {
    get(target, key, receiver) {
      return key === 'modifiedBy' ? 'Alice' : Reflect.get(target, key, receiver);
    },
  });
  const plain = await writeDocumentWithRust(await parse(new Uint8Array(source)), source, options);
  const saved = await repackDocx(await parse(new Uint8Array(source)), options);
  expect(new Uint8Array(saved)).toEqual(new Uint8Array(plain.buffer));
  expect(new Uint8Array(saved)).not.toEqual(new Uint8Array(source));
  const core = unzipContainer(new Uint8Array(saved))[CORE];
  expect(new TextDecoder().decode(core)).toContain('<cp:lastModifiedBy>Alice</cp:lastModifiedBy>');
});

test('the warning names a changed core with the same CRC and size', async () => {
  const source = await sourceWithModifier('plumless');
  const saved = await repackDocxWithWarnings(await parse(new Uint8Array(source)), {
    ...WITHOUT_STAMP,
    modifiedBy: 'buckeroo',
  });
  const before = readModelSaveDirectory(source)?.get(CORE);
  const after = readModelSaveDirectory(saved.buffer)?.get(CORE);
  expect(before).toBeDefined();
  expect(after?.crc32).toBe(before?.crc32);
  expect(after?.uncompressedSize).toBe(before?.uncompressedSize);
  const core = unzipContainer(new Uint8Array(saved.buffer))[CORE];
  expect(new TextDecoder().decode(core)).toContain('<cp:lastModifiedBy>buckeroo</cp:lastModifiedBy>');
  expect(saved.warnings.find((warning) => warning.startsWith(CODE))).toContain(`"${CORE}"`);
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

test('a timestamp getter cannot discard the body captured by the writer', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  const original = document.package.document;
  document.package.document = { ...original, content: structuredClone(original.content) };
  const { textNode } = editableRun(document);
  textNode.text = 'Edited body before the save';
  let reads = 0;
  const saved = await repackDocx(document, {
    get updateModifiedDate() {
      reads++;
      document.package.document = original;
      return false;
    },
  });
  expect(reads).toBe(1);
  const xml = unzipContainer(new Uint8Array(saved))['word/document.xml'];
  expect(new TextDecoder().decode(xml)).toContain('Edited body before the save');
});

test('the completed write survives a warnings iterator detaching the source', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  const plainDocument = await parse(bytes);
  editableRun(document).textNode.text = 'Edited model';
  editableRun(plainDocument).textNode.text = 'Edited model';
  const source = document.originalBuffer!;
  document.warnings = [];
  Object.defineProperty(document.warnings, Symbol.iterator, {
    value: function* () {
      structuredClone(source, { transfer: [source] });
      yield 'Caller warning';
    },
  });
  plainDocument.warnings = ['Caller warning'];

  const plain = await writeDocumentWithRust(plainDocument, plainDocument.originalBuffer!, WITHOUT_STAMP);
  const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
  expect(source.byteLength).toBe(0);
  expect(new Uint8Array(saved.buffer)).toEqual(new Uint8Array(plain.buffer));
  expect(saved.warnings).toEqual([
    ...plain.warnings,
    modelPartsRewrittenWarning(plainDocument.originalBuffer!, plain.buffer)!,
  ]);
});

test('a completed write reports unavailable part names when ZIP inspection throws', async () => {
  const { bytes } = oneSavePackage();
  const document = await parse(bytes);
  editableRun(document).textNode.text = 'Edited model';
  const Decoder = globalThis.TextDecoder;
  document.warnings = [];
  Object.defineProperty(document.warnings, Symbol.iterator, {
    value: function* () {
      globalThis.TextDecoder = class extends Decoder {
        constructor() {
          super();
          throw new Error('ZIP inspection failed');
        }
      };
      yield 'Caller warning';
    },
  });

  const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP).finally(() => {
    globalThis.TextDecoder = Decoder;
  });
  const xml = unzipContainer(new Uint8Array(saved.buffer))['word/document.xml'];
  expect(new TextDecoder().decode(xml)).toContain('Edited model');
  expect(saved.warnings).toContain('Caller warning');
  expect(saved.warnings.at(-1)).toBe(UNAVAILABLE_WARNING);
});

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

test('model matching preserves the writer clearing boxed NaN paragraph formatting', async () => {
  const parts = new Map(Object.entries(unzipContainer(new Uint8Array(await createEmptyDocx()))));
  const path = 'word/document.xml';
  const xml = new TextDecoder().decode(parts.get(path)).replace(
    '<w:p>', '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
  );
  parts.set(path, toBytes(xml));
  const document = await parse(new Uint8Array(rezipPartsToArrayBuffer(parts)));
  const paragraph = document.package.document.content[0];
  if (paragraph?.type !== 'paragraph') throw new Error('Missing the centered paragraph');
  expect(paragraph.formatting?.alignment).toBe('center');
  paragraph.formatting = Object.assign(
    Object.setPrototypeOf(new Number(NaN), Object.prototype),
    paragraph.formatting
  );

  const saved = await repackDocx(document, WITHOUT_STAMP);
  const written = unzipContainer(new Uint8Array(saved))[path];
  expect(new TextDecoder().decode(written)).not.toContain('<w:jc');
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

test('the rewritten parts warning includes additions, removals and changed contents', () => {
  const bytes = toBytes('member');
  const original = rezipPartsToArrayBuffer(new Map([
    ['removed.xml', bytes], ['unchanged.xml', bytes], ['changed.xml', bytes],
  ]));
  const written = rezipPartsToArrayBuffer(new Map([
    ['added.xml', bytes], ['unchanged.xml', bytes.slice()], ['changed.xml', toBytes('edited')],
  ]));
  expect(modelPartsRewrittenWarning(original, written)).toBe(
    `${CODE} ["added.xml","changed.xml","removed.xml"] Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`
  );
  expect(modelPartsRewrittenWarning(original, original)).toBeUndefined();
  expect(modelPartsRewrittenWarning(new ArrayBuffer(1), new ArrayBuffer(1))).toBeUndefined();
  expect(modelPartsRewrittenWarning(new ArrayBuffer(1), Uint8Array.of(1).buffer)).toBe(UNAVAILABLE_WARNING);
  expect(modelPartsRewrittenWarning(original, new ArrayBuffer(0))).toBe(UNAVAILABLE_WARNING);
});

function toZip64(bytes: Uint8Array): Uint8Array {
  const source = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const footer = bytes.byteLength - 22;
  const count = source.getUint16(footer + 10, true);
  const central = source.getUint32(footer + 16, true);
  const zip64ExtraSize = 28;
  const zip64 = new Uint8Array(bytes.byteLength + count * zip64ExtraSize + 56 + 20);
  const output = new DataView(zip64.buffer);
  zip64.set(bytes.subarray(0, central));
  let from = central;
  let to = central;
  for (let index = 0; index < count; index++) {
    const extra = from + 46 + source.getUint16(from + 28, true);
    const extraLength = source.getUint16(from + 30, true);
    const next = extra + extraLength + source.getUint16(from + 32, true);
    const local = source.getUint32(from + 42, true);
    zip64.set(bytes.subarray(from, extra), to);
    output.setUint16(local + 4, 45, true);
    output.setUint16(to + 6, 45, true);
    output.setUint32(to + 20, 0xffffffff, true);
    output.setUint32(to + 24, 0xffffffff, true);
    output.setUint16(to + 30, extraLength + zip64ExtraSize, true);
    output.setUint32(to + 42, 0xffffffff, true);
    const extraAt = to + extra - from;
    output.setUint16(extraAt, 1, true);
    output.setUint16(extraAt + 2, 24, true);
    output.setBigUint64(extraAt + 4, BigInt(source.getUint32(from + 24, true)), true);
    output.setBigUint64(extraAt + 12, BigInt(source.getUint32(from + 20, true)), true);
    output.setBigUint64(extraAt + 20, BigInt(local), true);
    zip64.set(bytes.subarray(extra, next), extraAt + zip64ExtraSize);
    to += next - from + zip64ExtraSize;
    from = next;
  }

  const locator = to + 56;
  output.setUint32(to, 0x06064b50, true);
  output.setBigUint64(to + 4, 44n, true);
  output.setUint16(to + 12, 45, true);
  output.setUint16(to + 14, 45, true);
  output.setBigUint64(to + 24, BigInt(count), true);
  output.setBigUint64(to + 32, BigInt(count), true);
  output.setBigUint64(to + 40, BigInt(to - central), true);
  output.setBigUint64(to + 48, BigInt(central), true);
  output.setUint32(locator, 0x07064b50, true);
  output.setBigUint64(locator + 8, BigInt(to), true);
  output.setUint32(locator + 16, 1, true);
  const end = locator + 20;
  zip64.set(bytes.subarray(footer), end);
  output.setUint16(end + 8, 0xffff, true);
  output.setUint16(end + 10, 0xffff, true);
  output.setUint32(end + 12, 0xffffffff, true);
  output.setUint32(end + 16, 0xffffffff, true);
  return zip64;
}

test('a ZIP64 source names the edited member in the warning', async () => {
  const { bytes } = oneSavePackage();
  const source = toZip64(bytes);
  expect(new Uint8Array(await repackDocx(await parse(source), WITHOUT_STAMP))).toEqual(new Uint8Array(source));
  const document = await parse(source);
  editableRun(document).textNode.text = 'Edited ZIP64 model';
  const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
  const warning = saved.warnings.find((entry) => entry.startsWith(CODE));
  expect(warning).toContain('"word/document.xml"');
  expect(warning).not.toContain(`"${CORE}"`);
});

test.each(['__proto__', 'length', 'word/media/'])('the warning accounts for the ZIP member %s', async (name) => {
  const { bytes } = oneSavePackage();
  const parts = new Map(Object.entries(unzipContainer(bytes)));
  parts.set(name, new Uint8Array());
  const source = rezipPartsToArrayBuffer(parts);
  const document = await parseDocx(source, { preloadFonts: false });
  editableRun(document).textNode.text = 'Edited model';
  const saved = await repackDocxWithWarnings(document, WITHOUT_STAMP);
  const before = readModelSaveDirectory(source);
  const after = readModelSaveDirectory(saved.buffer);
  const warning = saved.warnings.at(-1);
  expect(before?.has(name)).toBe(true);
  expect(warning).toContain('"word/document.xml"');
  if (name === 'word/media/') {
    expect(after?.has(name)).toBe(false);
    expect(warning).toContain(`"${name}"`);
  } else {
    expect(after?.has(name)).toBe(true);
    expect(after?.get(name)?.crc32).toBe(before?.get(name)?.crc32);
    expect(after?.get(name)?.uncompressedSize).toBe(before?.get(name)?.uncompressedSize);
    expect(warning).not.toContain(`"${name}"`);
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

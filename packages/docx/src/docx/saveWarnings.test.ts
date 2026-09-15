import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from './parser';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from '../yrs/documentToYrs';
import { createYrsSession } from '../yrs/index';
import { yrsToDocument } from '../yrs/yrsToDocument';
import type { Chart, Document } from '../types/document';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { repackDocxWithWarnings } from './rezip';
import { writeDocumentWithRust } from './rustSaveFacade';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"`;

function fixture(): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart1" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`);
  set('word/document.xml', `<w:document ${NAMESPACES}><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function unrecoverableChartDocument(): Document {
  const chart: Chart = {
    type: 'chart',
    chartType: 'column',
    rId: 'rIdNope',
    path: 'word/charts/chartNope.xml',
    series: [],
  };
  return {
    originalBuffer: fixture().buffer as ArrayBuffer,
    package: {
      document: {
        content: [
          { type: 'paragraph', paraId: '00000001', content: [{ type: 'run', content: [{ type: 'text', text: 'Text' }] }] },
          { type: 'paragraph', paraId: '00000002', content: [{ type: 'run', content: [{ type: 'chart', chart }] }] },
        ],
      },
    },
  };
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))));

it('production save results surface dropped chart runs instead of swallowing them', async () => {
  const document = unrecoverableChartDocument();
  const session = await createYrsSession({ clientId: 74510 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document);
    expect(saved.warnings?.join('\n')).toContain('rIdNope');
    const repacked = await repackDocxWithWarnings(saved);
    expect(repacked.warnings.join('\n')).toContain('rIdNope');
    expect(repacked.buffer.byteLength).toBeGreaterThan(0);
    const rust = await writeDocumentWithRust(saved, fixture().buffer as ArrayBuffer);
    expect(rust.buffer.byteLength).toBeGreaterThan(0);
    expect(rust.warnings.join('\n')).toContain('rIdNope');
  } finally {
    session.destroy();
  }
});

it('production saves warn about opaque drawings referencing missing relationships', async () => {
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`);
  set('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  set('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
  set('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p><w:r><w:object><o:OLEObject Type="Embed" ProgID="Eq" r:id="rIdOle"/></w:object></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`);
  const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId: 74512 });
  try {
    documentToYrs(session, parsed);
    const saved = yrsToDocument(session, parsed);
    const result = await writeDocumentWithRust(saved, bytes.buffer as ArrayBuffer);
    expect(result.warnings.join('\n')).toContain('rIdOle');
  } finally {
    session.destroy();
  }
});

it('production save results carry caller-visible document warnings through', async () => {
  const document = unrecoverableChartDocument();
  const session = await createYrsSession({ clientId: 74511 });
  try {
    documentToYrs(session, document);
    const saved = yrsToDocument(session, document);
    saved.warnings = [...(saved.warnings ?? []), 'caller note'];
    const repacked = await repackDocxWithWarnings(saved);
    expect(repacked.warnings).toContain('caller note');
    const rust = await writeDocumentWithRust(saved, fixture().buffer as ArrayBuffer);
    expect(rust.warnings).toContain('caller note');
  } finally {
    session.destroy();
  }
});

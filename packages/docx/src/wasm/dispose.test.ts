import { beforeAll } from 'bun:test';
import { readFileSync } from 'node:fs';
import { describeDisposeTable } from '../../../../bindings/dispose-cases';
import { preloadEditWasm } from './edit';
import * as engine from './engineEdit';

const font = new Uint8Array(readFileSync(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
let docx: Uint8Array;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('./generated/edit/docx_edit_bg.wasm', import.meta.url))));
  const encoder = new TextEncoder();
  docx = engine.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
  });
}, 30_000);

function retainedMeta(): engine.RetainedLayoutMeta {
  const session = new engine.EditSession(91);
  try {
    session.open_docx(docx, true, 'opening');
    const face = session.register_measure_font(font);
    return session.layout_document_with_regions_retained_meta(JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {
        pageWidth: 4320, pageHeight: 2880, marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
      } }] },
      measurement: { fontChains: { 'calibri|0|0': [face] }, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
      renderEnv: {},
    }));
  } finally {
    session.free();
  }
}

function runInheritance(): engine.RunInheritance {
  const session = new engine.EditSession(92);
  try {
    return session.run_inheritance('null');
  } finally {
    session.free();
  }
}

describeDisposeTable(engine.engineKind, engine, [
  { name: 'EditSession', engineClass: engine.EditSession, make: () => new engine.EditSession(7), use: (session) => (session as engine.EditSession).client_id() },
  { name: 'RetainedLayoutMeta', engineClass: engine.RetainedLayoutMeta, make: retainedMeta, use: (meta) => (meta as engine.RetainedLayoutMeta).page_count },
  { name: 'RunInheritance', engineClass: engine.RunInheritance, make: runInheritance, use: (runs) => (runs as engine.RunInheritance).attributes() },
]);

import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/></Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const SECTION = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;

function fixture(body: string): Uint8Array {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}${SECTION}</w:body></w:document>`;
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(document));
  parts.set('word/styles.xml', toBytes(STYLES));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

/** The saved `w:body` up to its `w:sectPr`. */
function savedBody(buffer: ArrayBuffer): string {
  const xml = new TextDecoder().decode(
    unzipContainer(new Uint8Array(buffer))['word/document.xml'] as Uint8Array
  );
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const end = xml.indexOf('<w:sectPr');
  return xml.slice(start, end === -1 ? xml.indexOf('</w:body>') : end);
}

const PAGE_BREAK = '<w:r><w:br w:type="page"/></w:r>';
let clientId = 68100;

/** `[through the engine, straight through the serializer]`. */
async function save(
  body: string,
  edit?: (session: YrsSession) => void
): Promise<[string, string]> {
  const bytes = fixture(body);
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const control = savedBody(await repackDocx(parsed));

  const session = await createYrsSession({ clientId: (clientId += 1) });
  try {
    session.seedFromDocx(bytes);
    edit?.(session);
    return [savedBody(await repackDocx(yrsToDocument(session, parsed))), control];
  } finally {
    session.destroy();
  }
}

describe('page break save projection', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  const shapes: Array<[string, string, number]> = [
    [
      'a break between two runs',
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`,
      1,
    ],
    [
      'a break closing a paragraph',
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`,
      1,
    ],
    ['a break opening a paragraph', `<w:p>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`, 1],
    [
      'two breaks in one paragraph',
      `<w:p><w:r><w:t>A</w:t></w:r>${PAGE_BREAK}<w:r><w:t>B</w:t></w:r>${PAGE_BREAK}<w:r><w:t>C</w:t></w:r></w:p>`,
      2,
    ],
    [
      'a break owning its paragraph',
      `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p>${PAGE_BREAK}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p>`,
      1,
    ],
    [
      'a column break',
      `<w:p><w:r><w:t>HEAD</w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>TARGET</w:t></w:r></w:p>`,
      0,
    ],
  ];

  for (const [name, body, pageBreaks] of shapes) {
    it(`saves ${name} as the serializer would`, async () => {
      const [through, control] = await save(body);
      expect(through).toBe(control);
      expect(through.match(/<w:br w:type="page"\/>/g) ?? []).toHaveLength(pageBreaks);
    });
  }

  it('saves a break through a paragraph whose formatting changed', async () => {
    const [through, control] = await save(
      `<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`,
      (session) => {
        const last = session.paragraphs('body').at(-1);
        if (last) session.setParagraphAttr(last.paraId, 'widowControl', false);
      }
    );
    expect(through).toContain('<w:br w:type="page"/>');
    expect(through).toContain('<w:widowControl w:val="0"/>');
    expect(control).toContain('<w:br w:type="page"/>');
  });

  it('records the same breaks through the projector and the engine', async () => {
    const bytes = fixture(
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`
    );
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const projected = await createYrsSession({ clientId: 68001 });
    const engine = await createYrsSession({ clientId: 68001 });
    try {
      documentToYrs(projected, parsed, { generation: 'parity' });
      engine.seedFromDocx(bytes, { generation: 'parity' });
      expect(engine.storySegments('body')).toEqual(projected.storySegments('body'));
      expect(engine.encodeState()).toEqual(projected.encodeState());
      expect(engine.paragraphs('body')[0]?.properties._originalRunBoundaries).toEqual([
        { text: 'HEAD', marksKey: '' },
        { text: '', breaks: [{ offset: 0, type: 'page' }] },
        { text: 'TARGET', marksKey: '' },
      ]);
    } finally {
      projected.destroy();
      engine.destroy();
    }
  });

  it('does not open a paragraph the source never had', async () => {
    const [through, control] = await save(
      `<w:p><w:r><w:t>HEAD</w:t></w:r>${PAGE_BREAK}<w:r><w:t>TARGET</w:t></w:r></w:p>`
    );
    expect(through.match(/<w:p[ >]/g) ?? []).toHaveLength(
      (control.match(/<w:p[ >]/g) ?? []).length
    );
  });
});

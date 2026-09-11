import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '@betteroffice/docx/docx/rezip/parts';
import { YrsPositionProjection } from './yrsPositionProjection';

const WASM = resolve(
  import.meta.dir,
  '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
);

const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

/** Paragraph, block embed, then two more paragraphs — drift starts at the second. */
function bodyDocx(embed: string): Uint8Array {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
    ${embed}
    <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
    <w:p><w:r><w:t>Charlie</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(document));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const TABLE = `<w:tbl>
  <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
  <w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`;

/** A break after visible text seeds a block-level `pageBreak` embed behind it. */
const PAGE_BREAK = `<w:p><w:r><w:t>Beta</w:t><w:br w:type="page"/></w:r></w:p>`;

const BLOCK_SDT = `<w:sdt>
  <w:sdtPr><w:alias w:val="Control"/><w:tag w:val="control"/></w:sdtPr>
  <w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent>
</w:sdt>`;

/** The pointer path: a display-list position to the yrs Loc typing inserts at. */
function clickToLoc(session: YrsSession, projection: YrsPositionProjection, position: number) {
  const target = projection.targetAt(position);
  const map = createYrsInputPositionMap(target.story, session.paragraphSpans(target.story));
  return displayPositionToYrsLoc(map, target.displayPosition);
}

function paragraphStart(projection: YrsPositionProjection, paraId: string): number {
  for (let position = 0; position <= projection.size; position += 1) {
    const node = projection.nodeAt(position);
    if (node?.kind === 'paragraph' && node.attrs.paraId === paraId) return position;
  }
  throw new Error(`no laid-out paragraph ${paraId}`);
}

describe('a body story carrying a block embed', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  for (const [name, embed] of [
    ['table', TABLE],
    ['page break', PAGE_BREAK],
    ['block content control', BLOCK_SDT],
  ] as const) {
    test(`types where the click landed after a ${name}`, async () => {
      const session = await createYrsSession({ clientId: 71001 });
      try {
        session.seedFromDocx(bodyDocx(embed));
        const projection = new YrsPositionProjection(session, 'body');
        const charlie = session.paragraphs('body').find((entry) => entry.text === 'Charlie');
        expect(charlie).toBeDefined();

        // click between "Cha" and "rlie"
        const clicked = paragraphStart(projection, charlie!.paraId) + 1 + 3;
        const loc = clickToLoc(session, projection, clicked);
        expect(loc).not.toBeNull();

        session.insertText(loc!, 'X');
        const typed = session.paragraphs('body').find((entry) => entry.paraId === charlie!.paraId);
        expect(typed?.text).toBe('ChaXrlie');
        expect(projection.positionForLoc(loc!)).toBe(clicked);
      } finally {
        session.destroy();
      }
    });
  }

  test('accumulates table, page, column, and block-SDT units before text', async () => {
    const session = await createYrsSession({ clientId: 71003 });
    try {
      session.seedFromDocx(bodyDocx(''));
      const bravo = session.paragraphs('body').find((entry) => entry.text === 'Bravo');
      expect(bravo).toBeDefined();
      const start = session.locateParagraph('body', bravo!.paraId).start;
      session.createStory('body:sdt-regression', 'control');
      session.applyRawOps('body', [
        { op: 'insertEmbed', index: start, kind: 'table', payload: { grid: [], rows: [] } },
        { op: 'insertEmbed', index: start + 1, kind: 'pageBreak' },
        { op: 'insertEmbed', index: start + 2, kind: 'columnBreak' },
        {
          op: 'insertEmbed',
          index: start + 3,
          kind: 'blockSdt',
          payload: { story: 'body:sdt-regression' },
        },
      ]);
      const projection = new YrsPositionProjection(session, 'body');
      const clicked = paragraphStart(projection, bravo!.paraId) + 4;
      const loc = clickToLoc(session, projection, clicked);

      expect(loc).toEqual({ story: 'body', paraId: bravo!.paraId, offset: 7 });
      expect(projection.positionForLoc(loc!)).toBe(clicked);
      session.insertText(loc!, 'X');
      expect(session.paragraphs('body').find((entry) => entry.paraId === bravo!.paraId)?.text).toBe(
        'BraXvo'
      );
    } finally {
      session.destroy();
    }
  });

  test('round-trips every clickable position back to the position clicked', async () => {
    const session = await createYrsSession({ clientId: 71002 });
    try {
      session.seedFromDocx(bodyDocx(TABLE));
      const projection = new YrsPositionProjection(session, 'body');
      const drifted: number[] = [];
      for (let start = 0; start <= projection.size; start += 1) {
        const node = projection.nodeAt(start);
        if (node?.kind !== 'paragraph') continue;
        for (let offset = 1; offset < node.nodeSize - 1; offset += 1) {
          const clicked = start + offset;
          const loc = clickToLoc(session, projection, clicked);
          if (loc && projection.positionForLoc(loc) !== clicked) drifted.push(clicked);
        }
      }
      expect(drifted).toEqual([]);
    } finally {
      session.destroy();
    }
  });
});

import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const paragraph = (text: string, ppr = '') =>
  `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (body: string, style = '') =>
  `<w:tbl><w:tblPr>${
    style ? `<w:tblStyle w:val="${style}"/>` : ''
  }</w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc>${body}</w:tc></w:tr></w:tbl>`;

function fixture(withDefaults: boolean, withControl = true): Uint8Array<ArrayBuffer> {
  const styles = `${
    withDefaults
      ? '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="279" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>'
      : ''
  }
<w:style w:type="paragraph" w:styleId="Spaced"><w:pPr><w:spacing w:after="100" w:line="360" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="table" w:styleId="Base"><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="table" w:styleId="Grid"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="Inner"><w:pPr><w:spacing w:after="60" w:line="300" w:lineRule="auto"/></w:pPr></w:style>`;
  const cell =
    paragraph('Table defaults') +
    paragraph('Paragraph style', '<w:pStyle w:val="Spaced"/>') +
    paragraph(
      'Direct formatting',
      '<w:pStyle w:val="Spaced"/><w:spacing w:after="200" w:line="480" w:lineRule="auto"/>'
    ) +
    (withControl
      ? `<w:sdt><w:sdtPr><w:id w:val="42"/><w:tag w:val="test"/></w:sdtPr><w:sdtContent>${paragraph(
          'Content control'
        )}</w:sdtContent></w:sdt>`
      : '') +
    table(paragraph('Nested table')) +
    paragraph('Outer restored');
  const body = paragraph('Before') + table(cell, 'Grid') + paragraph('After');
  const parts = new Map<string, Uint8Array>();
  const set = (name: string, xml: string) => parts.set(name, toBytes(xml));
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
  );
  set(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
  );
  set('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  set('word/styles.xml', `<w:styles xmlns:w="${W}">${styles}</w:styles>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function spacingByStory(session: YrsSession): Array<[string, unknown[]]> {
  return session.storyIds().map((story) => [
    story,
    session.paragraphs(story).map((paragraph) => ({
      spaceAfter: paragraph.properties.spaceAfter,
      lineSpacing: paragraph.properties.lineSpacing,
    })),
  ]);
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

for (const withDefaults of [true, false]) {
  it(`keeps table paragraph inheritance identical for both seeders (document defaults: ${withDefaults})`, async () => {
    const bytes = fixture(withDefaults);
    const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
    const original = structuredClone(parsed);
    const projected = await createYrsSession({ clientId: 74215 });
    const native = await createYrsSession({ clientId: 74215 });
    try {
      documentToYrs(projected, parsed);
      native.seedFromDocx(bytes);
      expect(projected.storyIds()).toEqual(native.storyIds());
      for (const story of native.storyIds()) {
        expect(projected.storySegments(story)).toEqual(native.storySegments(story));
        expect(projected.yrsBlocksForStory(story)).toEqual(native.yrsBlocksForStory(story));
      }
      const cellStory = 'body:t0:r0c0';
      expect(
        projected
          .paragraphs(cellStory)
          .map((p) => [p.properties.spaceAfter, p.properties.lineSpacing])
      ).toEqual([
        [0, 240],
        [100, 360],
        [200, 480],
        [0, 240],
      ]);
      expect(projected.paragraphs(`${cellStory}:sdt0`)[0].properties.spaceAfter).toBe(0);
      expect(projected.paragraphs(`${cellStory}:t0:r0c0`)[0].properties.spaceAfter).toBe(60);
      expect(parsed).toEqual(original);
    } finally {
      projected.destroy();
      native.destroy();
    }
  });
}

it('preserves table spacing through text edits and save/reopen for both seeders', async () => {
  const bytes = fixture(true, false);
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  for (const seeder of ['native', 'projected']) {
    const session = await createYrsSession({ clientId: 74217 });
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, parsed);
      const expectedSpacing = spacingByStory(session);
      const story = 'body:t0:r0c0';
      for (const edit of [false, true]) {
        if (edit) {
          session.insertText(
            { story, paraId: session.paragraphs(story)[0].paraId, offset: 0 },
            'Edited '
          );
        }
        const saved = await repackDocx(yrsToDocument(session, parsed));
        const reopened = await createYrsSession({ clientId: 74218 });
        try {
          reopened.seedFromDocx(new Uint8Array(saved));
          expect(spacingByStory(reopened)).toEqual(expectedSpacing);
          expect(reopened.paragraphs(story)[0].text).toBe(
            edit ? 'Edited Table defaults' : 'Table defaults'
          );
        } finally {
          reopened.destroy();
        }
      }
    } finally {
      session.destroy();
    }
  }
});

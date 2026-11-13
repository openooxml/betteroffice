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
  return documentFixture(body, styles);
}

function documentFixture(body: string, styles: string, theme?: string): Uint8Array<ArrayBuffer> {
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
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>${
      theme ? `<Relationship Id="theme" Type="${R}/theme" Target="theme/theme1.xml"/>` : ''
    }</Relationships>`
  );
  set('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  set('word/styles.xml', `<w:styles xmlns:w="${W}">${styles}</w:styles>`);
  if (theme) set('word/theme/theme1.xml', theme);
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
  ),
  30_000
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

it.each(['projected', 'native'])('applies conditional paragraph spacing (seeder: %s)', async (seeder) => {
  const fixturePath = resolve(import.meta.dir, '__fixtures__/table-conditional-spacing');
  const bytes = documentFixture(
    readFileSync(resolve(fixturePath, 'body.xml'), 'utf8'),
    readFileSync(resolve(fixturePath, 'styles.xml'), 'utf8')
  );
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const original = structuredClone(parsed);
  const projected = await createYrsSession({ clientId: 74219 });
  const native = await createYrsSession({ clientId: 74219 });
  try {
    documentToYrs(projected, parsed);
    if (seeder === 'native') {
      native.seedFromDocx(bytes);
      expect(projected.storyIds()).toEqual(native.storyIds());
      for (const story of native.storyIds()) {
        expect(projected.storySegments(story)).toEqual(native.storySegments(story));
        expect(projected.yrsBlocksForStory(story)).toEqual(native.yrsBlocksForStory(story));
      }
    }
    const expected = [
      [
        [190, 110, 110, 110, 200],
        [130, 170, 180, 170, 140],
        [130, 170, 180, 170, 140],
        [130, 170, 180, 170, 140],
        [210, 120, 120, 120, 220],
      ],
      [
        [110, 110, 110, 110, 110],
        [150, 150, 150, 150, 150],
        [150, 150, 150, 150, 150],
        [160, 160, 160, 160, 160],
        [160, 160, 160, 160, 160],
      ],
      Array.from({ length: 5 }, () => [10, 10, 10, 10, 10]),
      [
        [190, 110, 200],
        [170, 180, 170],
        [130, 180, 140],
        [210, 120, 220],
      ],
      [[400, 420, 120, 120, 220]],
      Array.from({ length: 5 }, () => [130, 170, 170, 180, 180]),
      [[190,110,110,110,110],[130,150,150,150,150],[130,160,160,160,160],
        [130,150,150,150,150],[130,160,160,160,160]],
      Array.from({ length: 5 }, () => [170,180,170,180,170]),
      Array.from({ length: 5 }, () => [170,180,170,180,170]),
    ];
    expected.forEach((rows, table) =>
      rows.forEach((cells, row) =>
        cells.forEach((after, cell) => {
          expect(
            projected.paragraphs(`body:t${table}:r${row}c${cell}`)[0].properties.spaceAfter
          ).toBe(after);
        })
      )
    );
    expect(projected.paragraphs('body:t0:r0c0')[0].properties.lineSpacing).toBe(360);
    expect(projected.paragraphs('body:t4:r0c0')[0].properties.lineSpacing).toBe(480);
    expect(parsed).toEqual(original);
  } finally {
    projected.destroy();
    native.destroy();
  }
});

it('applies conditional cell shading and borders identically for both seeders', async () => {
  const border = (side: string, size: number, color: string) =>
    `<w:${side} w:val="single" w:sz="${size}" w:space="0" w:color="${color}"/>`;
  const shading = (fill: string, theme = '') =>
    `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"${theme}/>`;
  const styles =
    '<w:style w:type="table" w:styleId="Shaded"><w:tblPr><w:tblStyleRowBandSize w:val="1"/>' +
    `<w:tblBorders>${border('insideH', 4, '000000')}</w:tblBorders></w:tblPr>` +
    `<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>${border('bottom', 12, 'FF0000')}</w:tcBorders>` +
    `${shading('4F81BD', ' w:themeFill="accent1"')}</w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="firstCol"><w:tcPr>${shading('00FF00')}</w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band1Horz"><w:tcPr>${shading('D3DFEE')}</w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="band2Horz"><w:tcPr><w:tcBorders>${border('top', 24, '000000')}</w:tcBorders></w:tcPr></w:tblStylePr></w:style>`;
  const theme =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom"><a:themeElements>' +
    '<a:clrScheme name="Custom"><a:accent1><a:srgbClr val="112233"/></a:accent1></a:clrScheme></a:themeElements></a:theme>';
  const rows = Array.from({ length: 4 }, (_, row) =>
    `<w:tr>${Array.from({ length: 3 }, (_, column) =>
      `<w:tc><w:tcPr>${row === 2 && column === 1 ? shading('ABCDEF') : ''}</w:tcPr>${paragraph(`r${row}c${column}`)}</w:tc>`
    ).join('')}</w:tr>`
  ).join('');
  const conditionalTable = (look: string) =>
    `<w:tbl><w:tblPr><w:tblStyle w:val="Shaded"/>${look}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(3)}</w:tblGrid>${rows}</w:tbl>`;
  const bytes = documentFixture(
    conditionalTable('<w:tblLook w:val="04A0"/>') +
      paragraph('Between') +
      conditionalTable(
        '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
      ),
    styles,
    theme
  );
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const projected = await createYrsSession({ clientId: 74220 });
  const native = await createYrsSession({ clientId: 74220 });
  try {
    documentToYrs(projected, parsed);
    native.seedFromDocx(bytes);
    for (const story of native.storyIds()) {
      expect(projected.yrsBlocksForStory(story)).toEqual(native.yrsBlocksForStory(story));
    }
    const cells = (table: number) =>
      (native.tablePayload('body', table)!.rows as Array<{ cells: Array<{ tcPr: Record<string, any> }> }>).map(
        (row) => row.cells.map((cell) => cell.tcPr)
      );
    const fills = (table: number) =>
      cells(table).map((row) => row.map((tcPr) => tcPr.backgroundColor ?? null));
    expect(fills(0)).toEqual([
      ['112233', '112233', '112233'],
      ['00FF00', 'D3DFEE', 'D3DFEE'],
      ['00FF00', 'ABCDEF', null],
      ['00FF00', 'D3DFEE', 'D3DFEE'],
    ]);
    expect(fills(1)).toEqual([
      ['112233', '112233', '112233'],
      [null, null, null],
      [null, 'ABCDEF', null],
      [null, null, null],
    ]);
    const [first, second] = [cells(0), cells(1)];
    expect(first[0][0].borders.bottom).toMatchObject({ size: 12, color: { rgb: 'FF0000' } });
    expect(first[1][1].borders.top).toMatchObject({ size: 4 });
    expect(first[2][2].borders.top).toMatchObject({ size: 24 });
    expect(second[2][2].borders.top).toMatchObject({ size: 4 });
    expect(first[2][1]._originalFormatting.shading.fill.rgb).toBe('ABCDEF');
    expect(first[1][1]._originalFormatting).toBeUndefined();
  } finally {
    projected.destroy();
    native.destroy();
  }
});

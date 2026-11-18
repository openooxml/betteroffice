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
      Array.from({ length: 5 }, () => [170, 180, 170, 180, 170]),
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

type SeededCell = { backgroundColor: string | null; borders: Record<string, any> | null; rowspan: number };

async function seedBoth(
  body: string,
  styles: string,
  edit?: (parsed: Awaited<ReturnType<typeof parseDocx>>) => void
): Promise<Record<string, SeededCell[][][]>> {
  const bytes = documentFixture(body, styles);
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  edit?.(parsed);
  const projected = await createYrsSession({ clientId: 74221 });
  const native = await createYrsSession({ clientId: 74221 });
  try {
    documentToYrs(projected, parsed);
    native.seedFromDocx(bytes);
    const seeded = (session: YrsSession) =>
      Array.from({ length: (body.match(/<w:tbl>/g) ?? []).length }, (_, table) =>
        (session.tablePayload('body', table)!.rows as Array<{ cells: Array<{ tcPr: Record<string, any> }> }>).map(
          (row) =>
            row.cells.map(({ tcPr }) => ({
              backgroundColor: tcPr.backgroundColor ?? null,
              borders: tcPr.borders ?? null,
              rowspan: tcPr.rowspan,
            }))
        )
      );
    return { projected: seeded(projected), native: seeded(native) };
  } finally {
    projected.destroy();
    native.destroy();
  }
}

const styledTable = (style: string, look: string, rows: string[][], columns: number) =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/>${look}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(columns)}</w:tblGrid>${rows
    .map((cells) => `<w:tr>${cells.map((tcPr, column) => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${paragraph(`c${column}`)}</w:tc>`).join('')}</w:tr>`)
    .join('')}</w:tbl>`;

it('lets a direct shd element replace the style shading in both seeders', async () => {
  const styles =
    '<w:style w:type="table" w:styleId="Header"><w:tblStylePr w:type="firstRow"><w:tcPr>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="0000FF"/></w:tcPr></w:tblStylePr></w:style>';
  const body = styledTable(
    'Header',
    '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
    [['<w:shd w:val="nil"/>', '<w:shd w:val="solid" w:color="FF0000" w:fill="auto"/>', '']],
    3
  );
  const seeded = await seedBoth(body, styles);
  for (const engine of ['projected', 'native']) {
    expect(seeded[engine]![0]![0]!.map((cell) => cell.backgroundColor)).toEqual([null, 'FF0000', '0000FF']);
  }
});

it('gives solid shading its pattern colour, black when automatic, in both seeders', async () => {
  const styles =
    '<w:style w:type="table" w:styleId="Solid">' +
    '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="solid" w:color="auto" w:fill="0000FF"/></w:tcPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="lastRow"><w:tcPr><w:shd w:val="solid" w:fill="0000FF"/></w:tcPr></w:tblStylePr></w:style>';
  const body = styledTable(
    'Solid',
    '<w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
    [
      ['', '', '', ''],
      [
        '<w:shd w:val="solid" w:color="auto" w:fill="FFFF00"/>',
        '<w:shd w:val="solid" w:fill="FFFF00"/>',
        '<w:shd w:val="solid" w:color="FF0000" w:fill="FFFF00"/>',
        '<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>',
      ],
      ['', '', '', ''],
    ],
    4
  );
  const seeded = await seedBoth(body, styles);
  for (const engine of ['projected', 'native']) {
    expect(seeded[engine]![0]!.map((row) => row.map((cell) => cell.backgroundColor))).toEqual([
      ['000000', '000000', '000000', '000000'],
      ['000000', '000000', 'FF0000', '00FF00'],
      ['000000', '000000', '000000', '000000'],
    ]);
  }
});

it('gives a vertically merged cell its region bottom border in both seeders', async () => {
  const styles =
    '<w:style w:type="table" w:styleId="Edged"><w:tblStylePr w:type="firstCol"><w:tcPr><w:tcBorders>' +
    '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="0000FF"/><w:insideH w:val="nil"/>' +
    '</w:tcBorders></w:tcPr></w:tblStylePr></w:style>';
  const body = styledTable(
    'Edged',
    '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
    [['', ''], ['<w:vMerge w:val="restart"/>', ''], ['<w:vMerge/>', '']],
    2
  );
  const seeded = await seedBoth(body, styles);
  for (const engine of ['projected', 'native']) {
    const anchor = seeded[engine]![0]![1]![0]!;
    expect(anchor.rowspan).toBe(2);
    expect(anchor.borders?.bottom).toMatchObject({ size: 8, color: { rgb: '0000FF' } });
  }
});

it('seeds identical conditional cell shading and borders in both engines', async () => {
  const line = (side: string, color: string) =>
    `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="${color}"/>`;
  const part = (type: string, fill: string, borders = '') =>
    `<w:tblStylePr w:type="${type}"><w:tcPr>${borders ? `<w:tcBorders>${borders}</w:tcBorders>` : ''}` +
    `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr></w:tblStylePr>`;
  const styles =
    '<w:style w:type="table" w:styleId="Full"><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>' +
    `<w:tblBorders>${line('top', '111111')}${line('bottom', '111111')}${line('insideH', '222222')}${line('insideV', '333333')}</w:tblBorders></w:tblPr>` +
    part('firstRow', 'F00000', line('bottom', 'F00000') + line('insideV', 'F0F0F0')) +
    part('lastRow', '0F0000', line('top', '0F0000')) +
    part('firstCol', '00F000', line('right', '00F000') + line('insideH', '00F0F0')) +
    part('lastCol', '000F00', line('left', '000F00')) +
    part('band1Horz', '0000F0', line('top', '0000F0') + line('bottom', '0000F0')) +
    part('band2Horz', '00000F') +
    part('band1Vert', 'A00000', line('left', 'A00000')) +
    part('band2Vert', '0A0000') +
    part('nwCell', '00A000', line('top', '00A000')) +
    part('neCell', '000A00', line('right', '000A00')) +
    part('swCell', '0000A0', line('left', '0000A0')) +
    part('seCell', '00000A', line('bottom', '00000A')) +
    '</w:style>';
  const grid = (rows: number, columns: number) =>
    Array.from({ length: rows }, (_, row) =>
      Array.from({ length: columns }, (_, column) =>
        row === 2 && column === 2 ? '<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"/>' : ''
      )
    );
  const body = [
    '<w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>',
    '<w:tblLook w:val="04A0"/>',
    '<w:tblLook w:firstRow="0" w:lastRow="1" w:firstColumn="0" w:lastColumn="1" w:noHBand="1" w:noVBand="0"/>',
  ]
    .map((look) => styledTable('Full', look, grid(6, 6), 6))
    .join(paragraph('Between'));
  const seeded = await seedBoth(body, styles);
  expect(seeded.projected).toEqual(seeded.native!);
  const [all] = seeded.native!;
  expect(all![0]![0]!.backgroundColor).toBe('00A000');
  expect(all![0]![5]!.backgroundColor).toBe('000A00');
  expect(all![5]![0]!.backgroundColor).toBe('0000A0');
  expect(all![5]![5]!.backgroundColor).toBe('00000A');
  expect(all![2]![2]!.backgroundColor).toBe('ABCDEF');
});

const shadingPart = (type: string, fill: string) =>
  `<w:tblStylePr w:type="${type}"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr></w:tblStylePr>`;
const edgeStyles = `<w:style w:type="table" w:styleId="Edges">${shadingPart('firstRow', 'FF0000')}${shadingPart('lastRow', '0000FF')}</w:style>`;
const bandedStyles = (bandSize: string) =>
  `<w:style w:type="table" w:styleId="Banded">${bandSize}${shadingPart('band1Horz', 'D3DFEE')}</w:style>`;
const flagsOff =
  '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
const firstColumnFills = (seeded: Record<string, SeededCell[][][]>, engine: string) =>
  seeded[engine]![0]!.map((row) => row[0]!.backgroundColor);

it('ignores the tblLook hex mask once a named flag is present in both seeders', async () => {
  const body = styledTable('Edges', '<w:tblLook w:val="04A0" w:lastRow="1"/>', [[''], [''], ['']], 1);
  const seeded = await seedBoth(body, edgeStyles);
  for (const engine of ['projected', 'native']) {
    expect(firstColumnFills(seeded, engine)).toEqual([null, null, '0000FF']);
  }
});

it('reads the tblLook hex mask when no named flag is present in both seeders', async () => {
  const body = styledTable('Edges', '<w:tblLook w:val="04A0"/>', [[''], [''], ['']], 1);
  const seeded = await seedBoth(body, edgeStyles);
  for (const engine of ['projected', 'native']) {
    expect(firstColumnFills(seeded, engine)).toEqual(['FF0000', null, null]);
  }
});

it('applies no horizontal banding without a row band size in both seeders', async () => {
  const seeded = await seedBoth(styledTable('Banded', flagsOff, [[''], [''], ['']], 1), bandedStyles(''));
  for (const engine of ['projected', 'native']) {
    expect(firstColumnFills(seeded, engine)).toEqual([null, null, null]);
  }
});

it('bands alternate rows with an explicit row band size of one in both seeders', async () => {
  const seeded = await seedBoth(
    styledTable('Banded', flagsOff, [[''], [''], ['']], 1),
    bandedStyles('<w:tblPr><w:tblStyleRowBandSize w:val="1"/></w:tblPr>')
  );
  for (const engine of ['projected', 'native']) {
    expect(firstColumnFills(seeded, engine)).toEqual(['D3DFEE', null, 'D3DFEE']);
  }
});

it('keeps wholeTable part shading off the cells in both seeders', async () => {
  const styles = `<w:style w:type="table" w:styleId="Whole">${shadingPart('wholeTable', '00FF00')}</w:style>`;
  const seeded = await seedBoth(styledTable('Whole', flagsOff, [['', ''], ['', '']], 2), styles);
  for (const engine of ['projected', 'native']) {
    expect(seeded[engine]![0]!.flat().map((cell) => cell.backgroundColor)).toEqual([null, null, null, null]);
  }
});

it('applies no vertical banding without a column band size in both seeders', async () => {
  const styles = `<w:style w:type="table" w:styleId="Columns">${shadingPart('band1Vert', 'D3DFEE')}</w:style>`;
  const look =
    '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="0"/>';
  const seeded = await seedBoth(styledTable('Columns', look, [['', '', '']], 3), styles);
  for (const engine of ['projected', 'native']) {
    expect(seeded[engine]![0]![0]!.map((cell) => cell.backgroundColor)).toEqual([null, null, null]);
  }
});

it('keeps wholeTable part paragraph spacing off the cells in both seeders', async () => {
  const styles =
    '<w:style w:type="table" w:styleId="Spaced"><w:tblStylePr w:type="wholeTable">' +
    '<w:pPr><w:spacing w:after="777"/></w:pPr></w:tblStylePr></w:style>';
  const bytes = documentFixture(styledTable('Spaced', flagsOff, [['', '']], 2), styles);
  const parsed = await parseDocx(bytes.buffer, { preloadFonts: false });
  const projected = await createYrsSession({ clientId: 74222 });
  const native = await createYrsSession({ clientId: 74222 });
  try {
    documentToYrs(projected, parsed);
    native.seedFromDocx(bytes);
    for (const session of [projected, native]) {
      for (const story of ['body:t0:r0c0', 'body:t0:r0c1']) {
        expect(session.paragraphs(story)[0]!.properties.spaceAfter).not.toBe(777);
      }
    }
  } finally {
    projected.destroy();
    native.destroy();
  }
});

it('keeps style-level tcBorders off the cells in both seeders', async () => {
  const line = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="0000FF"/>`;
  const styles =
    '<w:style w:type="table" w:styleId="Boxed"><w:tcPr><w:tcBorders>' +
    `${line('top')}${line('left')}${line('bottom')}${line('right')}</w:tcBorders></w:tcPr></w:style>`;
  const seeded = await seedBoth(styledTable('Boxed', flagsOff, [['', ''], ['', '']], 2), styles);
  for (const engine of ['projected', 'native']) {
    expect(
      seeded[engine]![0]!.flat().map((cell) => [cell.borders?.top ?? null, cell.borders?.left ?? null])
    ).toEqual(Array.from({ length: 4 }, () => [null, null]));
  }
});

it('keeps the table outer border where a band clears its own in both seeders', async () => {
  const line = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`;
  const nil = (side: string) => `<w:${side} w:val="nil"/>`;
  const styles =
    '<w:style w:type="table" w:styleId="Ruled"><w:tblPr><w:tblStyleRowBandSize w:val="1"/>' +
    `<w:tblBorders>${line('top')}${line('left')}${line('bottom')}${line('right')}</w:tblBorders></w:tblPr>` +
    `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:tcBorders>${nil('top')}${nil('bottom')}</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="firstCol"><w:tcPr><w:tcBorders>${nil('bottom')}</w:tcBorders></w:tcPr></w:tblStylePr></w:style>`;
  const look =
    '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
  const grid = Array.from({ length: 3 }, () => ['', '', '']);
  const seeded = await seedBoth(styledTable('Ruled', look, grid, 3), styles);
  for (const engine of ['projected', 'native']) {
    const cells = seeded[engine]![0]!;
    const size = (row: number, column: number, side: string) =>
      cells[row]![column]!.borders?.[side]?.size ?? null;
    expect([size(2, 0, 'bottom'), size(2, 1, 'bottom'), size(2, 2, 'bottom')]).toEqual([null, 8, 8]);
    expect([size(0, 1, 'top'), size(0, 2, 'top')]).toEqual([8, 8]);
  }
});

it('draws the header rule of a three-line table in both seeders', async () => {
  const line = (side: string, size: number) =>
    `<w:${side} w:val="single" w:sz="${size}" w:space="0" w:color="000000"/>`;
  const styles =
    '<w:style w:type="table" w:styleId="ThreeLine"><w:tblPr>' +
    `<w:tblBorders>${line('top', 12)}${line('bottom', 12)}</w:tblBorders></w:tblPr>` +
    `<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders>${line('bottom', 4)}</w:tcBorders></w:tcPr></w:tblStylePr></w:style>`;
  const look =
    '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>';
  const seeded = await seedBoth(styledTable('ThreeLine', look, [['', ''], ['', '']], 2), styles);
  for (const engine of ['projected', 'native']) {
    const cells = seeded[engine]![0]!;
    const size = (row: number, column: number, side: string) =>
      cells[row]![column]!.borders?.[side]?.size ?? null;
    expect([size(0, 0, 'top'), size(0, 1, 'top')]).toEqual([12, 12]);
    expect([size(0, 0, 'bottom'), size(0, 1, 'bottom')]).toEqual([4, 4]);
    expect([size(1, 0, 'top'), size(1, 1, 'top')]).toEqual([null, null]);
    expect([size(1, 0, 'bottom'), size(1, 1, 'bottom')]).toEqual([12, 12]);
  }
});

const single = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`;
const cleared = (side: string, value = 'nil') => `<w:${side} w:val="${value}"/>`;
const bandStyles = (tableBorders: string, regions: string[], sides: string) =>
  '<w:style w:type="table" w:styleId="Bands"><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>' +
  `<w:tblBorders>${tableBorders}</w:tblBorders></w:tblPr>` +
  regions
    .map((region) => `<w:tblStylePr w:type="${region}"><w:tcPr><w:tcBorders>${sides}</w:tcBorders></w:tcPr></w:tblStylePr>`)
    .join('') +
  '</w:style>';
const edgeSize = (cells: SeededCell[][], row: number, column: number, side: string) =>
  cells[row]![column]!.borders?.[side]?.size ?? null;

it('lets a band clear the bottom of a merged cell above the last row in both seeders', async () => {
  for (const value of ['nil', 'none']) {
    const styles = bandStyles(
      cleared('bottom') + single('insideH'),
      ['band1Horz', 'band2Horz'],
      cleared('bottom', value)
    );
    const body = styledTable('Bands', flagsOff, [['<w:vMerge w:val="restart"/>', ''], ['<w:vMerge/>', '']], 2);
    const seeded = await seedBoth(body, styles);
    for (const engine of ['projected', 'native']) {
      const cells = seeded[engine]![0]!;
      expect(cells[0]![0]!.rowspan).toBe(2);
      expect(edgeSize(cells, 0, 0, 'bottom')).toBeNull();
    }
  }
});

it('lets a band clear the right of a cell short of the grid in both seeders', async () => {
  for (const value of ['nil', 'none']) {
    const styles = bandStyles(single('right') + single('insideV'), ['band1Horz', 'band2Horz'], cleared('right', value));
    const body =
      `<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/>${flagsOff}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(3)}</w:tblGrid>` +
      `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc>${paragraph('c0')}</w:tc><w:tc>${paragraph('c1')}</w:tc></w:tr></w:tbl>`;
    const seeded = await seedBoth(body, styles);
    for (const engine of ['projected', 'native']) {
      expect(edgeSize(seeded[engine]![0]!, 0, 1, 'right')).toBeNull();
    }
  }
});

it('lets a band clear the sides of a right-to-left table in both seeders', async () => {
  const styles = bandStyles(
    single('left') + single('right') + cleared('insideV'),
    ['band1Horz', 'band2Horz'],
    cleared('left') + cleared('right')
  );
  const seeded = await seedBoth(styledTable('Bands', `<w:bidiVisual/>${flagsOff}`, [['', '']], 2), styles);
  for (const engine of ['projected', 'native']) {
    const cells = seeded[engine]![0]!;
    expect(edgeSize(cells, 0, 1, 'right')).toBeNull();
    expect(edgeSize(cells, 0, 0, 'left')).toBeNull();
  }
});

it('keeps the table sides where a vertical band clears its own in both seeders', async () => {
  const styles = bandStyles(
    single('left') + single('right') + single('insideV'),
    ['band1Vert', 'band2Vert'],
    cleared('left') + cleared('right')
  );
  const look =
    '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="0"/>';
  const body = styledTable('Bands', look, [['', '', '']], 3);
  const seeded = await seedBoth(body, styles);
  const unbanded = await seedBoth(withoutBands(body), styles);
  for (const engine of ['projected', 'native']) {
    const cells = seeded[engine]![0]!;
    const plain = unbanded[engine]![0]!;
    expect([edgeSize(cells, 0, 0, 'left'), edgeSize(cells, 0, 2, 'right')]).toEqual([8, 8]);
    expect(edgeSize(cells, 0, 1, 'left')).toBeNull();
    expect(edgeSize(plain, 0, 1, 'left')).toBe(8);
    expect(cells[0]![1]!.borders?.left ?? null).not.toEqual(plain[0]![1]!.borders?.left ?? null);
  }
});

const withoutBands = (body: string) =>
  body.replace(/w:noHBand="\d"/g, 'w:noHBand="1"').replace(/w:noVBand="\d"/g, 'w:noVBand="1"');
const sized = (side: string, size: number) =>
  `<w:${side} w:val="single" w:sz="${size}" w:space="0" w:color="000000"/>`;
const bothBands =
  '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>';
const layeredBandStyles = (tableBorders: string, parts: Array<[string, string]>) =>
  '<w:style w:type="table" w:styleId="Bands"><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>' +
  `<w:tblBorders>${tableBorders}</w:tblBorders></w:tblPr>` +
  parts
    .map(([region, sides]) => `<w:tblStylePr w:type="${region}"><w:tcPr><w:tcBorders>${sides}</w:tcBorders></w:tcPr></w:tblStylePr>`)
    .join('') +
  '</w:style>';
const offsetBandTable =
  `<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/>${bothBands}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(3)}</w:tblGrid>` +
  `<w:tr><w:tc>${paragraph('a')}</w:tc><w:tc>${paragraph('b')}</w:tc><w:tc>${paragraph('c')}</w:tc></w:tr>` +
  `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc>${paragraph('d')}</w:tc><w:tc>${paragraph('e')}</w:tc></w:tr></w:tbl>`;
const layeredBands = (value = 'nil') =>
  layeredBandStyles(sized('left', 4), [
    ['band2Horz', sized('insideV', 8)],
    ['band2Vert', cleared('left', value)],
  ]);

it('leaves no earlier band layer behind on a skipped band edge in both seeders', async () => {
  for (const value of ['nil', 'none']) {
    const seeded = await seedBoth(offsetBandTable, layeredBands(value));
    for (const engine of ['projected', 'native']) {
      expect(edgeSize(seeded[engine]![0]!, 1, 0, 'left')).toBe(4);
    }
  }
});

it('shows the table side where a vertical band clears a horizontal band edge in both seeders', async () => {
  const styles = layeredBandStyles(sized('left', 4), [
    ['band1Horz', sized('left', 12)],
    ['band1Vert', cleared('left')],
  ]);
  const seeded = await seedBoth(styledTable('Bands', bothBands, [['', '']], 2), styles);
  for (const engine of ['projected', 'native']) {
    expect(edgeSize(seeded[engine]![0]!, 0, 0, 'left')).toBe(4);
  }
});

it('takes the band edges of an explicit empty grid from the widest row in both seeders', async () => {
  const styles = bandStyles(sized('right', 12) + sized('insideV', 4), ['band1Vert', 'band2Vert'], cleared('right'));
  const look =
    '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="0"/>';
  const row = `<w:tr><w:tc>${paragraph('c0')}</w:tc><w:tc>${paragraph('c1')}</w:tc></w:tr>`;
  for (const grid of ['<w:tblGrid/>', '']) {
    const body = `<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/>${look}</w:tblPr>${grid}${row}</w:tbl>`;
    const seeded = await seedBoth(body, styles, (parsed) => {
      for (const block of parsed.package.document.content) {
        if (block.type !== 'table') continue;
        if (grid) block.columnWidths = [];
        else delete block.columnWidths;
      }
    });
    for (const engine of ['projected', 'native']) {
      const cells = seeded[engine]![0]!;
      expect(edgeSize(cells, 0, 0, 'right')).toBeNull();
      expect(edgeSize(cells, 0, 0, 'right')).not.toBe(12);
      expect(edgeSize(cells, 0, 1, 'right')).toBe(12);
    }
  }
});

it('gives a skipped band edge the border of the same table without bands in both seeders', async () => {
  const ruled =
    '<w:style w:type="table" w:styleId="Ruled"><w:tblPr><w:tblStyleRowBandSize w:val="1"/>' +
    `<w:tblBorders>${single('top')}${single('left')}${single('bottom')}${single('right')}</w:tblBorders></w:tblPr>` +
    `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:tcBorders>${cleared('top')}${cleared('bottom')}</w:tcBorders></w:tcPr></w:tblStylePr>` +
    `<w:tblStylePr w:type="firstCol"><w:tcPr><w:tcBorders>${cleared('bottom')}</w:tcBorders></w:tcPr></w:tblStylePr></w:style>`;
  const ruledLook =
    '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
  const offsetTable =
    `<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/>${flagsOff}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(3)}</w:tblGrid>` +
    `<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc>${paragraph('c0')}</w:tc><w:tc>${paragraph('c1')}</w:tc></w:tr></w:tbl>`;
  const constructions: Array<[string, string, Array<[number, number, string]>]> = [
    [
      styledTable('Ruled', ruledLook, Array.from({ length: 3 }, () => ['', '', '']), 3),
      ruled,
      [
        [0, 1, 'top'],
        [0, 2, 'top'],
        [2, 1, 'bottom'],
        [2, 2, 'bottom'],
      ],
    ],
    [
      styledTable('Bands', flagsOff, [['<w:vMerge w:val="restart"/>', ''], ['<w:vMerge/>', '']], 2),
      bandStyles(cleared('bottom') + single('insideH'), ['band1Horz', 'band2Horz'], cleared('bottom')),
      [[1, 0, 'bottom']],
    ],
    [
      offsetTable,
      bandStyles(single('left') + single('insideV'), ['band1Horz', 'band2Horz'], cleared('insideV')),
      [[0, 0, 'left']],
    ],
    [offsetBandTable, layeredBands(), [[1, 0, 'left']]],
  ];
  for (const [body, styles, edges] of constructions) {
    const banded = await seedBoth(body, styles);
    const unbanded = await seedBoth(withoutBands(body), styles);
    for (const engine of ['projected', 'native']) {
      for (const [row, column, side] of edges) {
        const edge = (seeded: Record<string, SeededCell[][][]>) =>
          seeded[engine]![0]![row]![column]!.borders?.[side] ?? null;
        expect(edge(banded)).toEqual(edge(unbanded));
      }
    }
  }
});

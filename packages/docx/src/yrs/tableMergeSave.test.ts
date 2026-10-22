import { beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx, type YrsSession, type YrsTableRange } from './index';
import { documentToYrs } from './documentToYrs';
import { captureSessionSave, writeSessionSave } from './saveYrsDocx';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

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

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const table = (...cells: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cells
    .map(() => '<w:gridCol w:w="4680"/>')
    .join('')}</w:tblGrid><w:tr>${cells.map(cell).join('')}</w:tr></w:tbl>`;

function docx(body: string, namespaces = ''): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set(
    'word/document.xml',
    toBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"${namespaces}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const FIRST_ROW: YrsTableRange = {
  anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 },
  head: { story: 'body', tableIndex: 0, row: 0, column: 1 },
};

async function opened(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession();
  session.seedFromDocx(bytes);
  return session;
}

async function saveLegacyTable(session: YrsSession, fallback = true) {
  const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
  try {
    expect(session.paragraphSavePlan().splicedParts?.flatMap(
      ({ tablePatches }) => tablePatches ?? []
    ) ?? []).toHaveLength(0);
    let baseline: (() => Promise<Uint8Array>) | undefined;
    if (fallback) {
      const base = session.materializeDocx()!;
      expect(session.paragraphSavePlan().splicedParts?.find(
        ({ part }) => part === 'word/document.xml'
      )).toBeUndefined();
      const capture = captureSessionSave(session);
      const document = yrsToDocument(session, base);
      baseline = async () => (await writeSessionSave(
        session, document, capture, base.originalBuffer!, { updateModifiedDate: false },
        () => false, true
      )).bytes;
    }
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    if (baseline) {
      const expected = await baseline();
      expect(unzipContainer(saved.bytes)['word/document.xml']).toEqual(
        unzipContainer(expected)['word/document.xml']
      );
    }
    expect(diagnostic.mock.calls.filter(([warning]) =>
      typeof warning === 'string' && warning.startsWith('DOCX_MINIMAL_PATCH replace counts')
    )).toEqual([]);
    return saved;
  } finally {
    diagnostic.mockRestore();
  }
}

async function reopenedStoryTexts(bytes: Uint8Array): Promise<Record<string, string[]>> {
  const session = await opened(bytes);
  try {
    return Object.fromEntries(
      session
        .storyIds()
        .filter((story) => story !== 'body' && story.startsWith('body'))
        .map((story) => [story, session.paragraphs(story).map((p) => p.text)])
    );
  } finally {
    session.destroy();
  }
}

describe('saving merged table cells after an earlier save', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('keeps the merged-in cell text', async () => {
    const session = await opened(
      docx(paragraph('Before') + table(paragraph('CellA'), paragraph('CellB')) + paragraph('After'))
    );
    try {
      await saveYrsDocx(session);
      session.mergeCells(FIRST_ROW);
      const saved = await saveLegacyTable(session);
      expect(await reopenedStoryTexts(saved.bytes)).toEqual({ 'body:t0:r0c0': ['CellA', 'CellB'] });
    } finally {
      session.destroy();
    }
  });

  it('keeps edits to a nested table the merge moved into the surviving cell', async () => {
    const session = await opened(
      docx(table(paragraph('CellA'), table(paragraph('Inner')) + paragraph('')))
    );
    try {
      const nested = 'body:t0:r0c1:t0:r0c0';
      expect(session.paragraphs(nested).map((p) => p.text)).toEqual(['Inner']);
      await saveYrsDocx(session);
      session.mergeCells(FIRST_ROW);
      expect(session.paragraphSavePlan().splicedParts?.find(
        ({ part }) => part === 'word/document.xml'
      )?.tablePatches ?? []).toHaveLength(0);
      await saveLegacyTable(session);
      const [inner] = session.paragraphs(nested);
      session.insertText({ story: nested, paraId: inner!.paraId, offset: 5 }, ' edited');
      const saved = await saveYrsDocx(session);
      expect(await reopenedStoryTexts(saved.bytes)).toEqual({
        'body:t0:r0c0': ['CellA', ''],
        'body:t0:r0c0:t0:r0c0': ['Inner edited'],
      });
    } finally {
      session.destroy();
    }
  });
});

describe('saving untouched vertical merges after reseeding', () => {
  beforeAll(async () => {
    await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
    await preloadOpcWasm(new Uint8Array(readFileSync(
      resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')
    )));
  });

  it('keeps grid offsets when merging the middle columns across rows', async () => {
    const offsets = '<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>';
    const left = cell(paragraph('Left'));
    const right = cell(paragraph('Right'));
    const grid = '<w:gridCol w:w="4680"/>'.repeat(4);
    const body = `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>` +
      `<w:tr>${left}${cell(paragraph('A'))}${cell(paragraph('B'))}${right}</w:tr>` +
      `<w:tr>${offsets}${cell(paragraph('C'))}${cell(paragraph('D'))}</w:tr></w:tbl>`;
    const unrelated = table(paragraph('Unrelated')).replace(
      '<w:tblPr>',
      '<w:tblPr><w:tblCaption w:val="caption"/><w:tblDescription w:val="description"/>'
    );
    const tail = paragraph('Tail');
    const session = await opened(docx(body + unrelated + tail));
    try {
      session.mergeCells({
        anchor: { story: 'body', tableIndex: 0, row: 0, column: 1 },
        head: { story: 'body', tableIndex: 0, row: 1, column: 2 },
      });
      const saved = await saveLegacyTable(session);
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      expect(xml).toContain('<w:tblCaption w:val="caption"/>');
      expect(xml).toContain('<w:tblDescription w:val="description"/>');
      const rows = xml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g)!;
      expect(rows[0]).toMatch(/<w:gridSpan w:val="2"\/><w:vMerge w:val="restart"\/>/);
      expect(rows[1]).toContain(offsets);
      expect(rows[1].match(/<w:tc>/g)).toHaveLength(1);
      expect(rows[1]).toMatch(/<w:gridSpan w:val="2"\/><w:vMerge(?: w:val="continue")?\/>/);
      expect(xml).toContain('Unrelated');
      expect(xml).toContain('Tail');
      expect(Object.values(await reopenedStoryTexts(saved.bytes)).flat().filter(Boolean)).toEqual([
        'Left', 'A', 'B', 'C', 'D', 'Right', 'Unrelated',
      ]);
    } finally {
      session.destroy();
    }
  });

  it('uses the base fallback for vertical cell merges', async () => {
    const owner = (text: string) => '<w:tc><w:tcPr><w:vMerge w:val="restart"/>' +
      `</w:tcPr>${paragraph(text)}</w:tc>`;
    const hidden = '<w:tc><w:tcPr><w:vMerge/><w:shd w:fill="778899"/></w:tcPr>' +
      '<w:p w14:paraId="00000006"><w:r><w:rPr><w:noProof/></w:rPr>' +
      '<w:t>Hidden one</w:t></w:r></w:p><w:p w14:paraId="00000007">' +
      '<w:r><w:rPr><w:noProof/></w:rPr><w:t>Hidden two</w:t></w:r></w:p></w:tc>';
    const empty = '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>';
    const grid = '<w:gridCol w:w="4680"/>'.repeat(3);
    const body = `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>` +
      `<w:tr>${owner('A')}${owner('B')}${owner('Keep')}</w:tr>` +
      `<w:tr>${empty}${empty}${hidden}</w:tr></w:tbl>`;
    const unrelated = table(paragraph('Unrelated'));
    const tail = paragraph('Tail');
    const session = await opened(docx(
      body + unrelated + tail,
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    ));
    try {
      session.mergeCells({
        anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 },
        head: { story: 'body', tableIndex: 0, row: 1, column: 1 },
      });
      const saved = await saveLegacyTable(session);
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      expect(xml).toContain('Unrelated');
      expect(xml).toContain('Tail');
      expect(Object.values(await reopenedStoryTexts(saved.bytes)).flat().filter(Boolean)).toEqual([
        'A', 'B', 'Keep', 'Hidden one', 'Hidden two', 'Unrelated',
      ]);
    } finally {
      session.destroy();
    }
  });

  it('uses the base row splice when appending beside continuations', async () => {
    const hidden = '<w:tc><w:tcPr><w:vMerge/></w:tcPr>' +
      '<w:p w14:paraId="00000002"><w:bookmarkStart w:id="7" w:name="Kept"/>' +
      '<w:r><w:t>Hidden one</w:t></w:r></w:p><w:p w14:paraId="00000003">' +
      '<w:r><w:t>Hidden two</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p></w:tc>';
    const ordinary = `<w:tr>${cell(paragraph('Ordinary'))}</w:tr>`;
    const prefix = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' +
      `${paragraph('Owner')}</w:tc></w:tr><w:tr>${hidden}</w:tr>${ordinary}`;
    const unrelated = table(paragraph('Unrelated'));
    const tail = paragraph('Tail');
    const session = await opened(docx(
      `${prefix}</w:tbl>${unrelated}${tail}`,
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    ));
    try {
      session.insertRow({ story: 'body', tableIndex: 0, row: 2, column: 0 }, 'below');
      const saved = await saveLegacyTable(session, false);
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      expect(xml).toContain(hidden);
      expect(xml).toContain('Unrelated');
      expect(xml).toContain('Tail');
      expect(xml).toContain('Ordinary');
    } finally {
      session.destroy();
    }
  });

  it('seeds the same row provenance through native and model loading', async () => {
    const bytes = docx(table(paragraph('Cell')) + paragraph('Tail'));
    const parsed = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
    const native = await createYrsSession({ clientId: 74290 });
    const projected = await createYrsSession({ clientId: 74290 });
    try {
      native.seedFromDocx(bytes);
      documentToYrs(projected, parsed);
      expect(projected.yrsBlocksForStory('body')).toEqual(native.yrsBlocksForStory('body'));
    } finally {
      native.destroy();
      projected.destroy();
    }
  });

  it('retains both continuation paragraphs and unrelated table bytes across reopen', async () => {
    const p = (id: number, text: string) =>
      `<w:p w14:paraId="${id.toString(16).padStart(8, '0').toUpperCase()}"><w:r><w:rPr><w:noProof/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    const merged = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${p(1, 'Owner')}</w:tc></w:tr>` +
      '<w:tr w:rsidR="11223344"><w:trPr><w:trHeight w:val="720"/></w:trPr>' +
      `<w:tc><w:tcPr><w:vMerge/><w:shd w:fill="ABCDEF"/></w:tcPr>${p(2, 'Hidden one')}${p(3, 'Hidden two')}</w:tc></w:tr></w:tbl>`;
    const unrelated = '<w:tbl><w:tblPr><w:tblCaption w:val="caption"/>' +
      '<w:tblDescription w:val="description"/></w:tblPr>' +
      `<w:tr><w:tc>${p(4, 'Unrelated')}</w:tc></w:tr></w:tbl>`;
    let bytes = docx(
      merged + unrelated + p(5, 'Tail'),
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    );
    for (let cycle = 0; cycle < 2; cycle += 1) {
      const session = await opened(bytes);
      try {
        const tail = session.paragraphs('body').find(({ text }) => text === 'Tail')!;
        expect(session.applyEdits({
          expectVersion: session.version(),
          steps: [{
            op: 'insertParagraphs', target: { story: 'body', paraId: tail.paraId },
            at: 'end', paragraphs: [{ text: `Inserted ${cycle}` }],
          }],
        }).ok).toBe(true);
        bytes = (await saveYrsDocx(session)).bytes;
        const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
        expect(xml).toContain(merged);
        expect(xml).toContain('Unrelated');
      } finally {
        session.destroy();
      }
    }
  });

  it.todo('retains continuation bookmarks through unrelated insertions and reopen (requires reverted continuation reseeding)', async () => {
    const merged = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4680"/>' +
      '<w:gridCol w:w="4680"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' +
      '<w:p w14:paraId="00000001"><w:r><w:t>Owner</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p w14:paraId="00000005"><w:r><w:t>Top side</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p w14:paraId="00000002">' +
      '<w:bookmarkStart w:id="7" w:name="Retained"/>' +
      '<w:r><w:t>Hidden one</w:t></w:r></w:p><w:p w14:paraId="00000003">' +
      '<w:r><w:t>Hidden two</w:t></w:r><w:bookmarkEnd w:id="7"/>' +
      '</w:p></w:tc><w:tc>' +
      '<w:p w14:paraId="00000006"><w:r><w:t>Bottom side</w:t></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl>';
    let bytes = docx(
      merged + '<w:p w14:paraId="00000004"><w:r><w:t>Tail</w:t></w:r></w:p>',
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
    );
    for (let cycle = 0; cycle < 2; cycle += 1) {
      const session = await opened(bytes);
      try {
        const tail = session.paragraphs('body').find(({ text }) => text === 'Tail')!;
        expect(session.applyEdits({
          expectVersion: session.version(),
          steps: [{
            op: 'insertParagraphs', target: { story: 'body', paraId: tail.paraId },
            at: 'end', paragraphs: [{ text: `Inserted ${cycle}` }],
          }],
        }).ok).toBe(true);
        bytes = (await saveYrsDocx(session)).bytes;
        const xml = new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
        expect(xml).toContain(merged);
      } finally {
        session.destroy();
      }
    }
  });

  it('seeds vertical merges after a grid offset on the native grid', async () => {
    const restart = (text: string) =>
      `<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${paragraph(text)}</w:tc>`;
    const grid = '<w:gridCol w:w="4680"/>'.repeat(4);
    const body = `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>` +
      '<w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>' +
      `${restart('A')}${restart('B')}</w:tr>` +
      `<w:tr>${cell(paragraph('X'))}<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>` +
      `<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${paragraph('D')}</w:tc></w:tr></w:tbl>`;
    const session = await opened(docx(body + table(paragraph('Unrelated')) + paragraph('Tail')));
    try {
      session.setCellShading({
        anchor: { story: 'body', tableIndex: 0, row: 1, column: 2 },
        head: { story: 'body', tableIndex: 0, row: 1, column: 3 },
      }, '112233');
      const saved = await saveLegacyTable(session);
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      const rows = xml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g)!;
      expect(rows[0]).toContain('<w:gridBefore w:val="1"/>');
      expect(rows[0].match(/<w:tc>/g)).toHaveLength(2);
      expect(rows[1].match(/<w:tc>/g)).toHaveLength(3);
      expect(rows[1]).toMatch(
        /X[\s\S]*<w:vMerge(?: w:val="continue")?\/>[\s\S]*<w:gridSpan w:val="2"\/>[\s\S]*112233[\s\S]*D/
      );
      const reopened = Object.entries(await reopenedStoryTexts(saved.bytes))
        .map(([story, texts]) => [story, texts.filter(Boolean)]);
      expect(Object.fromEntries(reopened)).toEqual({
        'body:t0:r0c0': ['A'],
        'body:t0:r0c1': ['B'],
        'body:t0:r1c0': ['X'],
        'body:t0:r1c1': ['D'],
        'body:t1:r0c0': ['Unrelated'],
      });
    } finally {
      session.destroy();
    }
  });
});

describe('saving deleted table cells', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('retains deleted cells through a save until a fresh seed replaces the session', async () => {
    const session = await opened(docx(table(paragraph('CellA')) + paragraph('After')));
    try {
      session.deleteTable({ story: 'body', tableIndex: 0 });
      const saved = await saveYrsDocx(session);
      expect(new TextDecoder().decode(session.encodeState())).toContain('CellA');
      expect(session.storyIds()).toEqual(['body']);
      const reseeded = await opened(saved.bytes);
      try {
        expect(new TextDecoder().decode(reseeded.encodeState())).not.toContain('CellA');
        expect(reseeded.storyIds()).toEqual(['body']);
      } finally {
        reseeded.destroy();
      }
    } finally {
      session.destroy();
    }
  });

  for (const removal of ['table', 'row', 'column'] as const) {
    it(`excludes deleted cells after deleting a ${removal}, undoing and redoing`, async () => {
      const secondRow = `<w:tr>${cell(paragraph('CellC'))}${cell(paragraph('CellD'))}</w:tr>`;
      const content = table(paragraph('CellA'), paragraph('CellB')).replace(
        '</w:tbl>',
        `${secondRow}</w:tbl>`
      );
      const session = await opened(docx(paragraph('Before') + content + paragraph('After')));
      try {
        await saveYrsDocx(session);
        const first = { story: 'body', tableIndex: 0, row: 0, column: 0 };
        if (removal === 'table') session.deleteTable(first);
        else if (removal === 'row') session.deleteRow({ anchor: first, head: first });
        else session.deleteColumn({ anchor: first, head: first });
        const expected = {
          table: [],
          row: [['CellC'], ['CellD']],
          column: [['CellB'], ['CellD']],
        }[removal];
        for (let cycle = 0; cycle < 2; cycle += 1) {
          const saved = await saveYrsDocx(session);
          expect(Object.values(await reopenedStoryTexts(saved.bytes))).toEqual(expected);
          expect(session.undo()).toBe(true);
          const restored = await saveYrsDocx(session);
          expect(Object.values(await reopenedStoryTexts(restored.bytes))).toEqual([
            ['CellA'],
            ['CellB'],
            ['CellC'],
            ['CellD'],
          ]);
          expect(session.redo()).toBe(true);
        }
      } finally {
        session.destroy();
      }
    });
  }
});

describe('saving table rows with grid offsets', () => {
  beforeAll(async () => {
    await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
    await preloadOpcWasm(new Uint8Array(readFileSync(
      resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')
    )));
  });

  it('keeps row grid offsets and their widths when another row is rewritten', async () => {
    const offsets = '<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>' +
      '<w:wBefore w:w="4680" w:type="dxa"/><w:wAfter w:w="4680" w:type="dxa"/></w:trPr>';
    const grid = '<w:gridCol w:w="4680"/>'.repeat(4);
    const body = `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>` +
      `<w:tr>${['Left', 'A', 'B', 'Right'].map((text) => cell(paragraph(text))).join('')}</w:tr>` +
      `<w:tr>${offsets}${cell(paragraph('C'))}${cell(paragraph('D'))}</w:tr></w:tbl>${paragraph('Tail')}`;
    const session = await opened(docx(body));
    try {
      session.mergeCells({
        anchor: { story: 'body', tableIndex: 0, row: 0, column: 1 },
        head: { story: 'body', tableIndex: 0, row: 0, column: 2 },
      });
      const saved = await saveYrsDocx(session, { updateModifiedDate: false });
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      const rows = xml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g)!;
      expect(rows).toHaveLength(2);
      expect(rows[0].match(/<w:tc>/g)).toHaveLength(3);
      expect(rows[0]).toContain('<w:gridSpan w:val="2"/>');
      expect(rows[1]).toContain(offsets);
      expect(rows[1].match(/<w:tc>/g)).toHaveLength(2);
      expect(rows[1]).toMatch(/>C<[\s\S]*>D</);
      expect(Object.values(await reopenedStoryTexts(saved.bytes)).flat().filter(Boolean).sort()).toEqual(
        ['A', 'B', 'C', 'D', 'Left', 'Right']
      );
    } finally {
      session.destroy();
    }
  });
});

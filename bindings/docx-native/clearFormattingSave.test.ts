import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { createEmptyDocx } from '../../packages/docx/src/docx/rezip/createEmpty';
import { applyClipboardFragment } from '../../packages/docx/src/yrs/clipboardFragment';
import { savedRunProperties } from '../../packages/docx/src/yrs/__fixtures__/textRuns';
import type { DocxEditStep } from '../../packages/docx/src/yrs/edits';

type Yrs = typeof wasmYrs;
type Session = Awaited<ReturnType<Yrs['createYrsSession']>>;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>';
const HEADING = '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>';

const FORMATTED_BODY =
  `<w:p w14:paraId="00000001">${HEADING}<w:r><w:t>Title</w:t></w:r></w:p>` +
  `<w:p w14:paraId="00000002">${HEADING}<w:r><w:rPr><w:b w:val="0"/><w:i/><w:color w:val="C00000"/></w:rPr><w:t>Unbolded</w:t></w:r></w:p>` +
  '<w:p w14:paraId="00000003"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold words</w:t></w:r></w:p>' +
  `<w:p w14:paraId="00000004">${HEADING}<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Directbold</w:t></w:r></w:p>`;

const FORMATTED_IDS = ['00000001', '00000002', '00000003', '00000004'];

const EDITED_BODY =
  `<w:p w14:paraId="00000011">${HEADING}<w:r><w:t>Plain heading</w:t></w:r></w:p>` +
  `<w:p w14:paraId="00000012">${HEADING}<w:r><w:rPr><w:b w:val="0"/><w:i/><w:color w:val="C00000"/><w:sz w:val="40"/></w:rPr><w:t>Styled heading</w:t></w:r></w:p>` +
  '<w:p w14:paraId="00000013"><w:r><w:t>Plain body</w:t></w:r></w:p>' +
  '<w:p w14:paraId="00000014"><w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>Styled body</w:t></w:r></w:p>' +
  '<w:p w14:paraId="00000015"><w:r><w:t>End</w:t></w:r></w:p>';

const UNFORMATTED_BODY =
  `<w:p w14:paraId="00000041">${HEADING}<w:r><w:t>Plain heading</w:t></w:r></w:p>` +
  '<w:p w14:paraId="00000042"><w:r><w:t>Plain one</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblStyle w:val="BoldGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
  '<w:p w14:paraId="00000043"><w:r><w:t>Cell words</w:t></w:r></w:p>' +
  '</w:tc></w:tr></w:tbl>' +
  '<w:p w14:paraId="00000044"><w:r><w:t>After</w:t></w:r></w:p>';

const BOLD_GRID = '<w:style w:type="table" w:styleId="BoldGrid"><w:name w:val="Bold Grid"/><w:rPr><w:b/></w:rPr></w:style>';

async function fixture(body = FORMATTED_BODY): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`, { date });
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', { date });
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', { date });
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}">${STYLES}${BOLD_GRID}</w:styles>`, { date });
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const whole = (session: Session, from: string, to = from, story = 'body') => {
  const length = session.paragraphs(story).find((entry) => entry.paraId === to)!.text.length;
  return { story, start: { paraId: from, offset: 0 }, end: { paraId: to, offset: length } };
};

const bold = (session: Session, marker: string) => {
  const segment = session.storySegments('body').find((entry) => entry.kind === 'text' && entry.text.includes(marker));
  return segment?.kind === 'text' ? Boolean(segment.attributes.bold) : undefined;
};

async function saveAndReopen(yrs: Yrs, session: Session) {
  const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
  const xml = await (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('string');
  const reopened = await yrs.createYrsSession({ clientId: 92 });
  reopened.seedFromDocx(saved.bytes);
  return { xml, reopened };
}

for (const [engine, yrs] of [['wasm', wasmYrs], ['native', nativeYrs as unknown as Yrs]] as const) {
  test(`${engine}: clear formatting keeps heading runs heading-bold and saves no run properties`, async () => {
    const clears: Record<string, (session: Session) => void> = {
      perParagraph: (session) => {
        for (const paraId of FORMATTED_IDS) session.clearFormatting(whole(session, paraId));
      },
      oneRange: (session) => session.clearFormatting(whole(session, FORMATTED_IDS[0]!, FORMATTED_IDS[3])),
      hostClearAll: (session) => {
        const steps = FORMATTED_IDS.map((paraId) => ({ op: 'formatText', target: { kind: 'paragraph', story: 'body', paraId }, clear: 'all' }) as DocxEditStep);
        const result = session.applyEdits({ expectVersion: session.version(), steps });
        expect(result.ok).toBe(true);
      },
    };
    for (const [name, clear] of Object.entries(clears)) {
      const session = await yrs.createYrsSession({ clientId: 91 });
      try {
        session.seedFromDocx(await fixture());
        clear(session);
        const markers = ['Title', 'Unbolded', 'Bold words', 'Directbold'];
        expect({ name, bold: markers.map((marker) => bold(session, marker)) }).toEqual({ name, bold: [true, true, false, true] });
        const { xml, reopened } = await saveAndReopen(yrs, session);
        try {
          expect(xml).not.toContain('<w:rPr></w:rPr>');
          expect(xml).not.toContain('<w:rPr/>');
          for (const marker of ['Title', 'Unbolded', 'Bold words']) {
            expect({ name, marker, saved: savedRunProperties(xml, marker) }).toEqual({ name, marker, saved: null });
          }
          expect({ name, bold: markers.map((marker) => bold(reopened, marker)) }).toEqual({ name, bold: [true, true, false, true] });
        } finally {
          reopened.destroy();
        }
      } finally {
        session.destroy();
      }
    }
  });

  test(`${engine}: in a cell of a bold table style the saved bold follows the table style`, async () => {
    const cells = ['Unbold cell', 'Grid bold', 'Italic cell'];
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="BoldGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
      '<w:p w14:paraId="00000021"><w:r><w:rPr><w:b w:val="0"/><w:i/></w:rPr><w:t>Unbold cell</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000022"><w:r><w:t>Grid bold</w:t></w:r></w:p>' +
      '<w:p w14:paraId="00000023"><w:r><w:rPr><w:b w:val="0"/><w:i/></w:rPr><w:t>Italic cell</w:t></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl><w:p w14:paraId="00000024"><w:r><w:t>After</w:t></w:r></w:p>';
    const session = await yrs.createYrsSession({ clientId: 98 });
    try {
      session.seedFromDocx(await fixture(body));
      const cell = session.listStories().stories.find((entry) => entry.kind === 'tableCell')!.story;
      session.clearFormatting(whole(session, '00000021', '00000021', cell));
      session.formatRange(whole(session, '00000022', '00000022', cell), { bold: false });
      session.formatRange(whole(session, '00000023', '00000023', cell), { bold: true });
      const { xml, reopened } = await saveAndReopen(yrs, session);
      try {
        expect(cells.map((marker) => savedRunProperties(xml, marker))).toEqual([null, '<w:b w:val="0"/>', '<w:i/>']);
        const reopenedCell = reopened.listStories().stories.find((entry) => entry.kind === 'tableCell')!.story;
        const boldIn = (marker: string) => {
          const segment = reopened.storySegments(reopenedCell).find((entry) => entry.kind === 'text' && entry.text.includes(marker));
          return segment?.kind === 'text' ? Boolean(segment.attributes.bold) : undefined;
        };
        expect(cells.map(boldIn)).toEqual([true, false, true]);
      } finally {
        reopened.destroy();
      }
    } finally {
      session.destroy();
    }
  });

  test(`${engine}: known limit: a direct value equal to the style stays after clear formatting`, async () => {
    const session = await yrs.createYrsSession({ clientId: 96 });
    try {
      session.seedFromDocx(await fixture());
      session.clearFormatting(whole(session, '00000004'));
      const { xml, reopened } = await saveAndReopen(yrs, session);
      reopened.destroy();
      expect(savedRunProperties(xml, 'Directbold')).toBe('<w:b/>');
    } finally {
      session.destroy();
    }
  });

  test(`${engine}: an edit saves the same run properties whether or not its range holds the paragraph mark`, async () => {
    const ids = ['00000011', '00000012', '00000013', '00000014', '00000015'];
    const edits: [string, string, (session: Session, range: ReturnType<typeof whole>) => void][] = [
      ['clear on a heading run', '00000012', (session, range) => session.clearFormatting(range)],
      ['clear on a body run', '00000014', (session, range) => session.clearFormatting(range)],
      ['bold on over a bold style', '00000012', (session, range) => session.formatRange(range, { bold: true })],
      ['bold off over a bold style', '00000011', (session, range) => session.formatRange(range, { bold: false })],
      ['bold on in a body paragraph', '00000013', (session, range) => session.formatRange(range, { bold: true })],
      ['bold off in a body paragraph', '00000014', (session, range) => session.formatRange(range, { bold: false })],
      ['italic on', '00000013', (session, range) => session.formatRange(range, { italic: true })],
      ['italic off', '00000014', (session, range) => session.formatRange(range, { italic: false })],
      ['italic off over a heading', '00000012', (session, range) => session.formatRange(range, { italic: false })],
      ['size set to the style size', '00000012', (session, range) => session.formatRange(range, { fontSize: 16 })],
      ['size set to another size', '00000011', (session, range) => session.formatRange(range, { fontSize: 11 })],
      ['underline off', '00000014', (session, range) => session.formatRange(range, { underline: null })],
    ];
    for (const [name, paraId, edit] of edits) {
      const saved: (string | null)[] = [];
      for (const withMark of [false, true]) {
        const session = await yrs.createYrsSession({ clientId: 97 });
        try {
          session.seedFromDocx(await fixture(EDITED_BODY));
          const range = whole(session, paraId);
          const next = ids[ids.indexOf(paraId) + 1]!;
          edit(session, withMark ? { ...range, end: { paraId: next, offset: 0 } } : range);
          const { xml, reopened } = await saveAndReopen(yrs, session);
          reopened.destroy();
          const marker = { '00000011': 'Plain heading', '00000012': 'Styled heading', '00000013': 'Plain body', '00000014': 'Styled body' }[paraId]!;
          saved.push(savedRunProperties(xml, marker));
        } finally {
          session.destroy();
        }
      }
      expect({ name, withMark: saved[1] }).toEqual({ name, withMark: saved[0] });
    }
  });

  test(`${engine}: text pasted into a heading keeps the heading bold and its own italic`, async () => {
    const session = await yrs.createYrsSession({ clientId: 93 });
    try {
      session.seedFromDocx(await fixture());
      const at = { paraId: '00000001', offset: 5 };
      applyClipboardFragment(session, { story: 'body', start: at, end: at }, {
        paragraphs: [{ text: 'Pasted', runs: [{ text: 'Pasted', format: { other: { italic: true } } }] }],
      });
      const segment = session.storySegments('body').find((entry) => entry.kind === 'text' && entry.text.includes('Pasted'));
      expect(segment?.kind === 'text' && segment.attributes.bold).toBe(true);
      expect(segment?.kind === 'text' && segment.attributes.italic).toBe(true);
    } finally {
      session.destroy();
    }
  });

  test(`${engine}: a new blank document clears direct formatting back to its defaults`, async () => {
    const session = await yrs.createYrsSession({ clientId: 94 });
    try {
      session.seedFromDocx(new Uint8Array(await createEmptyDocx()));
      const paraId = session.paragraphs('body')[0]!.paraId;
      session.insertText({ story: 'body', paraId, offset: 0 }, 'Fresh words');
      const range = whole(session, paraId);
      session.formatRange(range, { bold: true, fontSizePt: 20 });
      session.clearFormatting(range);
      const { xml, reopened } = await saveAndReopen(yrs, session);
      try {
        expect(savedRunProperties(xml, 'Fresh words')).toBeNull();
        expect(bold(reopened, 'Fresh words')).toBe(false);
      } finally {
        reopened.destroy();
      }
    } finally {
      session.destroy();
    }
  });

  test(`${engine}: clearing unformatted text writes no update and adds no undo step`, async () => {
    for (const suggesting of [undefined, { name: 'Ann', date: '2026-11-10T12:00:00Z' }]) {
      const session = await yrs.createYrsSession({ clientId: 95 });
      try {
        session.seedFromDocx(await fixture(UNFORMATTED_BODY));
        session.beginUndoCapture();
        const cell = session.listStories().stories.find((entry) => entry.kind === 'tableCell')!.story;
        const ranges = {
          heading: whole(session, '00000041'),
          paragraphs: whole(session, '00000041', '00000042'),
          cell: whole(session, '00000043', '00000043', cell),
          acrossTheTable: whole(session, '00000042', '00000044'),
        };
        for (const [name, range] of Object.entries(ranges)) {
          const vector = session.encodeStateVector();
          const update = session.encodeStateAsUpdate();
          session.clearFormatting(range, suggesting);
          expect({ name, vector: session.encodeStateVector() }).toEqual({ name, vector });
          expect({ name, update: session.encodeStateAsUpdate() }).toEqual({ name, update });
          expect({ name, canUndo: session.canUndo() }).toEqual({ name, canUndo: false });
        }
        const after = whole(session, '00000044');
        session.formatRange(after, { italic: true });
        const vector = session.encodeStateVector();
        session.clearFormatting(after, suggesting);
        expect(session.encodeStateVector()).not.toEqual(vector);
      } finally {
        session.destroy();
      }
    }
  });
}

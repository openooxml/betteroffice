import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { createEmptyDocx } from '../../packages/docx/src/docx/rezip/createEmpty';
import { applyClipboardFragment } from '../../packages/docx/src/yrs/clipboardFragment';
import { savedRunProperties } from '../../packages/docx/src/yrs/__fixtures__/textRuns';
import type { DocxEditStep } from '../../packages/docx/src/yrs/edits';
import type { TablePayload } from '../../packages/docx/src/yrs/storyPlainText';

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

const BOLD_GRID = '<w:style w:type="table" w:styleId="BoldGrid"><w:name w:val="Bold Grid"/><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="table" w:styleId="HeaderGrid"><w:name w:val="Header Grid"/><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>';

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
    const markers = ['Title', 'Unbolded', 'Bold words', 'Directbold'];
    const cleared = ['Title', 'Unbolded', 'Bold words'];
    const actual: Record<string, unknown>[] = [];
    const expected: Record<string, unknown>[] = [];
    for (const [name, clear] of Object.entries(clears)) {
      expected.push({ name, bold: [true, true, false, true], emptyRunProperties: false, saved: [null, null, null], reopenedBold: [true, true, false, true] });
      const session = await yrs.createYrsSession({ clientId: 91 });
      try {
        session.seedFromDocx(await fixture());
        clear(session);
        const sessionBold = markers.map((marker) => bold(session, marker));
        const { xml, reopened } = await saveAndReopen(yrs, session);
        try {
          actual.push({
            name,
            bold: sessionBold,
            emptyRunProperties: xml.includes('<w:rPr></w:rPr>') || xml.includes('<w:rPr/>'),
            saved: cleared.map((marker) => savedRunProperties(xml, marker)),
            reopenedBold: markers.map((marker) => bold(reopened, marker)),
          });
        } finally {
          reopened.destroy();
        }
      } catch (e) {
        actual.push({ name, error: String(e) });
      } finally {
        session.destroy();
      }
    }
    expect(actual).toEqual(expected);
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

  test(`${engine}: in a header row whose table style bolds the first row the saved bold follows that row`, async () => {
    const cell = (paraId: string, run: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="${paraId}">${run}</w:p></w:tc>`;
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="HeaderGrid"/><w:tblLook w:val="0020" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      `<w:tr>${cell('00000031', '<w:r><w:t>Header off</w:t></w:r>')}${cell('00000032', '<w:r><w:rPr><w:b w:val="0"/><w:i/></w:rPr><w:t>Header clear</w:t></w:r>')}</w:tr>` +
      `<w:tr>${cell('00000033', '<w:r><w:rPr><w:b/></w:rPr><w:t>Body off</w:t></w:r>')}${cell('00000034', '<w:r><w:t>Body plain</w:t></w:r>')}</w:tr>` +
      '</w:tbl><w:p w14:paraId="00000035"><w:r><w:t>After</w:t></w:r></w:p>';
    const markers = ['Header off', 'Header clear', 'Body off'];
    const storyOf = (session: Session, paraId: string) =>
      session.listStories().stories.find((entry) => entry.kind === 'tableCell' && session.paragraphs(entry.story).some((paragraph) => paragraph.paraId === paraId))!.story;
    const boldIn = (session: Session, paraId: string, marker: string) => {
      const segment = session.storySegments(storyOf(session, paraId)).find((entry) => entry.kind === 'text' && entry.text.includes(marker));
      return segment?.kind === 'text' ? Boolean(segment.attributes.bold) : undefined;
    };
    const session = await yrs.createYrsSession({ clientId: 99 });
    try {
      session.seedFromDocx(await fixture(body));
      expect([boldIn(session, '00000031', 'Header off'), boldIn(session, '00000033', 'Body off')]).toEqual([true, true]);
      session.formatRange(whole(session, '00000031', '00000031', storyOf(session, '00000031')), { bold: false });
      session.clearFormatting(whole(session, '00000032', '00000032', storyOf(session, '00000032')));
      session.formatRange(whole(session, '00000033', '00000033', storyOf(session, '00000033')), { bold: false });
      const { xml, reopened } = await saveAndReopen(yrs, session);
      try {
        expect(markers.map((marker) => savedRunProperties(xml, marker))).toEqual(['<w:b w:val="0"/>', null, null]);
        expect([
          boldIn(reopened, '00000031', 'Header off'), boldIn(reopened, '00000032', 'Header clear'), boldIn(reopened, '00000033', 'Body off'),
        ]).toEqual([false, true, false]);
      } finally {
        reopened.destroy();
      }
    } finally {
      session.destroy();
    }
  });

  test(`${engine}: after a table edit the cells of that table reopen with their edited bold`, async () => {
    const cell = (paraId: string, run: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="${paraId}">${run}</w:p></w:tc>`;
    const nested = '<w:tbl><w:tblPr><w:tblStyle w:val="HeaderGrid"/><w:tblLook w:val="0020" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr>${cell('00000046', '<w:r><w:t>Inner head</w:t></w:r>')}</w:tr></w:tbl><w:p w14:paraId="00000047"/>`;
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="HeaderGrid"/><w:tblLook w:val="0020" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      `<w:tr>${cell('00000041', '<w:r><w:t>Old head</w:t></w:r>')}${cell('00000042', '<w:r><w:t>Second head</w:t></w:r>')}</w:tr>` +
      `<w:tr>${cell('00000043', '<w:r><w:rPr><w:b/></w:rPr><w:t>Body one</w:t></w:r>')}<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${nested}</w:tc></w:tr>` +
      '</w:tbl><w:p w14:paraId="00000045"><w:r><w:t>After</w:t></w:r></w:p>';
    const at = (row: number, column: number) => ({ story: 'body', tableIndex: 0, row, column });
    const storyOf = (session: Session, paraId: string) =>
      session.listStories().stories.find((entry) => entry.kind === 'tableCell' && session.paragraphs(entry.story).some((paragraph) => paragraph.paraId === paraId))!.story;
    const format = (session: Session, paraId: string, value: boolean) =>
      session.formatRange(whole(session, paraId, paraId, storyOf(session, paraId)), { bold: value });
    const boldIn = (session: Session, marker: string) => {
      for (const { story } of session.listStories().stories) {
        const segment = session.storySegments(story).find((entry) => entry.kind === 'text' && entry.text.includes(marker));
        if (segment?.kind === 'text') return Boolean(segment.attributes.bold);
      }
      return undefined;
    };
    type Expected = { marker: string; session?: boolean; saved: string | null; reopens: boolean };
    const direct = (marker: string, on: boolean | null): Expected =>
      ({ marker, saved: on === null ? null : on ? '<w:b/>' : '<w:b w:val="0"/>', reopens: on === true });
    const typeIntoInsertedFirstRow = (session: Session) => {
      session.insertRow(at(0, 0), 'above');
      const table = session.tablePayload('body', 0) as unknown as TablePayload;
      const story = table.rows[0]!.cells[0]!.story;
      const paraId = session.paragraphs(story)[0]!.paraId;
      session.insertText({ story, paraId, offset: 0 }, 'New head');
      session.formatRange(whole(session, paraId, paraId, story), { bold: false });
    };
    const cases: [string, (session: Session) => Expected[]][] = [
      ['a row inserted above the header row, the old header cell now in row 1 untouched', (session) => {
        typeIntoInsertedFirstRow(session);
        return [direct('Old head', null)];
      }],
      ['a row inserted above the header row, text typed into a cell of the inserted first row', (session) => {
        typeIntoInsertedFirstRow(session);
        return [{ marker: 'New head', session: false, saved: '<w:b w:val="0"/>', reopens: false }];
      }],
      ['the header row deleted', (session) => {
        session.deleteRow({ anchor: at(0, 0), head: at(0, 0) });
        format(session, '00000043', false);
        return [direct('Body one', false)];
      }],
      ['a column deleted: bold the table style already gives is not written and the first row reopens bold', (session) => {
        session.deleteColumn({ anchor: at(0, 1), head: at(0, 1) });
        format(session, '00000041', true);
        return [{ marker: 'Old head', saved: null, reopens: true }];
      }],
      ['cells merged down the first column', (session) => {
        session.mergeCells({ anchor: at(0, 0), head: at(1, 0) });
        format(session, '00000043', false);
        return [direct('Body one', false)];
      }],
      ['a row inserted in the outer table of a nested one: the inner header cell keeps its own table style, so bold it already gives is not written and it reopens bold', (session) => {
        session.insertRow(at(0, 0), 'above');
        format(session, '00000046', true);
        return [{ marker: 'Inner head', saved: null, reopens: true }];
      }],
    ];
    const actual: Record<string, unknown>[] = [];
    const expected: Record<string, unknown>[] = [];
    for (const [name, edit] of cases) {
      const session = await yrs.createYrsSession({ clientId: 99 });
      try {
        session.seedFromDocx(await fixture(body));
        const wanted = edit(session);
        const first = await saveAndReopen(yrs, session);
        first.reopened.destroy();
        const { xml, reopened } = await saveAndReopen(yrs, session);
        try {
          for (const want of wanted) {
            const { marker } = want;
            const got: Record<string, unknown> = { name, marker, saved: savedRunProperties(xml, marker), reopens: boldIn(reopened, marker) };
            if ('session' in want) got.session = boldIn(session, marker);
            actual.push(got);
            expected.push({ name, ...want });
          }
        } finally {
          reopened.destroy();
        }
      } catch (e) {
        actual.push({ name, error: String(e) });
      } finally {
        session.destroy();
      }
    }
    expect(actual).toEqual(expected);
  });

  test(`${engine}: the old header cell, explicitly not bold and now in row 1 after a row is inserted above it, made bold saves explicit bold`, async () => {
    const cell = (paraId: string, run: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="${paraId}">${run}</w:p></w:tc>`;
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="HeaderGrid"/><w:tblLook w:val="0020" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
      `<w:tr>${cell('00000051', '<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Old head</w:t></w:r>')}</w:tr></w:tbl><w:p w14:paraId="00000052"><w:r><w:t>After</w:t></w:r></w:p>`;
    const actual: Record<string, unknown>[] = [];
    const expected: Record<string, unknown>[] = [];
    for (const saveBetween of [false, true]) {
      expected.push({ saveBetween, saved: '<w:b/>', bold: true });
      const session = await yrs.createYrsSession({ clientId: 99 });
      try {
        session.seedFromDocx(await fixture(body));
        session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
        if (saveBetween) (await saveAndReopen(yrs, session)).reopened.destroy();
        const story = session.listStories().stories.find((entry) => entry.kind === 'tableCell' && session.paragraphs(entry.story).some((paragraph) => paragraph.paraId === '00000051'))!.story;
        session.formatRange(whole(session, '00000051', '00000051', story), { bold: true });
        const { xml, reopened } = await saveAndReopen(yrs, session);
        try {
          const segment = reopened.listStories().stories
            .flatMap((entry) => reopened.storySegments(entry.story))
            .find((entry) => entry.kind === 'text' && entry.text.includes('Old head'));
          actual.push({ saveBetween, saved: savedRunProperties(xml, 'Old head'), bold: segment?.kind === 'text' && Boolean(segment.attributes.bold) });
        } finally {
          reopened.destroy();
        }
      } catch (e) {
        actual.push({ saveBetween, error: String(e) });
      } finally {
        session.destroy();
      }
    }
    expect(actual).toEqual(expected);
  });

  test(`${engine}: a peer bootstrapped from an update keeps the table style's bold when clearing formatting in a cell`, async () => {
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="BoldGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
      '<w:p w14:paraId="00000061"><w:r><w:rPr><w:i/></w:rPr><w:t>Peer cell</w:t></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl><w:p w14:paraId="00000062"><w:r><w:t>After</w:t></w:r></w:p>';
    const bytes = await fixture(body);
    const host = await yrs.createYrsSession({ clientId: 81 });
    const peer = await yrs.createYrsSession({ clientId: 82 });
    try {
      const hostMetadata = host.seedFromDocx(bytes);
      peer.bootstrapPeer(host.encodeState(), host.encodePeerMetadata(), bytes, hostMetadata);
      const cell = peer.listStories().stories.find((entry) => entry.kind === 'tableCell')!.story;
      peer.clearFormatting(whole(peer, '00000061', '00000061', cell));
      const segment = peer.storySegments(cell).find((entry) => entry.kind === 'text' && entry.text.includes('Peer cell'));
      expect(segment?.kind === 'text' && { bold: Boolean(segment.attributes.bold), italic: Boolean(segment.attributes.italic) })
        .toEqual({ bold: true, italic: false });
      const { xml, reopened } = await saveAndReopen(yrs, peer);
      try {
        expect(savedRunProperties(xml, 'Peer cell')).toBeNull();
        const reopenedCell = reopened.listStories().stories.find((entry) => entry.kind === 'tableCell')!.story;
        const saved = reopened.storySegments(reopenedCell).find((entry) => entry.kind === 'text' && entry.text.includes('Peer cell'));
        expect(saved?.kind === 'text' && { bold: Boolean(saved.attributes.bold), italic: Boolean(saved.attributes.italic) })
          .toEqual({ bold: true, italic: false });
      } finally {
        reopened.destroy();
      }
    } finally {
      peer.destroy();
      host.destroy();
    }
  });

  test(`${engine}: a peer bootstrapped from an update keeps the table style's bold when clearing formatting over a body range spanning the table`, async () => {
    const cell = (paraId: string, text: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="${paraId}"><w:r><w:rPr><w:i/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const body = '<w:p w14:paraId="00000071"><w:r><w:rPr><w:b/></w:rPr><w:t>Before table</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblStyle w:val="BoldGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
      `<w:tr>${cell('00000072', 'Peer left')}${cell('00000073', 'Peer right')}</w:tr>` +
      '</w:tbl><w:p w14:paraId="00000074"><w:r><w:rPr><w:b/></w:rPr><w:t>After table</w:t></w:r></w:p>';
    const markers = ['Before table', 'Peer left', 'Peer right', 'After table'];
    const formattingOf = (session: Session, marker: string) => {
      for (const { story } of session.listStories().stories) {
        const segment = session.storySegments(story).find((entry) => entry.kind === 'text' && entry.text.includes(marker));
        if (segment?.kind === 'text') return { bold: Boolean(segment.attributes.bold), italic: Boolean(segment.attributes.italic) };
      }
      return undefined;
    };
    const expected = [
      { bold: false, italic: false },
      { bold: true, italic: false },
      { bold: true, italic: false },
      { bold: false, italic: false },
    ];
    const bytes = await fixture(body);
    const host = await yrs.createYrsSession({ clientId: 83 });
    const peer = await yrs.createYrsSession({ clientId: 84 });
    try {
      const hostMetadata = host.seedFromDocx(bytes);
      peer.bootstrapPeer(host.encodeState(), host.encodePeerMetadata(), bytes, hostMetadata);
      peer.clearFormatting(whole(peer, '00000071', '00000074'));
      expect(markers.map((marker) => formattingOf(peer, marker))).toEqual(expected);
      const { xml, reopened } = await saveAndReopen(yrs, peer);
      try {
        expect(markers.map((marker) => savedRunProperties(xml, marker))).toEqual([null, null, null, null]);
        expect(markers.map((marker) => formattingOf(reopened, marker))).toEqual(expected);
      } finally {
        reopened.destroy();
      }
    } finally {
      peer.destroy();
      host.destroy();
    }
  });

  const italic = (paraId: string, text: string) =>
    `<w:p w14:paraId="${paraId}"><w:r><w:rPr><w:i/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const boldGridCell = (content: string) =>
    '<w:tbl><w:tblPr><w:tblStyle w:val="BoldGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
    `${content}</w:tc></w:tr></w:tbl>`;
  const blockControl = (content: string) =>
    `<w:sdt><w:sdtPr><w:id w:val="11"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const formattingIn = (session: Session, marker: string) => {
    for (const { story } of session.listStories().stories) {
      const segment = session.storySegments(story).find((entry) => entry.kind === 'text' && entry.text.includes(marker));
      if (segment?.kind === 'text') return { bold: Boolean(segment.attributes.bold), italic: Boolean(segment.attributes.italic) };
    }
    return undefined;
  };
  const controlStory = (session: Session) =>
    session.listStories().stories.find((entry) => entry.kind === 'contentControl')!.story;

  const peerClearsInControl = async (body: string, from: string, to: string, markers: string[], clientId: number) => {
    const bytes = await fixture(body);
    const host = await yrs.createYrsSession({ clientId });
    const peer = await yrs.createYrsSession({ clientId: clientId + 1 });
    try {
      const hostMetadata = host.seedFromDocx(bytes);
      peer.bootstrapPeer(host.encodeState(), host.encodePeerMetadata(), bytes, hostMetadata);
      peer.clearFormatting(whole(peer, from, to, controlStory(peer)));
      const inSession = markers.map((marker) => formattingIn(peer, marker));
      const { xml, reopened } = await saveAndReopen(yrs, peer);
      try {
        return {
          story: controlStory(peer),
          session: inSession,
          saved: markers.map((marker) => savedRunProperties(xml, marker)),
          reopened: markers.map((marker) => formattingIn(reopened, marker)),
        };
      } finally {
        reopened.destroy();
      }
    } finally {
      peer.destroy();
      host.destroy();
    }
  };

  test(`${engine}: a peer bootstrapped from an update keeps the table style's bold when clearing formatting inside a block content control in a cell`, async () => {
    const body = boldGridCell(`${blockControl(italic('00000081', 'Control cell'))}<w:p w14:paraId="00000082"/>`) +
      '<w:p w14:paraId="00000083"><w:r><w:t>After</w:t></w:r></w:p>';
    const markers = ['Control cell'];
    const kept = { bold: true, italic: false };
    expect(await peerClearsInControl(body, '00000081', '00000081', markers, 85)).toEqual({
      story: 'body:t0:r0c0:sdt0',
      session: [kept],
      saved: [null],
      reopened: [kept],
    });
  });

  test(`${engine}: a peer bootstrapped from an update keeps the table styles' bold when clearing formatting over the paragraphs of a block content control holding a nested table in a cell`, async () => {
    const control = blockControl(
      italic('00000091', 'Control own') + boldGridCell(italic('00000092', 'Nested cell')) + italic('00000093', 'Control tail'),
    );
    const body = boldGridCell(`${control}<w:p w14:paraId="00000094"/>`) +
      '<w:p w14:paraId="00000095"><w:r><w:t>After</w:t></w:r></w:p>';
    const markers = ['Control own', 'Nested cell', 'Control tail'];
    const kept = { bold: true, italic: false };
    expect(await peerClearsInControl(body, '00000091', '00000093', markers, 87)).toEqual({
      story: 'body:t0:r0c0:sdt0',
      session: [kept, kept, kept],
      saved: [null, null, null],
      reopened: [kept, kept, kept],
    });
  });

  test(`${engine}: known limit: Clear Formatting in a cell resets to the formatting of the row position the cell had when the document was opened`, async () => {
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="HeaderGrid"/><w:tblLook w:val="0020" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="00000071"><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Moved head</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl><w:p w14:paraId="00000072"><w:r><w:t>After</w:t></w:r></w:p>';
    const session = await yrs.createYrsSession({ clientId: 83 });
    try {
      session.seedFromDocx(await fixture(body));
      session.insertRow({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'above');
      const story = session.listStories().stories.find((entry) => entry.kind === 'tableCell' && session.paragraphs(entry.story).some((paragraph) => paragraph.paraId === '00000071'))!.story;
      session.clearFormatting(whole(session, '00000071', '00000071', story));
      const segment = session.storySegments(story).find((entry) => entry.kind === 'text' && entry.text.includes('Moved head'));
      expect(segment?.kind === 'text' && Boolean(segment.attributes.bold)).toBe(true);
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
    const actual: Record<string, unknown>[] = [];
    const expected: Record<string, unknown>[] = [];
    for (const [name, paraId, edit] of edits) {
      const saved: (string | null)[] = [];
      try {
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
        actual.push({ name, withMark: saved[1] });
      } catch (e) {
        actual.push({ name, error: String(e) });
      }
      expected.push({ name, withMark: saved[0] });
    }
    expect(actual).toEqual(expected);
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
    const sameBytes = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((byte, index) => byte === b[index]);
    const actual: Record<string, unknown>[] = [];
    const expected: Record<string, unknown>[] = [];
    for (const suggesting of [undefined, { name: 'Ann', date: '2026-11-10T12:00:00Z' }]) {
      const mode = suggesting ? 'suggesting' : 'direct';
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
          actual.push({ mode, name, vector: session.encodeStateVector(), update: session.encodeStateAsUpdate(), canUndo: session.canUndo() });
          expected.push({ mode, name, vector, update, canUndo: false });
        }
        const after = whole(session, '00000044');
        session.formatRange(after, { italic: true });
        const vector = session.encodeStateVector();
        session.clearFormatting(after, suggesting);
        actual.push({ mode, name: 'formatted', vectorChanged: !sameBytes(session.encodeStateVector(), vector) });
        expected.push({ mode, name: 'formatted', vectorChanged: true });
      } catch (e) {
        actual.push({ mode, error: String(e) });
      } finally {
        session.destroy();
      }
    }
    expect(actual).toEqual(expected);
  });
}

import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { createEmptyDocx } from '../../packages/docx/src/docx/rezip/createEmpty';
import { applyClipboardFragment } from '../../packages/docx/src/yrs/clipboardFragment';
import { savedRunProperties } from '../../packages/docx/src/yrs/__fixtures__/textRuns';

type Yrs = typeof wasmYrs;
type Session = Awaited<ReturnType<Yrs['createYrsSession']>>;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>';
const HEADING = '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>';

async function fixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`, { date });
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', { date });
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', { date });
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}">${STYLES}</w:styles>`, { date });
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
    `<w:p w14:paraId="00000001">${HEADING}<w:r><w:t>Title</w:t></w:r></w:p>` +
    `<w:p w14:paraId="00000002">${HEADING}<w:r><w:rPr><w:b w:val="0"/><w:i/><w:color w:val="C00000"/></w:rPr><w:t>Unbolded</w:t></w:r></w:p>` +
    `<w:p w14:paraId="00000003"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold words</w:t></w:r></w:p>` +
    '<w:sectPr/></w:body></w:document>', { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const whole = (session: Session, from: string, to = from) => {
  const length = session.paragraphs('body').find((entry) => entry.paraId === to)!.text.length;
  return { story: 'body', start: { paraId: from, offset: 0 }, end: { paraId: to, offset: length } };
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
    for (const oneRange of [false, true]) {
      const session = await yrs.createYrsSession({ clientId: 91 });
      try {
        session.seedFromDocx(await fixture());
        if (oneRange) session.clearFormatting(whole(session, '00000001', '00000003'));
        else for (const paraId of ['00000001', '00000002', '00000003']) session.clearFormatting(whole(session, paraId));
        expect([bold(session, 'Title'), bold(session, 'Unbolded'), bold(session, 'Bold words')]).toEqual([true, true, false]);
        const { xml, reopened } = await saveAndReopen(yrs, session);
        try {
          for (const marker of ['Title', 'Unbolded', 'Bold words']) expect(savedRunProperties(xml, marker)).toBeNull();
          expect([bold(reopened, 'Title'), bold(reopened, 'Unbolded'), bold(reopened, 'Bold words')]).toEqual([true, true, false]);
        } finally {
          reopened.destroy();
        }
      } finally {
        session.destroy();
      }
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
}

import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';

type Yrs = typeof wasmYrs;
type Session = Awaited<ReturnType<Yrs['createYrsSession']>>;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
const TEXT_DATA = '<w:ffData><w:name w:val="Text1"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:default w:val="Your name"/><w:maxLength w:val="20"/></w:textInput></w:ffData>';
const CHECK_DATA = '<w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox></w:ffData>';
const DROP_DATA = '<w:ffData><w:name w:val="Drop1"/><w:enabled/><w:calcOnExit w:val="0"/><w:ddList><w:result w:val="1"/><w:listEntry w:val="Alpha"/><w:listEntry w:val="Beta"/><w:listEntry w:val="Gamma"/></w:ddList></w:ffData>';
const DIRTY_BEGIN = '<w:fldChar w:fldCharType="begin" w:dirty="true"/>';

const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const field = (begin: string, instruction: string, result: string) =>
  `<w:r>${begin}</w:r><w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const formBegin = (data: string) => `<w:fldChar w:fldCharType="begin">${data}</w:fldChar>`;

const FORM_PARAGRAPH = `<w:p w14:paraId="73000001">${text('Name: ')}${field(formBegin(TEXT_DATA), 'FORMTEXT', text('Ada'))}${text(' Agree ')}${field(formBegin(CHECK_DATA), 'FORMCHECKBOX', '')}${text(' Pick ')}${field(formBegin(DROP_DATA), 'FORMDROPDOWN', text('Beta'))}${text(' end.')}</w:p>`;
const DIRTY_PARAGRAPH = `<w:p w14:paraId="73000002">${text('Page ')}${field(DIRTY_BEGIN, 'PAGE', text('1'))}${text(' of many')}</w:p>`;
const EXT_NAMESPACE = ' xmlns:ext="urn:example:form-extension"';
const EXT_DATA = '<w:ffData><w:name w:val="Ext1"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/><ext:hint ext:val="x"/></w:ffData>';
const EXT_PARAGRAPH = `<w:p w14:paraId="73000003">${text('Label ')}${field(formBegin(EXT_DATA), 'FORMTEXT', text('Ada'))}</w:p>`;
const OTHER_PARAGRAPH = `<w:p w14:paraId="73000004">${text('Other')}</w:p>`;

async function docx(body: string, namespaces = ''): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"${namespaces}><w:body>${body}${SECTION}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const documentXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');

const ROUTES: Record<string, (yrs: Yrs, session: Session) => Promise<Uint8Array>> = {
  'product save': async (yrs, session) => (await yrs.saveYrsDocx(session, { updateModifiedDate: false })).bytes,
  'whole model rewrite': async (yrs, session) => {
    const base = session.materializeDocx()!;
    const capture = yrs.captureSessionSave(session);
    const document = yrs.yrsToDocument(session, base);
    return (await yrs.writeSessionSave(session, document, capture, base.originalBuffer!, { updateModifiedDate: false }, () => false)).bytes;
  },
};

async function typedAndSaved(yrs: Yrs, paragraph: string, after: string, route: string) {
  const session = await yrs.createYrsSession();
  const reopened = await yrs.createYrsSession();
  try {
    session.openDocx(await docx(paragraph), true);
    const [match] = session.searchText(after, { caseSensitive: true });
    expect(match).toBeDefined();
    session.insertText({ story: 'body', paraId: match!.paraId, offset: match!.end }, 'x');
    const bytes = await ROUTES[route]!(yrs, session);
    reopened.openDocx(bytes, true);
    expect(reopened.searchText(`${after}x`, { caseSensitive: true })).toHaveLength(1);
    const xml = await documentXml(bytes);
    expect(xml).not.toContain(paragraph);
    return xml;
  } finally {
    session.destroy();
    reopened.destroy();
  }
}

async function typedInExtended(yrs: Yrs, after: string, route: string): Promise<string> {
  const session = await yrs.createYrsSession();
  try {
    session.openDocx(await docx(EXT_PARAGRAPH + OTHER_PARAGRAPH, EXT_NAMESPACE), true);
    const [match] = session.searchText(after, { caseSensitive: true });
    expect(match).toBeDefined();
    session.insertText({ story: 'body', paraId: match!.paraId, offset: match!.end }, 'x');
    return documentXml(await ROUTES[route]!(yrs, session));
  } finally {
    session.destroy();
  }
}

function expectInOrder(xml: string, fragments: string[]): void {
  let offset = 0;
  for (const fragment of fragments) {
    const at = xml.indexOf(fragment, offset);
    expect({ fragment, found: at >= 0 }).toEqual({ fragment, found: true });
    offset = at + fragment.length;
  }
}

const count = (xml: string, fragment: string) => xml.split(fragment).length - 1;

for (const [engine, yrs] of [['wasm', wasmYrs], ['native', nativeYrs as unknown as Yrs]] as const) {
  for (const route of Object.keys(ROUTES)) {
    test(`${engine}: keeps form field data when text beside the fields is typed (${route})`, async () => {
      const xml = await typedAndSaved(yrs, FORM_PARAGRAPH, 'Name', route);
      const begins = [TEXT_DATA, CHECK_DATA, DROP_DATA].map(formBegin);
      expect(begins.map((begin) => count(xml, begin))).toEqual([1, 1, 1]);
      expect(count(xml, '<w:ffData>')).toBe(3);
      expectInOrder(xml, ['Name', begins[0]!, ' FORMTEXT ', ' Agree ', begins[1]!, ' FORMCHECKBOX ', ' Pick ', begins[2]!, ' FORMDROPDOWN ', ' end.']);
    });

    test(`${engine}: keeps a dirty field marked dirty when text beside it is typed (${route})`, async () => {
      const xml = await typedAndSaved(yrs, DIRTY_PARAGRAPH, 'Page', route);
      expect(count(xml, DIRTY_BEGIN)).toBe(1);
      expect(count(xml, 'w:dirty=')).toBe(1);
    });

    test(`${engine}: refuses to save, naming the field, when form field data it cannot write is beside the typed text (${route})`, async () => {
      await expect(typedInExtended(yrs, 'Label', route)).rejects.toThrow(/Ext1/);
    });
  }

  test(`${engine}: keeps the source bytes of an untouched paragraph whose form field data it cannot write (product save)`, async () => {
    const xml = await typedInExtended(yrs, 'Other', 'product save');
    expect(count(xml, EXT_PARAGRAPH)).toBe(1);
    expect(xml).toContain('Otherx');
  });

  test(`${engine}: refuses to save, naming the field, when a whole rewrite would drop form field data it cannot write`, async () => {
    await expect(typedInExtended(yrs, 'Other', 'whole model rewrite')).rejects.toThrow(/Ext1/);
  });
}

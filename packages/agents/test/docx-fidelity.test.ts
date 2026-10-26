import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createYrsSession, saveYrsDocx } from '@betteroffice/docx/yrs';
import { openDocx, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const opened: DocxAgentDocument[] = [];
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}

const documentXml = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');

async function saveState(doc: DocxAgentDocument, source: Uint8Array) {
  const draft = await createYrsSession();
  try {
    draft.openDocx(source, false);
    draft.loadState(doc.session.encodeState());
    return (await saveYrsDocx(draft)).bytes;
  } finally { draft.destroy(); }
}

test('agent paragraph edits keep plain and clearing line breaks after a state reload', async () => {
  const zip = await JSZip.loadAsync(await editFixture());
  const untouched = '<w:p><w:r><w:t>Obsolete remark.</w:t></w:r></w:p>';
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:t>Revenue grew strongly in the third quarter.</w:t>',
    '<w:t>Revenue grew strongly</w:t><w:br/><w:t>again</w:t><w:br w:type="textWrapping" w:clear="all"/><w:t>today</w:t>'));
  const source = await zip.generateAsync({ type: 'uint8array' });
  const doc = await open(source);
  const target = doc.grep({ query: 'Revenue grew' }).matches[0];
  await doc.edit({ version: doc.list().version, edits: [{ op: 'set_paragraph_style', paragraph: target.ref, style: 'Quote' }] });
  for (const bytes of [await doc.export(), await saveState(doc, source)]) {
    const saved = await documentXml(bytes);
    expect(saved).toContain('<w:br/>');
    expect(saved).not.toContain('<w:br w:type="textWrapping"/>');
    expect(saved).toContain('<w:br w:type="textWrapping" w:clear="all"/>');
    expect(saved).toContain(untouched);
  }
});

import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx, type DocxAgentDocument } from '../src';
import { editFixture } from './fixture';

const HEADING = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>';
async function fixture(body: string) {
  const zip = await JSZip.loadAsync(await editFixture());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*?(?=<w:sectPr>)/, `<w:body>${body}`));
  zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace('</w:styles>', `${HEADING}</w:styles>`));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const opened: DocxAgentDocument[] = [];
async function open(bytes: Uint8Array) {
  const doc = await openDocx(bytes);
  opened.push(doc);
  return doc;
}
afterEach(() => { opened.splice(0).forEach(doc => doc.close()); });

const segment = (doc: DocxAgentDocument, text: string) => {
  const found = doc.session.storySegments('body').find(entry => entry.kind === 'text' && entry.text.includes(text));
  return found?.kind === 'text' ? found.attributes : {};
};
async function runProperties(doc: DocxAgentDocument, text: string) {
  const xml = await (await JSZip.loadAsync(await doc.export())).file('word/document.xml')!.async('string');
  const run = (xml.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).find(entry => entry.includes(text));
  return run?.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '';
}
async function format(doc: DocxAgentDocument, query: string, marks: Record<string, unknown>) {
  const match = doc.grep({ query }).matches[0].match;
  return doc.edit({ version: doc.list().version, edits: [{ op: 'format_text', match, ...marks }] });
}

test('format_text around a tracked insertion behaves as on main: no match spans it, a neighbouring match formats as the editor does', async () => {
  const body = '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:ins w:id="7" w:author="Ann" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>';
  const tool = await open(await fixture(body));
  const direct = await open(await fixture(body));
  expect(tool.grep({ query: 'Before inserted' }).matches).toEqual([]);
  const result = await format(tool, 'Before', { bold: true });
  expect(JSON.stringify(result)).not.toContain('error');
  const paraId = direct.session.paragraphs('body')[0].paraId;
  direct.session.formatRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 6 } }, { bold: true });
  expect(tool.session.storySegments('body')).toEqual(direct.session.storySegments('body'));
  expect(Boolean(segment(tool, 'inserted').bold)).toBe(false);
  expect(Boolean(segment(tool, 'inserted').ins)).toBe(true);
});

test('format_text writes the new marks', async () => {
  const doc = await open(await fixture('<w:p><w:r><w:t>Plain words</w:t></w:r></w:p>'));
  await format(doc, 'Plain', { strike: 'dblStrike', verticalAlign: 'superscript', color: 'C00000', highlight: 'yellow', fontSizePt: 10.5, fontFamily: 'Arial' });
  const properties = await runProperties(doc, 'Plain');
  for (const expected of ['<w:dstrike/>', '<w:vertAlign w:val="superscript"/>', 'w:val="C00000"', '<w:highlight w:val="yellow"/>', '<w:sz w:val="21"/>', 'w:ascii="Arial"']) {
    expect(properties).toContain(expected);
  }
  expect(Boolean(segment(doc, ' words').superscript)).toBe(false);
});

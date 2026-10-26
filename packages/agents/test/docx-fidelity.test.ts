import { afterEach, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createYrsSession, saveYrsDocx } from '@betteroffice/docx/yrs';
import { openDocx, type DocxAgentDocument } from '../src';
import { commentFixture, editFixture, tableCell, tableFixture } from './fixture';

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

test('merging cells preserves automatic shading in every retained cell', async () => {
  const properties = '<w:tcW w:w="1775" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/>';
  const lower = [tableCell('Lower left', properties), tableCell('Lower right', properties)];
  const source = await tableFixture([
    tableCell('Upper left', properties) + tableCell('Upper right', properties),
    lower.join(''),
  ]);
  const doc = await open(source);
  const from = doc.grep({ query: 'Upper left' }).matches[0].ref;
  const to = doc.grep({ query: 'Upper right' }).matches[0].ref;
  await doc.edit({ version: doc.list().version, edits: [{ op: 'merge_cells', from, to }] });
  for (const bytes of [await doc.export(), await saveState(doc, source)]) {
    const saved = await documentXml(bytes);
    expect(saved.match(/<w:shd w:val="clear" w:color="auto" w:fill="auto"\/>/g)).toHaveLength(3);
    expect(saved).toContain('<w:tcPr><w:tcW w:w="1775" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr>');
    for (const cell of lower) expect(saved).toContain(cell);
  }
});

test('reply_comment keeps the parent reference style and styles its own reference', async () => {
  const zip = await JSZip.loadAsync(await commentFixture());
  const styles = await zip.file('word/styles.xml')!.async('string');
  zip.file('word/styles.xml', styles.replace('</w:styles>', '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="Comment Reference"/></w:style></w:styles>'));
  const source = await zip.generateAsync({ type: 'uint8array' });
  const doc = await open(source);
  await doc.edit({ version: doc.list().version, edits: [{ op: 'reply_comment', comment: '0', author: 'Editor', text: 'Follow up' }] });
  for (const bytes of [await doc.export(), await doc.export(), await saveState(doc, source)]) {
    const saved = await documentXml(bytes);
    const reopened = await open(bytes);
    const id = reopened.list({ story: 'comments' }).items.find(comment => comment.parentId === '0')!.id;
    for (const comment of ['0', id]) {
      expect(saved).toContain(`<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${comment}"/></w:r>`);
      expect(saved.match(new RegExp(`<w:commentReference w:id="${comment}"/>`, 'g'))).toHaveLength(1);
    }
    expect(saved).toContain('<w:p><w:r><w:t>Obsolete remark.</w:t></w:r></w:p>');
  }
});

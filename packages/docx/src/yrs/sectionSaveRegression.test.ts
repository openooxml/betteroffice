import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from '../docx';
import type { SectionProperties } from '../types/document';
import { saveEditorDocument } from './editorSave';
import { createYrsSession, saveYrsDocx } from './index';
import { createResidentEngineSession } from './residentEngineSession';
import { sectionPackage } from './sectionTestPackage';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function sourcePackage(boundary = false): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await sectionPackage(boundary));
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml
    .replace('<w:r><w:t>First</w:t></w:r>', '<w:commentRangeStart w:id="1"/><w:r><w:t>First</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>')
    .replaceAll('<w:sectPr>', '<w:sectPr xmlns:u="urn:unknown"><u:opaque u:v=\'kept\'/>'));
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="${R}/comments" Target="comments.xml"/></Relationships>`);
  zip.file('word/comments.xml', `<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Reviewer"><w:p><w:r><w:t>Original comment</w:t></w:r></w:p></w:comment></w:comments>`);
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'));
  return zip.generateAsync({ type: 'uint8array' });
}

test('section saves keep comment body edits without moving anchors', async () => {
  const session = await createYrsSession();
  session.seedFromDocx(await sourcePackage());
  try {
    const anchors = session.resolveComment('1');
    session.applyRawOps('body', [{
      op: 'setComment', id: '1', body: 'Changed comment',
      ranges: anchors.map(anchor => [anchor.start, anchor.end] as const),
    }]);
    session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
    expect(session.resolveComment('1')).toEqual(anchors);
    const saved = await JSZip.loadAsync((await saveYrsDocx(session)).bytes);
    expect(await saved.file('word/comments.xml')!.async('string')).toContain('Changed comment');
    expect(await saved.file('word/document.xml')!.async('string')).toContain('w:val="continuous"');
  } finally { session.destroy(); }
});

test('editor section saves keep host comment metadata edits', async () => {
  const session = await createYrsSession();
  session.seedFromDocx(await sourcePackage());
  try {
    session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
    const projected = yrsToDocument(session, session.materializeDocx()!);
    const comments = projected.package.document.comments!.map(comment => ({ ...comment, author: 'Changed author' }));
    const saved = await JSZip.loadAsync(await saveEditorDocument(session, projected, comments));
    expect(await saved.file('word/comments.xml')!.async('string')).toContain('Changed author');
    expect(await saved.file('word/document.xml')!.async('string')).toContain('w:val="continuous"');
  } finally { session.destroy(); }
});

test('editor section saves keep numbering metadata additions', async () => {
  const session = await createYrsSession();
  session.seedFromDocx(await sectionPackage());
  try {
    session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
    const projected = yrsToDocument(session, session.materializeDocx()!);
    projected.package.numbering = {
      abstractNums: [{ abstractNumId: 9, levels: [] }],
      nums: [{ numId: 9, abstractNumId: 9 }],
    };
    const saved = await JSZip.loadAsync(await saveEditorDocument(session, projected, []));
    expect(await saved.file('word/numbering.xml')!.async('string')).toContain('w:abstractNumId="9"');
    expect(await saved.file('word/document.xml')!.async('string')).toContain('w:val="continuous"');
  } finally { session.destroy(); }
});

for (const boundary of [false, true]) {
  test(`editor text and ${boundary ? 'inline' : 'final'} page setup keep safe paragraph splices`, async () => {
    const source = await sourcePackage(boundary);
    const session = await createYrsSession();
    session.seedFromDocx(source);
    try {
      const paragraphs = session.paragraphs('body');
      session.replaceRange({ story: 'body', start: { paraId: paragraphs[0].paraId, offset: 0 }, end: { paraId: paragraphs[0].paraId, offset: 5 } }, 'Edited');
      const section = boundary ? paragraphs[1].paraId : 'final';
      session.setSectionProperties(section, { orientation: 'landscape', margins: { left: 1800 } });
      const projected = yrsToDocument(session, session.materializeDocx()!);
      const buffer = await saveEditorDocument(session, projected, projected.package.document.comments!);
      const saved = await JSZip.loadAsync(buffer);
      const xml = await saved.file('word/document.xml')!.async('string');
      expect(xml).toContain('Edited');
      expect(xml.match(/<u:opaque u:v='kept'\/>/g)).toHaveLength(boundary ? 2 : 1);
      const original = await JSZip.loadAsync(source);
      const originalXml = await original.file('word/document.xml')!.async('string');
      expect(xml).toContain(originalXml.match(/<w:p w14:paraId="10000003">.*?<\/w:p>/)![0]);
      const reopened = await parseDocx(buffer, { preloadFonts: false });
      const properties = boundary
        ? reopened.package.document.sections![0].properties
        : reopened.package.document.finalSectionProperties;
      expect(properties).toMatchObject({ orientation: 'landscape', marginLeft: 1800 });
    } finally { session.destroy(); }
  });
}

const residentPatches: SectionProperties[] = [
  { columnCount: 2, columnSpace: 720, equalWidth: true, separator: true },
  { footnotePr: { numStart: 3, numRestart: 'eachSect' } },
  { endnotePr: { numStart: 4, numRestart: 'eachSect' } },
  { pageWidth: 14000, columnCount: 2 },
];

for (const boundary of [false, true]) {
  for (const patch of residentPatches) {
    test(`resident ${boundary ? 'inline' : 'final'} section saves retain ${Object.keys(patch).join(', ')}`, async () => {
      const source = await sectionPackage(boundary);
      const session = await createResidentEngineSession();
      try {
        const hostJson = session.openDocx(source);
        const document = structuredClone(session.propertyDocument());
        if (boundary) {
          document.package.document.sections![0].properties = { ...document.package.document.sections![0].properties, ...patch };
        } else {
          document.package.document.finalSectionProperties = { ...document.package.document.finalSectionProperties, ...patch };
          const last = document.package.document.sections?.at(-1);
          if (last) last.properties = document.package.document.finalSectionProperties;
        }
        expect(session.setPropertyDocument(document)).toBe(true);
        const saved = await session.save(source, hostJson, undefined, [], { full: false });
        const reopened = await parseDocx(saved, { preloadFonts: false });
        expect(boundary ? reopened.package.document.sections![0].properties : reopened.package.document.finalSectionProperties).toMatchObject(patch);
      } finally { session.destroy(); }
    });
  }
}

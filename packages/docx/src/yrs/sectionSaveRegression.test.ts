import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import type { SectionProperties } from '../types/document';
import { saveEditorDocument } from './editorSave';
import { documentToYrs } from './documentToYrs';
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

async function tableSectionPackage(opaque = false, localPrefix = false): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await sectionPackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  const table = `<w:tbl${localPrefix ? ' xmlns:x="urn:custom"' : ''}>` +
    `<w:tblPr${localPrefix ? ' x:keep="1"' : ''}/>` +
    '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr/>' +
    '<w:tc><w:tcPr/><w:p w14:paraId="10000004"><w:r><w:t>Cell</w:t></w:r></w:p>' +
    '</w:tc></w:tr></w:tbl>';
  zip.file('word/document.xml', xml
    .replace('<w:p w14:paraId="10000002">', `${table}<w:p w14:paraId="10000002">`)
    .replace('<w:sectPr>', opaque ? '<w:sectPr xmlns:u="urn:unknown"><u:opaque/>' : '<w:sectPr>'));
  return zip.generateAsync({ type: 'uint8array' });
}

for (const boundary of [false, true]) {
  test(`simple table shading and ${boundary ? 'a new section boundary' : 'final page setup'} save together`, async () => {
    const session = await createYrsSession();
    session.seedFromDocx(await tableSectionPackage());
    try {
      const cell = { story: 'body', tableIndex: 0, row: 0, column: 0 };
      session.setCellShading({ anchor: cell, head: cell }, '112233');
      if (boundary) session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
      else session.setSectionProperties('final', { margins: { left: 1800 } });
      const splice = session.paragraphSavePlan().splicedParts?.find(({ part }) => part === 'word/document.xml');
      expect(splice?.tablePatches).toHaveLength(1);
      expect(splice!.changed.every(ordinal => splice!.sectionOnly?.includes(ordinal))).toBe(true);
      const saved = await JSZip.loadAsync((await saveYrsDocx(session)).bytes);
      const xml = await saved.file('word/document.xml')!.async('string');
      expect(xml).toContain('w:fill="112233"');
      expect(xml).toContain('>Cell<');
      expect(xml.match(/<w:sectPr\b/g)).toHaveLength(boundary ? 2 : 1);
      expect(xml).toContain(boundary ? 'w:val="continuous"' : 'w:left="1800"');
    } finally { session.destroy(); }
  });
}

test('section splice plans carry simple table shading beside text and opaque section XML', async () => {
  const session = await createYrsSession();
  session.seedFromDocx(await tableSectionPackage(true));
  try {
    const cell = { story: 'body', tableIndex: 0, row: 0, column: 0 };
    session.setCellShading({ anchor: cell, head: cell }, '112233');
    session.insertText({ story: 'body', paraId: session.paragraphs('body')[0].paraId, offset: 0 }, 'Edited ');
    session.setSectionProperties('final', { margins: { left: 1800 } });
    const plan = session.paragraphSavePlan();
    expect(plan.splicedParts?.find(({ part }) => part === 'word/document.xml')).toBeUndefined();
    expect(plan.sectionSplicedParts?.find(({ part }) => part === 'word/document.xml')?.tablePatches).toHaveLength(1);
    const saved = await JSZip.loadAsync((await saveYrsDocx(session)).bytes);
    const xml = await saved.file('word/document.xml')!.async('string');
    expect(xml.replace(/<[^>]*>/g, '')).toContain('Edited First');
    expect(xml).toContain('w:fill="112233"');
    expect(xml).toContain('w:left="1800"');
    expect(xml).toContain('<u:opaque/>');
  } finally { session.destroy(); }
});

for (const editor of [false, true]) {
  test(`${editor ? 'editor' : 'session'} refuses a lossy new boundary when a simple table patch is rejected`, async () => {
    const session = await createYrsSession();
    session.seedFromDocx(await tableSectionPackage(true, true));
    try {
      session.setTableWidth({ story: 'body', tableIndex: 0 }, 5000);
      session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
      expect(session.paragraphSavePlan().sectionSplicedParts?.find(
        ({ part }) => part === 'word/document.xml'
      )?.tablePatches).toHaveLength(1);
      const before = session.encodeState();
      const projected = yrsToDocument(session, session.materializeDocx()!);
      await expect(editor ? saveEditorDocument(session, projected, []) : saveYrsDocx(session)).rejects.toMatchObject({
        name: 'DocxSectionError', code: 'unsupported',
      });
      expect(session.encodeState()).toEqual(before);
    } finally { session.destroy(); }
  });
}

test('projected seeding leaves source section metadata intact through body edits', async () => {
  const source = await parseDocx(await sectionPackage(true), { preloadFonts: false });
  const session = await createYrsSession();
  try {
    documentToYrs(session, source);
    session.insertText({ story: 'body', paraId: session.paragraphs('body')[0].paraId, offset: 0 }, 'Edited ');
    const projected = yrsToDocument(session, source);
    expect(projected.package.document.finalSectionProperties).toBe(source.package.document.finalSectionProperties);
    expect(projected.package.document.sections).toBe(source.package.document.sections);
    const zip = await JSZip.loadAsync(await repackDocx(projected));
    expect((await zip.file('word/document.xml')!.async('string')).match(/<w:sectPr\b/g)).toHaveLength(2);
  } finally { session.destroy(); }
});

test('host page setup uses the ordinary save projection beside opaque body XML', async () => {
  const zip = await JSZip.loadAsync(await sectionPackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:t>First</w:t>', '<u:opaque xmlns:u="urn:unknown"/><w:t>First</w:t>'));
  const session = await createYrsSession();
  session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const projected = yrsToDocument(session, session.materializeDocx()!);
    projected.package.document.finalSectionProperties = { ...projected.package.document.finalSectionProperties, marginTop: 2880 };
    const expected = await JSZip.loadAsync(await repackDocx(projected));
    const saved = await JSZip.loadAsync(await saveEditorDocument(session, projected, []));
    expect(await saved.file('word/document.xml')!.async('string')).toBe(await expected.file('word/document.xml')!.async('string'));
  } finally { session.destroy(); }
});

test('text and final page setup beside opaque body XML fall back to the ordinary save', async () => {
  const zip = await JSZip.loadAsync(await sectionPackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:t>First</w:t>', '<u:opaque xmlns:u="urn:unknown"/><w:t>First</w:t>'));
  const session = await createYrsSession();
  session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    session.insertText({ story: 'body', paraId: session.paragraphs('body')[1].paraId, offset: 0 }, 'Edited ');
    session.setSectionProperties('final', { margins: { left: 2160 } });
    const saved = (await saveYrsDocx(session)).bytes;
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(JSON.stringify(reopened.package.document.content)).toContain('Edited ');
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 2160 });
  } finally { session.destroy(); }
});

test('a header edit beside opaque body XML keeps the final page setup on the ordinary save', async () => {
  const zip = await JSZip.loadAsync(await sectionPackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml
    .replace('<w:t>First</w:t>', '<u:opaque xmlns:u="urn:unknown"/><w:t>First</w:t>')
    .replace(/(<w:document [^>]*)>/, `$1 xmlns:r="${R}">`)
    .replace(/<w:sectPr>(?![\s\S]*<w:sectPr>)/, '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>'));
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/></Relationships>`);
  zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`);
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  const session = await createYrsSession();
  session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const header = session.storyIds().find(story => session.paragraphs(story).some(paragraph => paragraph.text === 'Header'))!;
    session.insertText({ story: header, paraId: session.paragraphs(header)[0].paraId, offset: 0 }, 'Edited ');
    session.setSectionProperties('final', { margins: { left: 2160 } });
    const saved = await JSZip.loadAsync((await saveYrsDocx(session)).bytes);
    expect(await saved.file('word/header1.xml')!.async('string')).toContain('Edited ');
    expect(await saved.file('word/document.xml')!.async('string')).toContain('w:left="2160"');
  } finally { session.destroy(); }
});

for (const editor of [false, true]) {
  test(`${editor ? 'editor' : 'session'} saves a new modeled section boundary beside a table insertion through the ordinary save`, async () => {
    const session = await createYrsSession();
    session.seedFromDocx(await sectionPackage());
    try {
      session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
      session.insertTable({ story: 'body', paraId: session.paragraphs('body')[1].paraId, offset: 0 }, 2, 2);
      const projected = yrsToDocument(session, session.materializeDocx()!);
      const saved = editor ? await saveEditorDocument(session, projected, []) : (await saveYrsDocx(session)).bytes;
      const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
      expect(xml).toContain('<w:tbl>');
      expect(xml.match(/<w:sectPr\b/g)).toHaveLength(2);
      expect(xml.match(/<w:pgSz w:w="12240"/g)).toHaveLength(2);
      const reopened = await parseDocx(saved, { preloadFonts: false });
      expect(reopened.package.document.finalSectionProperties).toMatchObject({ sectionStart: 'continuous' });
    } finally { session.destroy(); }
  });
}

test('a split boundary clones the section before the next source paragraph', async () => {
  const zip = await JSZip.loadAsync(await sectionPackage(true));
  const xml = await zip.file('word/document.xml')!.async('string');
  const raw = '<w:sectPr xmlns:u="urn:unknown"><u:opaque/></w:sectPr>';
  zip.file('word/document.xml', xml.replace(/<w:sectPr>.*?<\/w:sectPr>/, raw));
  const session = await createYrsSession();
  session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    session.insertSectionBreak({ story: 'body', paraId: session.paragraphs('body')[1].paraId, offset: 3 }, 'continuous');
    const saved = await JSZip.loadAsync((await saveYrsDocx(session)).bytes);
    const result = await saved.file('word/document.xml')!.async('string');
    expect(result.match(/<u:opaque\/>/g)).toHaveLength(2);
    expect(result).toContain(raw);
    expect(result).toContain('w:val="continuous"');
    expect(result).toContain('<w:pgSz w:w="12240" w:h="15840"/>');
  } finally { session.destroy(); }
});

for (const editor of [false, true]) {
  test(`${editor ? 'editor' : 'session'} saves text and page setup without source paragraphs through the ordinary save`, async () => {
    const zip = await JSZip.loadAsync(await sectionPackage());
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace(/<w:p\b[^>]*>.*?<\/w:p>/g, ''));
    const session = await createYrsSession();
    session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
    try {
      expect(session.paragraphIdentities().paragraphs.every(identity => identity.source === null)).toBe(true);
      session.insertText({ story: 'body', paraId: session.paragraphs('body')[0].paraId, offset: 0 }, 'Authored text');
      session.setSectionProperties('final', { orientation: 'landscape' });
      const projected = yrsToDocument(session, session.materializeDocx()!);
      const bytes = editor ? await saveEditorDocument(session, projected, []) : (await saveYrsDocx(session)).bytes;
      const reopened = await parseDocx(bytes, { preloadFonts: false });
      expect(JSON.stringify(reopened.package.document.content)).toContain('Authored text');
      expect(reopened.package.document.finalSectionProperties).toMatchObject({ orientation: 'landscape' });
    } finally { session.destroy(); }
  });
}

test('final page setup after a table insertion saves through the ordinary save', async () => {
  const session = await createYrsSession();
  session.seedFromDocx(await sectionPackage());
  try {
    session.insertTable({ story: 'body', paraId: session.paragraphs('body')[0].paraId, offset: 0 }, 2, 2);
    session.setSectionProperties('final', { margins: { left: 1800 } });
    const saved = (await saveYrsDocx(session)).bytes;
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:tbl>');
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 1800 });
  } finally { session.destroy(); }
});

test('resident saves mix unsupported inline and supported final section changes', async () => {
  const source = await sectionPackage(true);
  const session = await createResidentEngineSession();
  try {
    const hostJson = session.openDocx(source);
    const document = structuredClone(session.propertyDocument());
    const body = document.package.document;
    body.sections![0].properties = { ...body.sections![0].properties, columnCount: 2 };
    body.finalSectionProperties = { ...body.finalSectionProperties, marginLeft: 1800 };
    const last = body.sections?.at(-1);
    if (last) last.properties = body.finalSectionProperties;
    expect(session.setPropertyDocument(document)).toBe(true);
    const reopened = await parseDocx(await session.save(source, hostJson, undefined, [], { full: false }), { preloadFonts: false });
    expect(reopened.package.document.sections![0].properties).toMatchObject({ columnCount: 2 });
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 1800 });
  } finally { session.destroy(); }
});

test('resident saves round fractional section margins', async () => {
  const source = await sectionPackage();
  const session = await createResidentEngineSession();
  try {
    const hostJson = session.openDocx(source);
    const document = structuredClone(session.propertyDocument());
    const body = document.package.document;
    body.finalSectionProperties = { ...body.finalSectionProperties, marginLeft: 1440.5 };
    const last = body.sections?.at(-1);
    if (last) last.properties = body.finalSectionProperties;
    expect(session.setPropertyDocument(document)).toBe(true);
    const reopened = await parseDocx(await session.save(source, hostJson, undefined, [], { full: false }), { preloadFonts: false });
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 1441 });
  } finally { session.destroy(); }
});

test('editor comment saves keep host section changes beside a session page setup', async () => {
  const session = await createYrsSession();
  session.seedFromDocx(await sourcePackage());
  try {
    session.setSectionProperties('final', { orientation: 'landscape' });
    const projected = yrsToDocument(session, session.materializeDocx()!);
    projected.package.document.finalSectionProperties = { ...projected.package.document.finalSectionProperties, columnCount: 2 };
    const reopened = await parseDocx(await saveEditorDocument(session, projected, projected.package.document.comments!), { preloadFonts: false });
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ orientation: 'landscape', columnCount: 2 });
  } finally { session.destroy(); }
});

test('editor comment saves keep host section deletions beside a session page setup', async () => {
  const zip = await JSZip.loadAsync(await sourcePackage());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('</w:sectPr>', '<w:cols w:num="2"/></w:sectPr>'));
  const session = await createYrsSession();
  session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    session.setSectionProperties('final', { orientation: 'landscape' });
    const projected = yrsToDocument(session, session.materializeDocx()!);
    expect(projected.package.document.finalSectionProperties?.columnCount).toBe(2);
    delete projected.package.document.finalSectionProperties!.columnCount;
    const reopened = await parseDocx(await saveEditorDocument(session, projected, projected.package.document.comments!), { preloadFonts: false });
    expect(reopened.package.document.finalSectionProperties).toMatchObject({ orientation: 'landscape' });
    expect(reopened.package.document.finalSectionProperties?.columnCount ?? 1).toBe(1);
  } finally { session.destroy(); }
});

for (const editor of [false, true]) {
  test(`${editor ? 'editor' : 'session'} refuses a new section boundary beside a table insertion`, async () => {
    const zip = await JSZip.loadAsync(await sectionPackage());
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace('<w:sectPr>', '<w:sectPr><w:textDirection w:val="tbRl"/>'));
    const session = await createYrsSession();
    session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
    try {
      session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
      session.insertTable({ story: 'body', paraId: session.paragraphs('body')[1].paraId, offset: 0 }, 2, 2);
      const before = session.encodeState();
      const projected = yrsToDocument(session, session.materializeDocx()!);
      await expect(editor ? saveEditorDocument(session, projected, []) : saveYrsDocx(session)).rejects.toMatchObject({
        name: 'DocxSectionError', code: 'unsupported',
      });
      expect(session.encodeState()).toEqual(before);
    } finally { session.destroy(); }
  });
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

for (const inline of [false, true]) {
  for (const patch of residentPatches.slice(0, 3)) {
    test(`editor refuses a new section with ${inline ? 'inline' : 'final'} ${Object.keys(patch).join(', ')}`, async () => {
      const zip = await JSZip.loadAsync(await sectionPackage());
      const xml = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', xml.replace('<w:sectPr>', '<w:sectPr><w:textDirection w:val="tbRl"/>'));
      const session = await createYrsSession();
      session.seedFromDocx(await zip.generateAsync({ type: 'uint8array' }));
      try {
        session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
        const projected = yrsToDocument(session, session.materializeDocx()!);
        const body = projected.package.document;
        if (inline) {
          const first = body.content[0];
          if (first.type !== 'paragraph') throw new Error('Missing section paragraph');
          first.sectionProperties = { ...first.sectionProperties, ...patch };
          body.sections![0].properties = first.sectionProperties;
        } else {
          body.finalSectionProperties = { ...body.finalSectionProperties, ...patch };
          body.sections!.at(-1)!.properties = body.finalSectionProperties;
        }
        const before = session.encodeState();
        await expect(saveEditorDocument(session, projected, [])).rejects.toMatchObject({
          name: 'DocxSectionError', code: 'unsupported',
        });
        expect(session.encodeState()).toEqual(before);
      } finally { session.destroy(); }
    });
  }
}

for (const boundary of [false, true]) {
  test(`resident text saves after ${boundary ? 'inline' : 'final'} page setup replay source splices`, async () => {
    const source = await sourcePackage(boundary);
    const session = await createResidentEngineSession();
    try {
      const hostJson = session.openDocx(source);
      const comments = session.propertyDocument().package.document.comments ?? [];
      const section = boundary ? session.sectionProperties().sections[0].section : 'final';
      expect(session.setSectionMargins(section, { marginLeft: 1800 })).toBe(true);
      const record = { full: false };
      await session.save(source, hostJson, undefined, comments, record);
      const paraId = session.paragraphs('body')[0].paraId;
      let previous = 'First';
      for (const text of ['Edited', 'Again']) {
        session.replaceRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: previous.length } }, text);
        const saved = await session.save(source, hostJson, undefined, comments, record);
        const zip = await JSZip.loadAsync(saved);
        const xml = await zip.file('word/document.xml')!.async('string');
        expect(xml).toContain(text);
        expect(xml.match(/<u:opaque u:v='kept'\/>/g)).toHaveLength(boundary ? 2 : 1);
        expect(xml).toContain('<w:p w14:paraId="10000003"><w:r><w:t>Third</w:t></w:r></w:p>');
        const reopened = await parseDocx(saved, { preloadFonts: false });
        expect(boundary ? reopened.package.document.sections![0].properties : reopened.package.document.finalSectionProperties).toMatchObject({ marginLeft: 1800 });
        previous = text;
      }
    } finally { session.destroy(); }
  });
}

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

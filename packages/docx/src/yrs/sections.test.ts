import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';

import { sectionPackage } from './sectionTestPackage';

async function open(bytes?: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession();
  session.seedFromDocx(bytes ?? await sectionPackage());
  return session;
}

for (const kind of ['nextPage', 'continuous', 'evenPage', 'oddPage'] as const) {
  test(`section break after a paragraph preserves its text and starts ${kind}`, async () => {
    const session = await open();
    try {
      const before = session.paragraphs('body');
      session.insertSectionBreakAfter(before[0].paraId, kind);
      expect(session.paragraphs('body').map(p => p.text)).toEqual(before.map(p => p.text));
      const sections = session.listSections();
      expect(sections.map(s => [s.firstParagraph, s.lastParagraph])).toEqual([[1, 1], [2, 3]]);
      expect(sections[0].closingParagraphId).toBe(before[0].paraId);
      expect(sections[1].closingParagraphId).toBe('final');
      expect(sections[1].authored.sectionStart).toBe(kind);
      const reopened = await open((await saveYrsDocx(session)).bytes);
      try { expect(reopened.listSections()[1].authored.sectionStart).toBe(kind); }
      finally { reopened.destroy(); }
    } finally { session.destroy(); }
  });
}

test('insert before an existing boundary clones that section and preserves the next start', async () => {
  const session = await open(await sectionPackage(true));
  try {
    const paragraphs = session.paragraphs('body');
    session.insertSectionBreakAfter(paragraphs[0].paraId, 'continuous');
    expect(session.listSections().map(s => [s.firstParagraph, s.lastParagraph])).toEqual([[1, 1], [2, 2], [3, 3]]);
    expect(session.listSections()[1].authored.sectionStart).toBe('continuous');
    expect(() => session.insertSectionBreakAfter(paragraphs[1].paraId, 'oddPage')).toThrow();
  } finally { session.destroy(); }
});

test('final-only page setup survives save and orientation repeats are no-ops', async () => {
  const source = await sectionPackage();
  const session = await open(source);
  try {
    session.setSectionProperties('final', { orientation: 'landscape', margins: { left: 1800, header: 600 } });
    expect(session.listSections()[0].authored).toMatchObject({ pageWidth: 15840, pageHeight: 12240, marginLeft: 1800, headerDistance: 600 });
    const version = session.version();
    session.setSectionProperties('final', { orientation: 'landscape' });
    expect(session.version()).toBe(version);
    const saved = (await saveYrsDocx(session)).bytes;
    const reopened = await open(saved);
    try { expect(reopened.listSections()[0].authored).toEqual(session.listSections()[0].authored); }
    finally { reopened.destroy(); }
    const originalZip = await JSZip.loadAsync(source);
    const savedZip = await JSZip.loadAsync(saved);
    for (const path of Object.keys(originalZip.files).filter(path => !originalZip.files[path].dir && path !== 'word/document.xml' && path !== 'docProps/core.xml')) {
      expect(await savedZip.file(path)!.async('uint8array')).toEqual(await originalZip.file(path)!.async('uint8array'));
    }
  } finally { session.destroy(); }
});

test('section edits undo, redo and replicate as whole values on a fork', async () => {
  const source = await sectionPackage();
  const session = await open(source);
  const peer = await open(source);
  try {
    const id = session.paragraphs('body')[0].paraId;
    session.insertSectionBreakAfter(id, 'evenPage');
    expect(session.undo()).toBe(true);
    expect(session.listSections()).toHaveLength(1);
    expect(session.redo()).toBe(true);
    peer.loadState(session.encodeState());
    peer.setSectionProperties('final', { pageWidth: 14000, margins: { gutter: 120 } });
    session.applyUpdate(peer.encodeStateAsUpdate(session.encodeStateVector()));
    expect(session.listSections()).toEqual(peer.listSections());
  } finally { session.destroy(); peer.destroy(); }
});

test('invalid section patches and non-body targets refuse without any mutation', async () => {
  const session = await open();
  try {
    const before = session.encodeState();
    expect(() => session.insertSectionBreakAfter('missing', 'nextPage')).toThrow();
    expect(() => session.setSectionProperties('final', { pageWidth: -1 })).toThrow();
    expect(session.encodeState()).toEqual(before);
  } finally { session.destroy(); }
});

test('section insertion after the final paragraph appends one empty final paragraph', async () => {
  const session = await open();
  try {
    const last = session.paragraphs('body').at(-1)!;
    session.insertSectionBreakAfter(last.paraId, 'oddPage');
    expect(session.paragraphs('body').map(p => p.text)).toEqual(['First', 'Second', 'Third', '']);
    expect(session.listSections().map(s => [s.firstParagraph, s.lastParagraph])).toEqual([[1, 3], [4, 4]]);
    const reopened = await open((await saveYrsDocx(session)).bytes);
    try { expect(reopened.listSections()[1].authored.sectionStart).toBe('oddPage'); }
    finally { reopened.destroy(); }
  } finally { session.destroy(); }
});

test('undoing a section edit restores the same default export bytes', async () => {
  const session = await open();
  try {
    const baseline = (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
    session.setSectionProperties('final', { orientation: 'landscape' });
    expect(session.undo()).toBe(true);
    expect((await saveYrsDocx(session, { updateModifiedDate: false })).bytes).toEqual(baseline);
  } finally { session.destroy(); }
});

test('mixed text and section changes retain raw section XML', async () => {
  const source = await sectionPackage();
  const zip = await JSZip.loadAsync(source);
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:sectPr>', '<w:sectPr xmlns:u="urn:unknown"><u:opaque u:v=\'kept\'/>'));
  const session = await open(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const id = session.paragraphs('body')[0].paraId;
    session.insertSectionBreakAfter(id, 'continuous');
    session.replaceRange({ story: 'body', start: { paraId: id, offset: 0 }, end: { paraId: id, offset: 5 } }, 'Edited');
    const saved = await JSZip.loadAsync((await saveYrsDocx(session)).bytes);
    const result = await saved.file('word/document.xml')!.async('string');
    expect(result.match(/<u:opaque u:v='kept'\/>/g)).toHaveLength(2);
    expect(result).toContain('Edited');
  } finally { session.destroy(); }
});

test('non-body section targets return a typed refusal and keep the batch unchanged', async () => {
  const session = await open();
  try {
    const { paraId } = session.createStory('hf:section-test', 'Header');
    const before = session.encodeState();
    const result = session.applyEdits({ expectVersion: session.version(), steps: [{ op: 'insertSectionBreakAfter', target: { story: 'hf:section-test', paraId }, kind: 'nextPage' }] });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.failure.code).toBe('unsupported');
    expect(session.encodeState()).toEqual(before);
  } finally { session.destroy(); }
});

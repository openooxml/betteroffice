import { expect, test } from 'bun:test';
import { createCanvas, loadImage } from '@napi-rs/canvas';
import JSZip from 'jszip';
import { openPresentation } from '@betteroffice/pptx';
import { attachPptx, openPptx } from '../src/pptx';
import type { PptxAgentPresentation } from '../src/pptx';
import type { PptxAgentEdit } from '../src/pptx-schema';
import { pptxFixture } from './format-fixtures';
import { presentationFixture } from './pptx-fixture';

const apply = (deck: PptxAgentPresentation, edits: PptxAgentEdit[]) => deck.edit({ version: deck.overview().version, edits });

test('PPTX replacements compact 65 source runs and fall back within the run limit', async () => {
  const zip = await JSZip.loadAsync(await presentationFixture());
  const path = 'ppt/slides/slide1.xml';
  const xml = await zip.file(path)!.async('string');
  const runs = Array.from({ length: 65 }, (_, i) => `<a:r><a:rPr b="${i % 2}" sz="2400"/><a:t>x</a:t></a:r>`).join('');
  zip.file(path, xml.replace(/<a:p>[\s\S]*?<\/a:p>/u, `<a:p>${runs}</a:p>`));
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  for (const text of ['New text', 'x'.repeat(65), '']) {
    const deck = await openPptx(bytes);
    try {
      const target = first(deck);
      apply(deck, [{ op: 'replace_text', ...target, text: text === 'x'.repeat(65) ? 'y'.repeat(65) : text }]);
      const paragraph = deck.readSlide({ slide: target.slide }).shapes[0].stories[0].paragraphs[0];
      expect(paragraph.runs.length).toBeLessThanOrEqual(64);
      expect(paragraph.runs.every(run => run.text.length > 0)).toBe(true);
      if (text.length === 65) {
        expect(paragraph.runs).toHaveLength(1);
        expect(paragraph.runs[0].formatting).toMatchObject({ bold: false, fontSizePt: 24 });
      }
      const saved = await JSZip.loadAsync(await deck.export());
      const paragraphXml = (await saved.file(path)!.async('string')).match(/<a:p(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/a:p>)/u)![0];
      const savedText = [...paragraphXml.matchAll(/<a:t>([^<]*)<\/a:t>/gu)].map(match => match[1]).join('');
      expect(savedText).toBe(text.length === 65 ? 'y'.repeat(65) : text);
    } finally { deck.close(); }
  }
});
function first(deck: PptxAgentPresentation) {
  const slide = deck.outline().slides[0];
  return { slide: slide.id, shape: slide.shapes.find(shape => shape.text.length > 0)!.id };
}

test('PPTX outline and slide read expose layouts, tree, points, runs, placeholders and notes', async () => {
  const deck = await openPptx(await presentationFixture({ group: true }));
  try {
    const page = deck.outline({ shapeLimit: 1 });
    expect(page.slides[0]).toMatchObject({ index: 1, layoutName: 'Blank', shapeCount: 2, nextShapeOffset: 1 });
    expect(page.layouts[0]).toMatchObject({ id: 'ppt/slideLayouts/slideLayout1.xml', name: 'Blank' });
    const next = deck.outline({ shapeOffset: 1, shapeLimit: 1 });
    const shape = next.slides[0].shapes[0];
    expect(shape.parentId).toBe(page.slides[0].shapes[0].id);
    expect(shape.position.emu.x).toBe(100000);
    expect(shape.position.points.x).toBe(100000 / 12700);
    const read = deck.readSlide({ slide: page.slides[0].id, limit: 1 });
    expect(read.nextOffset).toBe(1);
    expect(read.notes.text).toBe('Present the quarterly result.');
    const child = deck.readSlide({ slide: page.slides[0].id, offset: read.nextOffset! }).shapes[0];
    expect(child.placeholder).toMatchObject({ placeholderType: 'title' });
    expect(child.stories[0].paragraphs[0].runs[0].formatting).toMatchObject({ bold: true, fontSizePt: 24, color: '#1565C0' });
    expect(child.stories[0].paragraphs).toHaveLength(2);
    expect(child.stories[0].id).toBeString();
  } finally { deck.close(); }
});

test('PPTX text, paragraph, bullet and notes edits preserve paragraph and run markup', async () => {
  const zip = await JSZip.loadAsync(await presentationFixture());
  const props = '<a:pPr algn="r" lvl="2" marL="3000" indent="-1000"><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buNone/></a:pPr>';
  const end = '<a:endParaRPr lang="de-DE" sz="1700"/>';
  for (const path of ['ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml']) {
    let xml = await zip.file(path)!.async('string');
    xml = xml.replaceAll('<a:p>', `<a:p>${props}`).replaceAll('</a:p>', `${end}</a:p>`);
    zip.file(path, xml);
  }
  const deck = await openPptx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const target = first(deck);
    apply(deck, [{ op: 'replace_text', ...target, text: 'First replacement\nSecond replacement' }, { op: 'find_replace', query: 'replacement', replacement: 'result' }, { op: 'set_notes', slide: target.slide, text: 'New notes' }]);
    expect(deck.readSlide({ slide: target.slide }).shapes[0].stories[0].paragraphs[0].runs[0].formatting).toMatchObject({ bold: true, fontSizePt: 24 });
    expect(deck.readSlide({ slide: target.slide }).shapes[0].stories[0].paragraphs[1].runs[0].formatting.bold).toBeNull();
    let saved = await JSZip.loadAsync(await deck.export());
    for (const path of ['ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml']) {
      const xml = await saved.file(path)!.async('string');
      for (const markup of ['algn="r"', 'lvl="2"', 'marL="3000"', 'indent="-1000"', '<a:spcBef><a:spcPts val="500"/></a:spcBef>', '<a:buNone/>', end]) expect(xml).toContain(markup);
    }
    apply(deck, [{ op: 'set_paragraphs', ...target, paragraphs: [{ bullet: true, runs: [{ text: 'Bullet one', bold: true }] }, { bullet: false, runs: [{ text: 'Plain two' }] }] }]);
    saved = await JSZip.loadAsync(await deck.export());
    const xml = await saved.file('ppt/slides/slide1.xml')!.async('string');
    expect(xml).toContain('<a:buChar char="•"/>');
    expect(xml).toContain('<a:buNone/>');
    expect(xml.match(/<a:endParaRPr/g)).toHaveLength(2);
    expect(xml.match(/indent="-1000"/g)).toHaveLength(2);
  } finally { deck.close(); }
});

test('PPTX inserts a title and one body shape with three real bullet paragraphs', async () => {
  const deck = await openPptx(await presentationFixture());
  try {
    const slide = apply(deck, [{ op: 'add_slide', index: 1, layout: deck.outline().layouts[0].id }]).results[0].slideId as string;
    expect(deck.readSlide({ slide }).shapeCount).toBe(0);
    const added = apply(deck, [
      { op: 'add_text_box', slide, name: 'Title', rect: { x: 36, y: 36, width: 648, height: 54 }, text: 'Pilot plan' },
      { op: 'add_text_box', slide, name: 'Body', rect: { x: 36, y: 126, width: 648, height: 270 }, text: '' },
    ]);
    apply(deck, [{ op: 'set_paragraphs', slide, shape: added.results[1].shapeId as string, paragraphs: ['Start with ten users', 'Collect feedback daily', 'Report results in two weeks'].map(text => ({ bullet: true, runs: [{ text }] })) }, { op: 'set_notes', slide, text: 'Keep the pilot to two weeks.' }]);
    const read = deck.readSlide({ slide });
    expect(read.shapeCount).toBe(2);
    expect(read.shapes[1].stories[0].paragraphs.map(p => JSON.parse(p.bullet!))).toEqual(Array(3).fill({ type: 'character', value: '•' }));
    const zip = await JSZip.loadAsync(await deck.export());
    expect((await zip.file('ppt/slides/slide2.xml')!.async('string')).match(/<a:buChar/g)).toHaveLength(3);
  } finally { deck.close(); }
});

test('PPTX replaces full text and sets formatted paragraphs with stable shape IDs', async () => {
  const deck = await openPptx(await pptxFixture());
  try {
    const target = first(deck);
    const replacement = apply(deck, [{ op: 'replace_text', ...target, text: 'First 😀\nSecond' }]);
    expect(replacement.results[0]).toMatchObject({ changed: true, shapeId: target.shape, paragraphs: 2 });
    let read = deck.readSlide({ slide: target.slide });
    expect(read.shapes[0].stories[0].text).toBe('First 😀\nSecond');
    expect(read.shapes[0].stories[0].paragraphs[0].runs[0].formatting.bold).toBe(true);
    apply(deck, [{ op: 'set_paragraphs', ...target, paragraphs: [{ alignment: 'ctr', runs: [{ text: 'Blue', bold: false, italic: true, fontSizePt: 18, fontFamily: 'Arial', color: '#0000FF' }, { text: ' bold', bold: true, underline: 'sng' }] }, { runs: [{ text: 'Next' }] }] }]);
    read = deck.readSlide({ slide: target.slide });
    expect(read.shapes[0].id).toBe(target.shape);
    expect(read.shapes[0].stories[0].paragraphs[0]).toMatchObject({ alignment: 'ctr', runs: [{ text: 'Blue', formatting: { bold: false, italic: true, fontSizePt: 18, color: '#0000FF' } }, { text: ' bold', formatting: { bold: true, underline: 'sng' } }] });
    const reopened = await openPptx(await deck.export());
    try { expect(reopened.readSlide({ slide: reopened.outline().slides[0].id }).shapes[0].stories[0].text).toBe('Blue bold\nNext'); }
    finally { reopened.close(); }
  } finally { deck.close(); }
});

test('PPTX find/replace edits the entire deck with counts and preserves run styles', async () => {
  const deck = await openPptx(await pptxFixture());
  try {
    const target = first(deck);
    apply(deck, [{ op: 'duplicate_slide', slide: target.slide }]);
    const receipt = apply(deck, [{ op: 'find_replace', query: 'risk', replacement: 'growth' }]);
    expect(receipt.results[0]).toMatchObject({ replacements: 4, changed: true });
    for (const slide of deck.outline().slides) expect(deck.readSlide({ slide: slide.id }).shapes[0].stories[0].text).toContain('growth growth');
    expect(apply(deck, [{ op: 'find_replace', query: 'GROWTH', replacement: 'x', caseSensitive: true }]).results[0].changed).toBe(false);
    expect(deck.readSlide({ slide: target.slide }).shapes[0].stories[0].paragraphs[0].runs[0].formatting.bold).toBe(true);
  } finally { deck.close(); }
});

test('PPTX adds, duplicates, moves and deletes slides with layout validation and persisted notes', async () => {
  const deck = await openPptx(await presentationFixture({ group: true }));
  try {
    const original = deck.outline().slides[0];
    const added = apply(deck, [{ op: 'add_slide', index: 1, layout: deck.outline().layouts[0].id }]).results[0].slideId as string;
    expect(deck.outline().slides[0].id).toBe(added);
    const clone = apply(deck, [{ op: 'duplicate_slide', slide: original.id, index: 1 }]).results[0].slideId as string;
    expect(deck.readSlide({ slide: clone }).shapes).toHaveLength(2);
    expect(deck.readSlide({ slide: clone }).notes.text).toBe('Present the quarterly result.');
    apply(deck, [{ op: 'move_slide', slide: clone, index: 3 }, { op: 'set_notes', slide: clone, text: 'Cloned notes 😀' }, { op: 'delete_slide', slide: original.id }]);
    expect(deck.outline().slides.map(s => s.id)).toEqual([added, clone]);
    const bytes = await deck.export();
    const zip = await JSZip.loadAsync(bytes);
    const cloneXml = await zip.file('ppt/slides/slide3.xml')!.async('string');
    expect(cloneXml).toContain('<p:grpSp>');
    expect(cloneXml).toContain('<p:transition spd="slow">');
    expect(await zip.file('customXml/item1.xml')!.async('string')).toBe('<custom>preserve me</custom>');
    const reopened = await openPptx(bytes);
    try {
      expect(reopened.overview().slides).toBe(2);
      expect(reopened.readSlide({ slide: reopened.outline().slides[1].id }).notes.text).toBe('Cloned notes 😀');
    } finally { reopened.close(); }
    expect(() => apply(deck, [{ op: 'add_slide', layout: 'missing' }])).toThrow('Copy a layout ID');
  } finally { deck.close(); }
});

test('PPTX adds text boxes, preset shapes and images in points and exports them', async () => {
  const deck = await openPptx(await pptxFixture());
  const image = createCanvas(20, 20);
  image.getContext('2d').fillStyle = '#ff0000';
  image.getContext('2d').fillRect(0, 0, 20, 20);
  try {
    const slide = first(deck).slide;
    const rect = { x: 20, y: 140, width: 80, height: 80 };
    const receipt = apply(deck, [
      { op: 'add_text_box', slide, name: 'Caption', rect, text: 'Added caption', style: { bold: true, fontSizePt: 16 } },
      { op: 'add_shape', slide, name: 'Circle', rect, geometry: 'ellipse', fill: '#00FF00' },
      { op: 'add_image', slide, name: 'Logo', rect: { ...rect, x: 140 }, contentType: 'image/png', base64: image.toBuffer('image/png').toString('base64') },
    ]);
    expect(receipt.results).toHaveLength(3);
    const shapes = deck.readSlide({ slide }).shapes;
    expect(shapes[1]).toMatchObject({ name: 'Caption', position: { emu: { x: 254000 }, points: rect } });
    expect(shapes[2]).toMatchObject({ geometry: 'ellipse', fill: '#00FF00' });
    expect(shapes[3].type).toBe('picture');
    const reopened = await openPptx(await deck.export());
    try { expect(reopened.readSlide({ slide: reopened.outline().slides[0].id }).shapes).toHaveLength(4); }
    finally { reopened.close(); }
    const png = await deck.preview(slide);
    expect([...png.png.slice(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    const pixels = createCanvas(png.width, png.height).getContext('2d');
    pixels.drawImage(await loadImage(Buffer.from(png.png)), 0, 0);
    const data = pixels.getImageData(0, 0, png.width, png.height).data;
    let red = 0;
    for (let i = 0; i < data.length; i += 4) if (data[i] > 240 && data[i + 1] < 10 && data[i + 2] < 10) red++;
    expect(red).toBeGreaterThan(1000);
  } finally { deck.close(); }
});

test('PPTX rejects a failed batch without applying earlier items and returns exact repair IDs', async () => {
  const deck = await openPptx(await pptxFixture());
  try {
    const target = first(deck);
    const before = await deck.export();
    const version = deck.overview().version;
    let failure: any;
    try { apply(deck, [{ op: 'set_notes', slide: target.slide, text: 'Should roll back' }, { op: 'replace_text', slide: target.slide, shape: target.shape + 'x', text: 'invalid' }]); }
    catch (error) { failure = error; }
    expect(failure.code).toBe('UNKNOWN_SHAPE');
    expect(failure.details.failedIndex).toBe(1);
    expect(failure.details.nearestShapeIds[0].id).toBe(target.shape);
    expect(failure.details.results.map((item: any) => item.status)).toEqual(['rolled_back', 'failed']);
    expect(deck.overview().version).toBe(version);
    expect(await deck.export()).toEqual(before);
    expect(() => apply(deck, [{ op: 'delete_slide', slide: target.slide }])).toThrow('last slide');
    expect(() => apply(deck, [{ op: 'set_notes', slide: 'missing', text: '' }])).toThrow('Copy a slide ID');
    apply(deck, [{ op: 'set_notes', slide: target.slide, text: 'Changed' }]);
    expect(() => deck.edit({ version, edits: [{ op: 'set_notes', slide: target.slide, text: 'Stale' }] })).toThrow('Deck changed');
  } finally { deck.close(); }
});

test('PPTX attaches to a live handle, commits one update, sees host changes and leaves ownership intact', async () => {
  const source = await pptxFixture();
  const seed = await openPptx(source);
  seed.close();
  const handle = openPresentation(source);
  handle.setSlideNotes(handle.snapshot().slides[0].id, 'Before attachment');
  const deck = attachPptx(handle, { name: 'Live deck' });
  let updates = 0;
  const off = handle.onUpdate(() => { updates++; });
  try {
    const target = first(deck);
    apply(deck, [{ op: 'replace_text', ...target, text: 'Live content' }, { op: 'set_notes', slide: target.slide, text: 'Live notes' }]);
    expect(updates).toBe(1);
    expect(handle.snapshot().slides[0].notes).toBe('Live notes');
    handle.setSlideNotes(target.slide, 'Host edit');
    expect(deck.readSlide({ slide: target.slide }).notes.text).toBe('Host edit');
    expect((await deck.verify()).reopened).toBe(true);
    deck.close();
    expect(handle.snapshot().slides).toHaveLength(1);
    expect(() => deck.outline()).toThrow('closed');
  } finally { off(); deck.close(); handle.dispose(); }
});

test('PPTX refuses unsafe Unicode, images, text fields and oversized find/replace batches', async () => {
  const deck = await openPptx(await pptxFixture('risk '.repeat(200)));
  try {
    const target = first(deck);
    expect(() => apply(deck, [{ op: 'replace_text', ...target, text: '\ud800' }])).toThrow('complete Unicode');
    expect(() => apply(deck, [{ op: 'find_replace', query: 'risk', replacement: 'x' }])).toThrow('128 occurrences');
    expect(deck.readSlide({ slide: target.slide }).shapes[0].stories[0].text).toContain('risk');
    expect(() => apply(deck, [{ op: 'add_image', slide: target.slide, name: 'Bad', rect: { x: 0, y: 0, width: 1, height: 1 }, contentType: 'image/png', base64: 'a===' }])).toThrow('canonical base64');
    expect(() => apply(deck, [{ op: 'set_paragraphs', ...target, paragraphs: [{ runs: [{ text: 'bad\nrun' }] }] }])).toThrow('one paragraph');
  } finally { deck.close(); }
  const fields = await openPptx(await presentationFixture({ fields: true }));
  try { expect(() => apply(fields, [{ op: 'replace_text', ...first(fields), text: 'Would remove field' }])).toThrow('text fields'); }
  finally { fields.close(); }
});

test('PPTX paginates long run/text/notes reads without silently dropping formatting', async () => {
  const deck = await openPptx(await presentationFixture({ runs: 100 }));
  try {
    const target = first(deck);
    apply(deck, [{ op: 'set_notes', slide: target.slide, text: 'x'.repeat(6000) }]);
    const read = deck.readSlide({ slide: target.slide, paragraphOffset: 1, textLength: 4000 });
    const paragraph = read.shapes[0].stories[0].paragraphs[0];
    expect(paragraph.nextRunOffset).toBe(8);
    const next = deck.readSlide({ slide: target.slide, paragraphOffset: 1, runOffset: paragraph.nextRunOffset! });
    expect(next.shapes[0].stories[0].paragraphs[0].runs[0].text).not.toBe(paragraph.runs[0].text);
    expect(read.notes.nextStart).toBe(4000);
    expect(deck.readSlide({ slide: target.slide, notesStart: 4000 }).notes.text.length).toBe(2000);
    expect(JSON.stringify(read).length).toBeLessThan(30000);
  } finally { deck.close(); }
});

test('PPTX previews pending proposals without mutating the presentation', async () => {
  const deck = await openPptx(await pptxFixture());
  try {
    const hit = deck.grep({ query: 'Risk' }).matches[0];
    const pending = deck.propose({ author: 'test', edits: [{ match: hit.match, newText: 'Opportunity' }] });
    const before = await deck.render(1);
    const after = await deck.render(1, pending.id);
    expect(after.png).not.toEqual(before.png);
    expect(deck.grep({ query: 'Opportunity' }).matches).toHaveLength(0);
    await expect(deck.render(2)).rejects.toThrow('Choose a slide index');
  } finally { deck.close(); }
});

test('PPTX shape-heavy outlines remain bounded and paginate every shape', async () => {
  const source = await pptxFixture();
  const initialized = await openPptx(source);
  initialized.close();
  const handle = openPresentation(source);
  const slide = handle.snapshot().slides[0].id;
  for (let i = 0; i < 80; i++) handle.addShape(slide, { name: `Shape ${i} ${'x'.repeat(180)}`, geometry: 'rect', rect: { x: i * 100, y: 0, width: 10000, height: 10000 } });
  const deck = attachPptx(handle);
  try {
    let offset = 0;
    const ids = new Set<string>();
    while (true) {
      const page = deck.outline({ limit: 100, shapeLimit: 100, shapeOffset: offset });
      expect(JSON.stringify(page).length).toBeLessThan(17000);
      for (const shape of page.slides[0].shapes) ids.add(shape.id);
      const next = page.slides[0].nextShapeOffset;
      if (next === null) break;
      expect(next).toBeGreaterThan(offset);
      offset = next;
    }
    expect(ids.size).toBe(81);
  } finally { deck.close(); handle.dispose(); }
});

test('PPTX duplication preserves added images and edits the copy independently', async () => {
  const deck = await openPptx(await presentationFixture({ group: true }));
  try {
    const slide = deck.outline().slides[0];
    const image = createCanvas(2, 2).toBuffer('image/png').toString('base64');
    apply(deck, [{ op: 'add_image', slide: slide.id, name: 'Picture', rect: { x: 5, y: 5, width: 10, height: 10 }, contentType: 'image/png', base64: image }]);
    const clone = apply(deck, [{ op: 'duplicate_slide', slide: slide.id }]).results[0].slideId as string;
    const shape = deck.readSlide({ slide: clone }).shapes.find(shape => shape.name === 'Revenue')!.id;
    apply(deck, [{ op: 'replace_text', slide: clone, shape, text: 'Independent clone' }]);
    expect(deck.outline().slides[0].title).toContain('Revenue');
    expect(deck.outline().slides[1].title).toBe('Independent clone');
    const reopened = await openPptx(await deck.export());
    try {
      for (const slide of reopened.outline().slides) expect(reopened.readSlide({ slide: slide.id }).shapes.some(shape => shape.type === 'picture')).toBe(true);
      expect(reopened.outline().slides[1].title).toBe('Independent clone');
    } finally { reopened.close(); }
  } finally { deck.close(); }
});

import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createFontProvider } from '@betteroffice/fonts';
import { openPresentation } from '@betteroffice/pptx';
import { attachPptx, openPptx } from '../src/pptx';
import { registerPptxTools } from '../src/pptx-tools';
import type { FileWorkspace } from '../src/workspace';
import { presentationFixture } from './pptx-fixture';

const slidePath = 'ppt/slides/slide1.xml';
const untouched = '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Untouched"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr i="1" lang="de-DE"/><a:t>Keep exactly</a:t></a:r></a:p></p:txBody></p:sp>';

async function fixture(trailingBreak = false, autofit?: 'layout' | 'slide') {
  const zip = await JSZip.loadAsync(await presentationFixture());
  const paragraphs = `<a:p><a:r><a:rPr b="0"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr b="1"/><a:t>Second</a:t></a:r>${trailingBreak ? '<a:br/>' : ''}</a:p><a:p><a:r><a:t>Tail</a:t></a:r></a:p>`;
  let slide = (await zip.file(slidePath)!.async('string')).replace(/<a:p>[\s\S]*<\/a:p>/u, paragraphs).replace('</p:spTree>', `${untouched}</p:spTree>`);
  if (autofit === 'slide') slide = slide.replace('<a:bodyPr/>', '<a:bodyPr><a:normAutofit/></a:bodyPr>');
  if (autofit === 'layout') {
    const path = 'ppt/slideLayouts/slideLayout1.xml';
    const placeholder = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p/></p:txBody></p:sp>';
    zip.file(path, (await zip.file(path)!.async('string')).replace('</p:spTree>', `${placeholder}</p:spTree>`));
  }
  zip.file(slidePath, slide);
  return zip.generateAsync({ type: 'uint8array' });
}

function target(deck: Awaited<ReturnType<typeof openPptx>>) {
  const slide = deck.outline().slides[0];
  return { slide: slide.id, shape: slide.shapes[0].id };
}

async function verifyUntouched(before: Uint8Array, after: Uint8Array) {
  const source = await JSZip.loadAsync(before);
  const saved = await JSZip.loadAsync(after);
  expect(await saved.file(slidePath)!.async('string')).toContain(untouched);
  for (const path of Object.keys(source.files)) {
    if (path !== slidePath && !source.files[path].dir) expect(await saved.file(path)!.async('uint8array')).toEqual(await source.file(path)!.async('uint8array'));
  }
}

test('PPTX slide read distinguishes soft breaks and paragraph breaks and round trips unchanged', async () => {
  const bytes = await fixture();
  const deck = await openPptx(bytes);
  try {
    const address = target(deck);
    const read = deck.readSlide(address);
    const story = read.shapes[0].stories[0];
    expect(story.text).toBe('First\vSecond\nTail');
    expect(read.shapes[0].text.text).toBe('First\vSecond\nTail');
    expect(read.title).toBe('First\vSecond');
    expect(story.paragraphCount).toBe(2);
    expect(story.paragraphs[0].runs.map(run => run.text).join('')).toBe('First\vSecond');
    expect(story.paragraphs[1].start).toBe(13);
    const result = deck.edit({ version: read.version, edits: [{ op: 'replace_text', ...address, text: story.text }] });
    expect(result.results[0].changed).toBe(false);
    const saved = await JSZip.loadAsync(await deck.export());
    const source = await JSZip.loadAsync(bytes);
    expect(await saved.file(slidePath)!.async('string')).toBe(await source.file(slidePath)!.async('string'));
  } finally { deck.close(); }
});

test('PPTX whole shape replacement keeps matching styled runs around soft breaks', async () => {
  const bytes = await fixture(true);
  const deck = await openPptx(bytes);
  let saved: Uint8Array;
  try {
    const address = target(deck);
    deck.edit({ version: deck.overview().version, edits: [{ op: 'replace_text', ...address, text: 'First\vSecond\nEdited tail' }] });
    const read = deck.readSlide(address).shapes[0].stories[0];
    expect(read.text).toBe('First\vSecond\nEdited tail');
    expect(read.paragraphCount).toBe(2);
    expect(read.paragraphs[0].runs.find(run => run.text === 'Second')?.formatting.bold).toBe(true);
    saved = await deck.export();
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file(slidePath)!.async('string');
    expect(xml.match(/<a:br\b/gu)).toHaveLength(1);
    expect(xml.match(/<a:p>/gu)).toHaveLength(3);
    await verifyUntouched(bytes, saved);
  } finally { deck.close(); }
  const reopened = await openPptx(saved!);
  try {
    const story = reopened.readSlide(target(reopened)).shapes[0].stories[0];
    expect(story.text).toBe('First\vSecond\nEdited tail');
    expect(story.paragraphs[0].runs.find(run => run.text === 'Second')?.formatting.bold).toBe(true);
  } finally { reopened.close(); }
});

test('PPTX replacement with unrelated text uses the first run format', async () => {
  const deck = await openPptx(await fixture());
  try {
    const address = target(deck);
    deck.edit({ version: deck.overview().version, edits: [{ op: 'replace_text', ...address, text: 'Other content' }] });
    const runs = deck.readSlide(address).shapes[0].stories[0].paragraphs[0].runs;
    expect(runs).toHaveLength(1);
    expect(runs[0]).toMatchObject({ text: 'Other content', formatting: { bold: false } });
  } finally { deck.close(); }
});

for (const autofit of ['layout', 'slide'] as const) test(`PPTX text edits preserve ${autofit} autofit ownership`, async () => {
  const bytes = await fixture(false, autofit);
  const deck = await openPptx(bytes);
  let saved: Uint8Array;
  try {
    await deck.preview(target(deck).slide);
    deck.edit({ version: deck.overview().version, edits: [{ op: 'replace_text', ...target(deck), text: 'Updated title' }] });
    saved = await deck.export();
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file(slidePath)!.async('string');
    expect(xml.includes('<a:normAutofit')).toBe(autofit === 'slide');
    if (autofit === 'layout') expect(xml).toContain('<a:bodyPr/>');
    await verifyUntouched(bytes, saved);
  } finally { deck.close(); }
  const reopened = await openPptx(saved!);
  try { expect(reopened.readSlide(target(reopened)).shapes[0].stories[0].text).toBe('Updated title'); }
  finally { reopened.close(); }
});

test('PPTX replacement treats newline as a paragraph boundary even when flattened text matches', async () => {
  const deck = await openPptx(await fixture());
  try {
    const address = target(deck);
    const result = deck.edit({ version: deck.overview().version, edits: [{ op: 'replace_text', ...address, text: 'First\nSecond\nTail' }] });
    expect(result.results[0].changed).toBe(true);
    const read = deck.readSlide(address).shapes[0].stories[0];
    expect(read.paragraphCount).toBe(3);
    expect(read.paragraphs.map(paragraph => paragraph.runs.map(run => run.text).join(''))).toEqual(['First', 'Second', 'Tail']);
  } finally { deck.close(); }
});

test('PPTX tool descriptions explain soft breaks and replacement formatting', () => {
  const descriptions = new Map<string, string>();
  registerPptxTools((name, description) => { descriptions.set(name, description); }, {} as FileWorkspace);
  expect(descriptions.get('pptx_read_slide')).toContain('\\u000b');
  const edit = descriptions.get('pptx_edit')!;
  expect(edit).toContain('soft line break');
  expect(edit).toContain("first run's format");
  expect(edit).toContain('find_replace or office_propose');
});

test('PPTX matching paragraphs above the run limit refuse without changing the deck', async () => {
  const zip = await JSZip.loadAsync(await fixture());
  const runs = Array.from({ length: 65 }, (_, index) => `<a:r><a:rPr b="${index % 2}"/><a:t>x</a:t></a:r>`).join('');
  zip.file(slidePath, (await zip.file(slidePath)!.async('string')).replace(/<a:p>[\s\S]*?<\/a:p>/u, `<a:p>${runs}</a:p>`));
  const deck = await openPptx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const address = target(deck);
    const version = deck.overview().version;
    expect(() => deck.edit({ version, edits: [{ op: 'replace_text', ...address, text: `${'x'.repeat(65)}\nChanged tail` }] })).toThrow('more than 64 formatted runs');
    expect(deck.overview().version).toBe(version);
    expect(deck.readSlide(address).shapes[0].stories[0].text).toBe(`${'x'.repeat(65)}\nTail`);
  } finally { deck.close(); }
});

test('PPTX inherited autofit writes changed fit and reopens with the same layout', async () => {
  const bytes = await fixture(false, 'layout');
  const seed = await openPptx(bytes);
  seed.close();
  const provider = createFontProvider();
  const fonts = await Promise.all([
    { family: 'Arial', bold: false, italic: false },
    { family: 'Arial', bold: true, italic: false },
    { family: 'Arial', bold: false, italic: true },
  ].map(async face => {
    const load = provider.resolve(face.family, face.bold, face.italic) ?? provider.resolveLastResort(face.family, face.bold, face.italic, 'powerpoint');
    return { ...face, bytes: new Uint8Array(await load()) };
  }));
  const handle = openPresentation(bytes, { fonts });
  const deck = attachPptx(handle);
  let saved: Uint8Array;
  let live: ReturnType<typeof handle.layoutSlide>;
  try {
    const text = 'Words that must shrink inside the title box. '.repeat(20);
    deck.edit({ version: deck.overview().version, edits: [{ op: 'replace_text', ...target(deck), text }] });
    live = handle.layoutSlide(0);
    saved = await deck.export();
    const zip = await JSZip.loadAsync(saved);
    const xml = await zip.file(slidePath)!.async('string');
    const scale = Number(xml.match(/fontScale="([0-9]+)"/u)?.[1]);
    expect(scale).toBeGreaterThan(0);
    expect(scale).toBeLessThan(100000);
    expect(xml).toContain('lnSpcReduction=');
    await verifyUntouched(bytes, saved);
  } finally { deck.close(); handle.dispose(); }
  const reopenedHandle = openPresentation(saved!, { fonts });
  const reopened = attachPptx(reopenedHandle);
  try {
    expect(reopenedHandle.layoutSlide(0)).toEqual(live!);
  } finally { reopened.close(); reopenedHandle.dispose(); }
});

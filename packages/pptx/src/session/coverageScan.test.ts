import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { activeSlideCoverageScripts, slideCoverageScripts } from './coverageScan';

const run = (text: string, lang?: string) =>
  `<a:r><a:rPr${lang ? ` lang="${lang}"` : ''}/><a:t>${text}</a:t></a:r>`;
const slide = (runs: string) =>
  `<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p>${runs}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;

async function deck(slides: string[], order: number[]): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('ppt/presentation.xml', `<p:presentation><p:sldIdLst>${order
    .map((number, index) => `<p:sldId id="${256 + index}" r:id="rId${number}"/>`).join('')}</p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships>${slides
    .map((_, index) => `<Relationship Id="rId${index + 1}" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
  slides.forEach((xml, index) => zip.file(`ppt/slides/slide${index + 1}.xml`, xml));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

test('slide text with certain CJK needs selects its coverage scripts', () => {
  expect(slideCoverageScripts(slide(run('かな')))).toEqual(['cjk-jp']);
  expect(slideCoverageScripts(slide(run('推計', 'ja-JP')))).toEqual(['cjk-jp']);
  expect(slideCoverageScripts(slide(run('한글')))).toEqual(['cjk-kr']);
  expect(slideCoverageScripts(slide(run('繁體', 'zh-TW') + run('香港', 'zh-HK')))).toEqual(['cjk-tc']);
  expect(slideCoverageScripts(slide(run('汉字', 'zh-CN')))).toEqual(['cjk-sc']);
  expect(slideCoverageScripts(slide(run('漢字')))).toEqual([]);
  expect(slideCoverageScripts(slide(run('Latin', 'ja-JP')))).toEqual([]);
  expect(slideCoverageScripts(slide(run('推計', 'ja-JP') + run('汉字')))).toEqual(['cjk-jp']);
});

test('only the slide that opens first is scanned', async () => {
  const bytes = await deck([slide(run('English')), slide(run('かな'))], [2, 1]);
  expect(await activeSlideCoverageScripts(bytes)).toEqual(['cjk-jp']);
  expect(await activeSlideCoverageScripts(bytes, 1)).toEqual([]);
  expect(await activeSlideCoverageScripts(bytes.buffer as ArrayBuffer, 9)).toEqual([]);
});

test('unreadable packages need no coverage', async () => {
  expect(await activeSlideCoverageScripts(new Uint8Array([1, 2, 3]))).toEqual([]);
  expect(await activeSlideCoverageScripts(new Uint8Array(0))).toEqual([]);
});

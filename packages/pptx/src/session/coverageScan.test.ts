import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { activeSlideCoverageScripts, slideCoverageScripts } from './coverageScan';

const run = (text: string, lang?: string) =>
  `<a:r><a:rPr${lang ? ` lang="${lang}"` : ''}/><a:t>${text}</a:t></a:r>`;
const slide = (runs: string) =>
  `<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p>${runs}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;

async function deck(slides: string[], order: (number | string)[], ids?: string[]): Promise<Uint8Array> {
  const zip = new JSZip();
  const rid = (number: number | string) => (typeof number === 'number' ? `rId${number}` : number);
  zip.file('ppt/presentation.xml', `<p:presentation><p:sldIdLst>${order
    .map((number, index) => `<p:sldId id="${256 + index}" r:id="${rid(number)}"/>`).join('')}</p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships>${slides
    .map((_, index) => `<Relationship Id="${ids?.[index] ?? `rId${index + 1}`}" Target="slides/slide${index + 1}.xml"/>`).join('')}</Relationships>`);
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

function declareSize(bytes: Uint8Array, name: string, size: number): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const encoded = new TextEncoder().encode(name);
  for (let at = 0; at + 46 <= bytes.byteLength; at++) {
    if (view.getUint32(at, true) !== 0x02014b50 || view.getUint16(at + 28, true) !== encoded.length) continue;
    if (!encoded.every((byte, index) => bytes[at + 46 + index] === byte)) continue;
    view.setUint32(at + 24, size, true);
    view.setUint32(view.getUint32(at + 42, true) + 22, size, true);
  }
}

test('a crafted relationship id is compared as text, not run as a pattern', async () => {
  const bytes = await deck([slide(run('かな'))], ['(a+)+$'], [`${'a'.repeat(28)}!`]);
  const started = performance.now();
  expect(await activeSlideCoverageScripts(bytes)).toEqual([]);
  expect(performance.now() - started).toBeLessThan(250);
  const literal = await deck([slide(run('かな'))], ['(a+)+$'], ['(a+)+$']);
  expect(await activeSlideCoverageScripts(literal)).toEqual(['cjk-jp']);
}, 60_000);

test('unterminated tags are scanned in linear time', () => {
  const started = performance.now();
  expect(slideCoverageScripts('<a:rPr lang="ja-JP" '.repeat(10_000))).toEqual([]);
  expect(slideCoverageScripts(`${'<a:t>'.repeat(20_000)}かな`)).toEqual(['cjk-jp']);
  expect(performance.now() - started).toBeLessThan(250);
}, 60_000);

test('a part that inflates past its declared size is abandoned at the scan limit', async () => {
  const zip = new JSZip();
  zip.file('ppt/presentation.xml', `<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>${' '.repeat(24 * 1024 * 1024)}</p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slide(run('かな')));
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  declareSize(bytes, 'ppt/presentation.xml', 100);
  const native = globalThis.DecompressionStream;
  let inflated = 0;
  globalThis.DecompressionStream = class {
    readonly writable: WritableStream<BufferSource>;
    readonly readable: ReadableStream<Uint8Array>;
    constructor(format: CompressionFormat) {
      const inner = new native(format);
      this.writable = inner.writable;
      this.readable = inner.readable.pipeThrough(new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, controller) { inflated += chunk.byteLength; controller.enqueue(chunk); },
      }));
    }
  } as unknown as typeof DecompressionStream;
  try {
    expect(await activeSlideCoverageScripts(bytes)).toEqual([]);
  } finally {
    globalThis.DecompressionStream = native;
  }
  expect(inflated).toBeLessThan(6 * 1024 * 1024);
}, 60_000);

test('stored parts must match their declared size', async () => {
  const zip = new JSZip();
  zip.file('ppt/presentation.xml', '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', slide(run('かな')));
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
  expect(await activeSlideCoverageScripts(bytes)).toEqual(['cjk-jp']);
  declareSize(bytes, 'ppt/slides/slide1.xml', 10);
  expect(await activeSlideCoverageScripts(bytes)).toEqual([]);
});

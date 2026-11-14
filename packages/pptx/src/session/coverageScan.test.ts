import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { activeSlideCoverageScripts, activeSlideMetafileFlags, scanOpeningSlide, slideCoverageScripts } from './coverageScan';

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
  let largestInput = 0;
  globalThis.DecompressionStream = class {
    readonly writable: WritableStream<BufferSource>;
    readonly readable: ReadableStream<Uint8Array>;
    constructor(format: CompressionFormat) {
      const inner = new native(format);
      const input = new TransformStream<Uint8Array, Uint8Array>({
        transform(chunk, controller) { largestInput = Math.max(largestInput, chunk.byteLength); controller.enqueue(chunk); },
      });
      input.readable.pipeTo(inner.writable as unknown as WritableStream<Uint8Array>).catch(() => {});
      this.writable = input.writable as unknown as WritableStream<BufferSource>;
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
  expect(largestInput).toBeGreaterThan(0);
  expect(largestInput).toBeLessThanOrEqual(1024);
}, 60_000);

test('a long language tag is read once per run, not once per character', () => {
  const language = `ja-x-${'a-'.repeat(10_000)}b`;
  const started = performance.now();
  expect(slideCoverageScripts(slide(run('漢'.repeat(10_000), language)))).toEqual(['cjk-jp']);
  expect(performance.now() - started).toBeLessThan(250);
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

for (const extension of ['emf', 'WMF']) test(`opening metafile relationships are found in presentation order (${extension})`, async () => {
  const zip = await JSZip.loadAsync(await deck([slide(run('English')), slide(run('かな'))], [2, 1]));
  zip.file('ppt/slides/_rels/slide2.xml.rels', `<Relationships><Relationship Id="picture" Target="../media/picture.${extension}"/></Relationships>`);
  zip.file(`ppt/media/picture.${extension}`, new Uint8Array([1]));
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  expect(await scanOpeningSlide(bytes)).toEqual({ scripts: ['cjk-jp'], openingSlideMetafile: true, packageMetafile: true });
  expect(await activeSlideMetafileFlags(bytes, 1)).toEqual({ openingSlideMetafile: false, packageMetafile: true });
  expect(await activeSlideMetafileFlags(bytes, 99)).toEqual({ openingSlideMetafile: false, packageMetafile: true });
});

test('metafile flags distinguish relationships from package media', async () => {
  const zip = await JSZip.loadAsync(await deck([slide(run('English'))], [1]));
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="picture" Target="../media/picture.png"/></Relationships>');
  zip.file('ppt/media/picture.png', new Uint8Array([1]));
  zip.file('other/picture.emf', new Uint8Array([1]));
  expect(await activeSlideMetafileFlags(await zip.generateAsync({ type: 'uint8array' })))
    .toEqual({ openingSlideMetafile: false, packageMetafile: false });
  zip.file('ppt/media/unused.EMF', new Uint8Array([1]));
  expect(await activeSlideMetafileFlags(await zip.generateAsync({ type: 'uint8array' })))
    .toEqual({ openingSlideMetafile: false, packageMetafile: true });
  zip.remove('ppt/media/unused.EMF');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="picture" Target="../media/missing.wmf"/></Relationships>');
  expect(await activeSlideMetafileFlags(await zip.generateAsync({ type: 'uint8array' })))
    .toEqual({ openingSlideMetafile: true, packageMetafile: false });
});

test('malformed packages have no metafile flags', async () => {
  for (const bytes of [new Uint8Array(), new Uint8Array([1, 2, 3]), new Uint8Array(64)]) {
    expect(await activeSlideMetafileFlags(bytes)).toEqual({ openingSlideMetafile: false, packageMetafile: false });
  }
  const bytes = await deck([slide(run('English'))], [1]);
  expect(await activeSlideMetafileFlags(bytes.subarray(0, bytes.length - 20)))
    .toEqual({ openingSlideMetafile: false, packageMetafile: false });
});

test('metafile relationship scans keep the part size limit', async () => {
  const zip = await JSZip.loadAsync(await deck([slide(run('English'))], [1]));
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="picture" Target="../media/picture.emf"/></Relationships>');
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  declareSize(bytes, 'ppt/slides/_rels/slide1.xml.rels', 4 * 1024 * 1024 + 1);
  expect(await activeSlideMetafileFlags(bytes)).toEqual({ openingSlideMetafile: false, packageMetafile: false });
});

for (const holder of ['layout', 'master'] as const) test(`a metafile the opening slide shows through its ${holder} is found`, async () => {
  const layoutType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
  const masterType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
  const zip = await JSZip.loadAsync(await deck([slide(run('English')), slide(run('English'))], [1, 2]));
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships><Relationship Id="rId1" Type="${layoutType}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<Relationships><Relationship Id="rId1" Type="${masterType}" Target="../slideMasters/slideMaster1.xml"/>${holder === 'layout' ? '<Relationship Id="rId2" Target="../media/logo.emf"/>' : ''}</Relationships>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<Relationships>${holder === 'master' ? '<Relationship Id="rId9" Target="../media/logo.emf"/>' : ''}</Relationships>`);
  zip.file('ppt/media/logo.emf', new Uint8Array([1]));
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  expect(await activeSlideMetafileFlags(bytes)).toEqual({ openingSlideMetafile: true, packageMetafile: true });
  expect(await activeSlideMetafileFlags(bytes, 1)).toEqual({ openingSlideMetafile: false, packageMetafile: true });
});

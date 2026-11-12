import { afterAll, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import JSZip from 'jszip';
import {
  REGISTRY_EXPORTS_SINCE, compareSemver, describeAddons, fontPathFor, guardRegistrations, layOutFonts, loadProduct, missingAddons,
  serializeChains, writeDocumentChains, type FontPackages, type Product,
} from './native-chains';

type Faces = Record<string, string[]>;

const repo = resolve(import.meta.dir, '../..');
const productRoot = resolve(process.env.NATIVE_CHAINS_PRODUCT ?? join(repo, 'packages/docx'));
const built = existsSync(join(productRoot, 'dist'));
const binary = process.env.NATIVE_BENCH_BINARY;
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

const latin = (family: string) => `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>`;
const RUNS = [
  ['Consolas monospace 0123', latin('Consolas'), ''],
  ['Rockwell slab serif', latin('Rockwell'), ''],
  ['Microsoft Sans Serif text', latin('Microsoft Sans Serif'), ''],
  ['Verdana humanist sans', latin('Verdana'), ''],
  ['Arial Black heavy face', latin('Arial Black'), ''],
  ['日本語の明朝体の本文です', '<w:rFonts w:ascii="Noto Serif JP" w:hAnsi="Noto Serif JP" w:eastAsia="Noto Serif JP"/>', '<w:lang w:eastAsia="ja-JP"/>'],
  ['مرحبا بالعالم', latin('Times New Roman'), '<w:rtl/><w:lang w:bidi="ar-SA"/>'],
  ['Calibri body text', latin('Calibri'), ''],
];

async function fontDocx(): Promise<Uint8Array> {
  const body = RUNS.map(([text, fonts, extra]) =>
    `<w:p><w:r><w:rPr>${fonts}${extra}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function replayedNames(product: Product, fonts: Array<Uint8Array | { substituteOf: number; family: string }>): Promise<string[]> {
  const session = await product.yrs.createYrsSession();
  try {
    const names: string[] = [];
    for (const font of fonts) {
      const isBytes = ArrayBuffer.isView(font);
      const id: number = isBytes ? session.registerFont(font) : session.registerSubstituteFont(font.substituteOf, font.family);
      if (id === names.length) names.push(isBytes ? hash(font) : `${names[font.substituteOf]}~${font.family}`);
      else expect({ id, base: isBytes ? undefined : font.substituteOf }).toEqual({ id, base: id });
    }
    return names;
  } finally {
    session.destroy();
  }
}

async function productFaces(product: Product, bytes: Uint8Array): Promise<{ faces: Faces; measurement: Record<string, unknown> }> {
  const laid = await layOutFonts(product, bytes);
  try {
    const { fontChains: _, ...measurement } = laid.measurement;
    product.editor.computeLayout({ document: laid.document, pageGap: 0, session: laid.session, renderEnv: laid.renderEnv, measurement: laid.measurement });
    const fonts = laid.session.residentWorkerSnapshot()?.fonts as Array<Uint8Array | { substituteOf: number; family: string }> | undefined;
    if (!fonts) throw new Error('the session handed no resident fonts to wasm');
    console.log(`RAW product: ${JSON.stringify(fonts.map((font) => (ArrayBuffer.isView(font) ? hash(font) : font)))}`);
    const names = await replayedNames(product, fonts);
    return { faces: Object.fromEntries(Object.entries(laid.measurement.fontChains).map(([key, chain]) => [key, chain.map((id) => names[id])])), measurement };
  } finally {
    laid.session.destroy();
  }
}

async function chainsFaces(path: string): Promise<Faces> {
  const chains = JSON.parse(await readFile(path, 'utf8'));
  console.log(`RAW recorded: ${JSON.stringify(chains.faces.map((face: any) => ('substituteOf' in face ? face : face.sha256)))}`);
  const names: string[] = [];
  for (const face of chains.faces) {
    if ('substituteOf' in face) {
      expect(face.substituteOf).toBeLessThan(names.length);
      names.push(`${names[face.substituteOf]}~${face.family}`);
      continue;
    }
    expect(face.file).toMatch(/^[^/\\]+$/);
    const bytes = new Uint8Array(await readFile(join(dirname(path), face.file)));
    expect(hash(bytes)).toBe(face.sha256);
    names.push(hash(bytes));
  }
  return Object.fromEntries(Object.entries<number[]>(chains.fontChains).map(([key, chain]) => [key, chain.map((id) => names[id])]));
}

let directory: string | undefined;
const prepared = built ? (async () => {
  directory = await mkdtemp(join(tmpdir(), 'native-chains-'));
  const product = await loadProduct(productRoot);
  const bytes = await fontDocx();
  const docx = join(directory, 'fonts.docx');
  await writeFile(docx, bytes);
  const { faces: expected, measurement } = await productFaces(product, bytes);
  const chainsPath = await writeDocumentChains(product, bytes, join(directory, 'chains'), 'fonts');
  return { directory, docx, chainsPath, expected, measurement, product };
})() : undefined;
prepared?.catch(() => {});

afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

test('the recorded chains file replays the product registry face for face', async () => {
  if (!prepared) {
    console.log(`native-chains: ${productRoot} has no dist build, skipping the product chain checks`);
    return;
  }
  const { chainsPath, expected, measurement, product } = await prepared;
  expect(describeAddons(product.fonts)).not.toContain('absent');
  expect(JSON.parse(await readFile(chainsPath, 'utf8')).measurement).toEqual(measurement);
  const keys = Object.keys(expected);
  for (const family of ['consolas', 'rockwell', 'microsoft sans serif', 'verdana', 'arial black', 'noto serif jp', 'times new roman', 'calibri'])
    expect({ family, present: keys.some((key) => key.toLowerCase().startsWith(`${family}|`)) }).toEqual({ family, present: true });
  expect(await chainsFaces(chainsPath)).toEqual(expected);
}, 300_000);

test('the native binary measures with the product chains it is handed', async () => {
  if (!binary) {
    console.log('native-chains: NATIVE_BENCH_BINARY is unset, skipping the native binary checks');
    return;
  }
  if (!prepared) throw new Error(`NATIVE_BENCH_BINARY needs the product built at ${productRoot}`);
  const { directory, docx, chainsPath, expected } = await prepared;
  let manifest = process.env.NATIVE_BENCH_FONTS;
  if (!manifest) {
    const fonts = spawnSync(process.execPath, [join(import.meta.dir, 'native-fonts.ts'), join(directory, 'fonts')], { encoding: 'utf8' });
    if (fonts.status !== 0) throw new Error(`native-fonts.ts exited ${fonts.status}: ${fonts.stderr}`);
    manifest = join(directory, 'fonts', 'manifest.json');
  }
  const run = (...chains: string[]) => {
    const result = spawnSync(binary, [docx, join(directory, 'page.png'), manifest!, ...chains], { encoding: 'utf8' });
    if (result.status !== 0) throw new Error(`${binary} exited ${result.status}: ${result.stderr}`);
    return JSON.parse(result.stdout) as { font_path: string; font_faces: Faces };
  };
  const replayed = run(chainsPath);
  expect(replayed.font_path).toBe('product-registry');
  expect(replayed.font_faces).toEqual(expected);
  const old = run();
  expect(old.font_path).toBe('bench-manifest');
  const files = new Map<string, string>(JSON.parse(await readFile(manifest, 'utf8')).faces.map((face: any) => [face.sha256, face.file]));
  const label = (chain: string[] | undefined) => chain?.map((face) =>
    face.split('~').map((part, index) => index ? part : files.get(part) ?? part.slice(0, 12)).join('~')).join(' > ') ?? '-';
  console.log(['family | old face | product face', ...[...new Set([...Object.keys(old.font_faces), ...Object.keys(expected)])].sort()
    .map((key) => `${key.split('|')[0]} | ${label(old.font_faces[key])} | ${label(expected[key])}`)].join('\n'));
  if (process.env.NATIVE_CHAINS_RED === '1') expect(old.font_faces).toEqual(expected);
}, 300_000);

test('the chains file keeps the measurement and the substitute arguments the product passed', () => {
  const measurement = { defaults: { fontSize: 11, fontFamily: 'Calibri' }, compat: { noLeading: false }, authoritativeShaping: true };
  const faces = [{ bytes: new Uint8Array([1]) }, { substituteOf: 0, family: 'Calibri' }, { substituteOf: 0, family: 'Arial', bold: true, italic: false }];
  const { chains } = serializeChains(faces, { fontChains: { 'calibri|0|0': [1, 0] }, ...measurement });
  expect(chains.measurement).toEqual(measurement);
  expect(chains.faces.slice(1)).toEqual([{ substituteOf: 0, family: 'Calibri' }, { substituteOf: 0, family: 'Arial', bold: true, italic: false }]);
  expect(() => serializeChains(faces, { fontChains: {}, ...measurement, fallbackIds: [0] })).toThrow('fallbackIds');
});

test('a document fails only when it needs an add-on its channel lacks', () => {
  const packages: FontPackages = {
    version: '0.0.0',
    fonts: {
      resolveBundledFamilyFace: (family: string) => (family === 'Noto Serif JP' ? { file: 'jp.ttf' } : { file: 'latin.ttf' }),
      resolveScriptFallbackFace: (script: string) => ({ file: script === 'cjk-sc' ? 'sc.ttf' : 'arabic.ttf' }),
    },
    addons: {
      'fonts-latin': { version: '1.0.0', files: new Set(['latin.ttf']), manifest: 'latin' },
      'fonts-intl': { version: '1.0.0', files: new Set(['thai.ttf']), manifest: 'intl' },
      'fonts-cjk': { files: new Set(['jp.ttf', 'sc.ttf']), manifest: 'cjk' },
    },
  };
  const requirement = (family: string, scripts: string[] = []) => ({ key: family, family, bold: false, italic: false, scripts });
  expect(missingAddons(packages, [requirement('Calibri', ['arabic'])])).toEqual([]);
  expect(missingAddons(packages, [requirement('Calibri', ['cjk-sc'])])).toEqual(['Calibri (cjk-sc script, sc.ttf) needs @betteroffice/fonts-cjk, which is absent']);
  expect(missingAddons(packages, [requirement('Noto Serif JP')])[0]).toContain('Noto Serif JP (family face');
  expect(describeAddons(packages)).toBe('fonts-latin 1.0.0, fonts-intl 1.0.0, fonts-cjk absent');
  const withoutLatin = { ...packages, addons: { ...packages.addons, 'fonts-latin': { files: new Set(['latin.ttf']), manifest: 'latin', error: 'broken' } } };
  expect(missingAddons(withoutLatin, [requirement('Calibri')])).toEqual(['Calibri (family face, latin.ttf) needs @betteroffice/fonts-latin, which is broken']);
});

test('the old font map is chosen only for a published release older than the first one with the registry exports', () => {
  expect(REGISTRY_EXPORTS_SINCE).toBe('0.6.0');
  expect(fontPathFor('published', '0.5.1')).toBe('bench-manifest');
  expect(fontPathFor('published', '0.6.0-next.3')).toBe('bench-manifest');
  expect(fontPathFor('published', '0.6.0')).toBe('product-registry');
  expect(fontPathFor('published', '0.10.0')).toBe('product-registry');
  expect(fontPathFor('commit', '0.5.1')).toBe('product-registry');
  expect(() => fontPathFor('published', 'latest')).toThrow('not a semver version');
  expect(compareSemver('0.6.0-next.2', '0.6.0-next.10')).toBe(-1);
  expect(compareSemver('0.6.0-next', '0.6.0-next.1')).toBe(-1);
  expect(compareSemver('0.6.0-1', '0.6.0-alpha')).toBe(-1);
  expect(compareSemver('1.0.0+build', '1.0.0')).toBe(0);
  expect(compareSemver('0.5.10', '0.6.0')).toBe(-1);
});

test('the bench reads the same first release with the registry exports as the recorder', async () => {
  const python = await readFile(join(import.meta.dir, 'docx_benchmark.py'), 'utf8');
  expect(python).toContain(`REGISTRY_EXPORTS_SINCE = '${REGISTRY_EXPORTS_SINCE}'`);
});

test('a registration the product registry drops fails the recording, naming the family', async () => {
  const registered: number[] = [];
  const engine = {
    registerFont: (bytes: Uint8Array) => {
      if (bytes[0] === 0) throw new Error('rejected');
      registered.push(bytes[0]);
      return registered.length - 1;
    },
    clearFonts: () => {},
  };
  const loader = (byte: number | Error) => Object.assign(async () => {
    if (byte instanceof Error) throw byte;
    return new Uint8Array([byte]).buffer;
  }, { measureAs: 'Carlito' });
  const loaders: Record<string, ReturnType<typeof loader>> = { Calibri: loader(1), Cambria: loader(new Error('offline')), Consolas: loader(0) };
  const marker = Symbol.for('@betteroffice/docx/hostedFontFamilies');
  const provider = { [marker]: undefined, resolve: (family: string) => loaders[family], resolveLastResort: (family: string) => loaders[family] };
  const guard = guardRegistrations(provider, engine);
  expect(guard.bundled!.resolve('Calibri')).toBe(guard.bundled!.resolve('Calibri'));
  expect(guard.bundled!.resolve('Calibri').measureAs).toBe('Carlito');
  expect(marker in guard.bundled!).toBe(true);
  const register = async (loaded: any) => guard.engine.registerFont(new Uint8Array(await loaded()));
  expect(await register(guard.bundled!.resolve('Calibri'))).toBe(0);
  expect(guard.failures.size).toBe(0);
  await expect(register(guard.bundled!.resolve('Cambria'))).rejects.toThrow('offline');
  await expect(register(guard.bundled!.resolveLastResort('Consolas'))).rejects.toThrow('rejected');
  expect(() => guard.engine.registerFont(new Uint8Array([0]))).toThrow('rejected');
  expect([...guard.failures]).toEqual(['Cambria: offline', 'Consolas (last resort): rejected']);
});

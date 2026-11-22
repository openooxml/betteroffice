import { afterAll, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import type { buildWorkerStub as BuildWorkerStub } from './worker-stub';
import { validateNativeWorker, validatePublishedWorker } from './worker-entry-check.mjs';
import { createFontPackResponder, type FontPackRequest, type FontPackSource } from '../shared/office-session/optionalFontPacks';

const buildWorkerStub: typeof BuildWorkerStub | undefined = await import('./worker-stub').then(
  (module) => module.buildWorkerStub,
  () => undefined,
);

const packages = resolve(import.meta.dir, '../packages');
const fontModules = resolve(import.meta.dir, '../shared/office-session/fontModules');
const stubs = [
  { name: 'docx', entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docx/docxWorker.mjs' },
  { name: 'docx legacy', entry: 'docx/src/yrs/docxWorker.browser.ts', outfile: 'docx/yrs/docxWorker.mjs', sibling: '../docxWorker.mjs' },
  { name: 'pptx', entry: 'pptx/src/session/worker.ts', outfile: 'pptx/pptxSessionWorker.mjs' },
  { name: 'xlsx', entry: 'xlsx/src/session/worker.ts', outfile: 'xlsx/xlsxSessionWorker.mjs' },
];
const directory = await mkdtemp(join(tmpdir(), 'worker-stub-'));

afterAll(() => rm(directory, { recursive: true, force: true }));

for (const { name, entry, outfile, ...options } of stubs) {
  test(`${name} worker bundles its start module and font runtime without imports`, async () => {
    expect(buildWorkerStub).toBeFunction();
    const file = join(directory, outfile);
    await buildWorkerStub!({ entry: join(packages, entry), outfile: file, ...options }, { browserFonts: true });
    const contents = await readFile(file, 'utf8');
    expect((await validatePublishedWorker(contents, 'sibling' in options ? options.sibling : null)).problems).toEqual([]);
    if ('sibling' in options) {
      expect(contents.trim()).toBe('import "../docxWorker.mjs";');
    } else {
      const urls = [...contents.matchAll(/new URL\(["']([^"']+)["'],\s*import\.meta\.url\)/g)].map(([, asset]) => asset!);
      expect(urls.length).toBeGreaterThan(0);
      for (const asset of urls) {
        expect(asset).toStartWith('./generated/');
        expect(new URL(asset, `https://workers.example/${outfile}`).pathname).toStartWith(`/${outfile.split('/')[0]}/generated/`);
      }
      expect(contents).not.toMatch(/data:application\/wasm/);
      expect(contents).toContain('font-pack-request');
      expect(contents).toContain('stopImmediatePropagation');
    }
  });
}

test('only browser worker bundles inline font imports and install the pack channel', async () => {
  const entry = join(directory, 'font-imports.ts');
  await writeFile(entry, `import { importFonts, importProgressiveFonts } from ${JSON.stringify(fontModules)};
globalThis.fonts = importFonts;
globalThis.progressive = importProgressiveFonts;`);
  const browser = join(directory, 'font-imports.browser.mjs');
  await buildWorkerStub!({ entry, outfile: browser }, { browserFonts: true });
  const source = await readFile(browser, 'utf8');
  expect((await validatePublishedWorker(source)).problems).toEqual([]);
  expect(source).toContain('font-pack-request');
  expect(source).toContain('decodeFontFaceOutsideEngine');
  expect(source).toContain('openSlicedFace');
  const native = join(directory, 'font-imports.native.mjs');
  await buildWorkerStub!({ entry, outfile: native }, { external: ['@betteroffice/fonts', '@betteroffice/fonts/*'] });
  const plain = await readFile(native, 'utf8');
  expect((await validateNativeWorker(plain)).problems).toEqual([]);
  expect(plain).toContain('import("@betteroffice/fonts")');
  expect(plain).toContain('import("@betteroffice/fonts/progressive")');
  expect(plain).not.toContain('font-pack-request');
  expect(plain).not.toContain('decodeFontFaceOutsideEngine');
});

async function runFontProbe<T>(name: string, contents: string, source: FontPackSource): Promise<T> {
  const entry = join(directory, `${name}.ts`);
  const outfile = join(directory, `${name}.mjs`);
  await writeFile(entry, contents);
  await buildWorkerStub!({ entry, outfile }, { browserFonts: true });
  const self = Object.assign(new EventTarget(), {
    postMessage(message: FontPackRequest) { responder.receive(message); },
  });
  const responder = createFontPackResponder((data) => self.dispatchEvent(new MessageEvent('message', { data })), source);
  const scope = globalThis as typeof globalThis & { fontProbe?: () => Promise<T> };
  const previous = Object.getOwnPropertyDescriptor(scope, 'self');
  Object.defineProperty(scope, 'self', { value: self, configurable: true });
  try {
    await import(outfile);
    return await scope.fontProbe!();
  } finally {
    if (previous) Object.defineProperty(scope, 'self', previous);
    else Reflect.deleteProperty(scope, 'self');
    delete scope.fontProbe;
  }
}

test('XLSX, PPTX and DOCX diagnostics preserve the missing font module rejection in browser workers', async () => {
  const formats = ['xlsx', 'pptx', 'docx'] as const;
  const imports = formats.map((format) => `import { fontPackDiagnostics as ${format} } from ${JSON.stringify(join(packages, `${format}/src/fontPackDiagnostics.ts`))};`);
  const results = await runFontProbe<PromiseSettledResult<unknown>[]>('missing-font-diagnostics', `
    ${imports.join('\n')}
    globalThis.fontProbe = () => Promise.allSettled([xlsx(), pptx(), docx()]);
  `, async () => undefined);
  for (const [index, format] of formats.entries()) {
    expect({ format, result: results[index] }).toMatchObject({ format, result: {
      status: 'rejected', reason: { name: 'TypeError', message: expect.stringContaining('@betteroffice/fonts') },
    } });
  }
});

test('PPTX command defaults settle without installing fonts when the module is unavailable', async () => {
  const result = await runFontProbe<{ landed: unknown[]; registered: unknown[]; rejected: string }>('missing-pptx-defaults', `
    import { commandFontProvider, createCommandDefaultFonts } from ${JSON.stringify(join(packages, 'pptx/src/session/commandDefaultFonts.ts'))};
    globalThis.fontProbe = async () => {
      const registered = [];
      const provider = () => commandFontProvider.load();
      const defaults = createCommandDefaultFonts([], () => [{ family: 'Calibri', bold: false, italic: false }],
        (face) => registered.push(face), () => {}, undefined, undefined, provider);
      try {
        const landed = await defaults.load();
        const rejected = await provider().then(() => '', (error) => error.name);
        return { landed, registered, rejected };
      } finally { defaults.dispose(); }
    };
  `, async () => undefined);
  expect(result).toEqual({ landed: [], registered: [], rejected: 'TypeError' });
});

test('DOCX export font selection rejects and progressive outlines request a whole-face reload without fonts', async () => {
  const result = await runFontProbe<{ selection: string; recorded: string; refused: number[]; logs: string[] }>('missing-docx-fonts', `
    import { resolveCompleteFontSelection } from ${JSON.stringify(join(packages, 'docx/src/layout/measure/defaultFontProvider.ts'))};
    import { recordProviderFaceLoads } from ${JSON.stringify(join(packages, 'docx/src/wasm/enginePngFonts.ts'))};
    import { ResidentProgressiveFonts } from ${JSON.stringify(join(packages, 'docx/src/yrs/residentProgressiveFonts.ts'))};
    globalThis.fontProbe = async () => {
      const selection = await resolveCompleteFontSelection().then(() => '', (error) => error.name);
      const recorded = await recordProviderFaceLoads({}).then(() => '', (error) => error.name);
      const logs = [];
      const refused = [];
      const spec = { slot: 0, source: { ref: { format: 1, index: 'slices.json', indexSha256: '' }, urls: {} } };
      const fonts = new ResidentProgressiveFonts([spec], [10], [],
        { updateFontOutlines() { throw new Error('unexpected outline install'); } }, async () => new ArrayBuffer(0),
        (message) => logs.push(message));
      fonts.onRefused((owner) => refused.push(owner));
      try {
        await fonts.ensure(new Map([[10, new Set([1])]]));
        await fonts.whenInstalled();
        return { selection, recorded, refused, logs };
      } finally { fonts.dispose(); }
    };
  `, async () => undefined);
  expect(result).toEqual({ selection: 'TypeError', recorded: 'TypeError', refused: [10],
    logs: [expect.stringContaining('the progressive font module is unavailable here')] });
});

test('a missing file decoder in the bundled worker takes the structured decoder error path', async () => {
  const urls = { 'fontpack.wasm': new URL('missing-worker-decoder/fontpack.wasm', import.meta.url).href };
  const error = await runFontProbe<Error>('missing-worker-decoder', `
    import { importFonts } from ${JSON.stringify(fontModules)};
    globalThis.fontProbe = async () => {
      const fonts = await importFonts();
      const frame = new Uint8Array([0x28, 0xb5, 0x2f, 0xfd, 0x24, 0, 1, 0, 0, 0, 0, 0, 0]);
      return fonts.decodeFontFaceOutsideEngine(frame, { codec: 'woff2t+zstd' }).catch((error) => error);
    };
  `, async (kind) => kind === 'base' ? {} : kind === 'decoder' ? urls : undefined);
  expect(error).toMatchObject({ message: expect.stringContaining('Font decoder wasm not found:'), cause: { code: 'ENOENT' } });
});

test('a docx build into another directory writes both worker stubs there and none into the package dist', async () => {
  const docx = join(packages, 'docx');
  const out = join(directory, 'docx-out');
  const published = ['dist/docxWorker.mjs', 'dist/yrs/docxWorker.mjs'].map((file) => join(docx, file));
  const snapshot = () => Promise.all(published.map((file) => readFile(file).then(
    (bytes) => createHash('sha256').update(bytes).digest('hex'),
    () => null,
  )));
  const before = await snapshot();
  const build = spawnSync(process.execPath, ['x', 'tsup', '--out-dir', out], {
    cwd: docx,
    encoding: 'utf8',
    env: { ...process.env, NODE_OPTIONS: '--max-old-space-size=8192' },
  });
  if (build.status !== 0) throw new Error(build.stdout + build.stderr);
  const emitted = ['docxWorker.mjs', 'yrs/docxWorker.mjs', 'yrs/residentEngineWorker.mjs'];
  const present = await Promise.all(emitted.map((file) => access(join(out, file)).then(() => file, () => `${file} missing`)));
  expect({ present, packageDist: await snapshot() }).toEqual({ present: emitted, packageDist: before });
}, 300_000);

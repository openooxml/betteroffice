import { afterEach, expect, test } from 'bun:test';
import { BUNDLED_FONTS } from '@betteroffice/fonts';
import { createFontProvider } from '@betteroffice/fonts/cdn';
import { prefetchFontChains, TextMeasureFontRegistry } from './fontRegistry';
import { createRustMeasureSource, type ResidentFontRequirement } from './rustMeasureSource';

const originalFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = originalFetch;
});

test("an unknown family asks the provider for Word's last-resort face", async () => {
  const offices: unknown[] = [];
  const registry = new TextMeasureFontRegistry(
    { registerFont: () => 3 },
    {
      bundled: {
        resolve: () => undefined,
        resolveLastResort: (_family, _bold, _italic, office) => {
          offices.push(office);
          return async () => new ArrayBuffer(16);
        },
      },
    }
  );
  expect(await registry.getFontIdChain('Lato', false, false)).toEqual([3]);
  expect(offices).toEqual(['word']);
});

const requirements: ResidentFontRequirement[] = [
  { family: 'Calibri', bold: false, italic: false },
  { family: 'Times New Roman', bold: true, italic: true },
  { family: 'Lato', bold: false, italic: true },
].map(({ family, bold, italic }) => ({
  key: `${family}|${Number(bold)}|${Number(italic)}`,
  family,
  bold,
  italic,
}));
const families = [...requirements.map(({ family }) => family), 'Courier New'];

const fileOf = (bytes: Uint8Array) =>
  new TextDecoder().decode(bytes).match(/^.+?\.(?:ttf|otf)/)![0];

async function run(scenario: string, warm: 'none' | 'settled' | 'in-flight') {
  const fetched: string[] = [];
  const registered: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    fetched.push(url);
    const face = BUNDLED_FONTS.find((face) => url.endsWith('/' + face.file));
    await Promise.resolve();
    if (!face) return new Response(null, { status: 404 });
    const body = new Uint8Array(face.byteLength);
    body.set(new TextEncoder().encode(face.file));
    return new Response(body);
  }) as typeof fetch;
  const provider = createFontProvider({ baseUrl: `https://fonts.test/${scenario}/` });
  const source = createRustMeasureSource({
    bundled: provider,
    engine: {
      registerFont(bytes) {
        registered.push(fileOf(bytes));
        return registered.length;
      },
      clearFonts() {},
    },
  });
  let prefetch: Promise<void> | undefined;
  let warmFetchCount = 0;
  if (warm === 'settled') {
    await prefetchFontChains(provider, families);
    expect(registered).toEqual([]);
    warmFetchCount = fetched.length;
  } else if (warm === 'in-flight') {
    prefetch = prefetchFontChains(provider, families);
  }
  await source.prepareFontRequirements(requirements);
  await prefetch;
  const { fontChains } = source.measurementConfigForRequirements(requirements)!;
  const chains = Object.fromEntries(
    requirements.map(({ key }) => [key, fontChains[key]!.map((id) => registered[id - 1])])
  );
  return { registered, chains, fetchedDuringPrepare: fetched.slice(warmFetchCount), fetched };
}

test('a chain prefetch registers the same faces in the same order as preparation without it', async () => {
  const cold = await run('none', 'none');
  for (const warm of ['settled', 'in-flight'] as const) {
    const warmed = await run(warm, warm);
    expect(warmed.registered).toEqual(cold.registered);
    expect(warmed.chains).toEqual(cold.chains);
  }
  expect(cold.chains).toEqual({
    'Calibri|0|0': ['Carlito-Regular.ttf', 'LiberationSans-Regular.ttf'],
    'Times New Roman|1|1': ['LiberationSerif-BoldItalic.ttf'],
    'Lato|0|1': ['LiberationSans-Italic.ttf'],
  });
});

test('a prefetched family no requirement uses registers nothing', async () => {
  const { registered, fetchedDuringPrepare, fetched } = await run('unused-family', 'settled');
  expect(fetched.some((url) => url.endsWith('/LiberationMono-Regular.ttf'))).toBe(true);
  expect(fetchedDuringPrepare).toEqual([]);
  const courierFiles = BUNDLED_FONTS.filter((face) => face.metricCompatWith === 'Courier New').map(
    (face) => face.file
  );
  expect(registered.filter((file) => courierFiles.includes(file))).toEqual([]);
});

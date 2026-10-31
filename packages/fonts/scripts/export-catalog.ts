import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import * as ts from 'typescript';
import * as manifest from '../src/manifest';
import { SCRIPT_PATTERNS } from '../src/scripts';
import latin from '../../fonts-latin/assets/manifest.json';
import intl from '../../fonts-intl/assets/manifest.json';
import cjk from '../../fonts-cjk/assets/manifest.json';

const packs = [latin, intl, cjk];
const output = new URL('../../../crates/betteroffice-fonts/data/', import.meta.url);
const source = readFileSync(new URL('../src/manifest.ts', import.meta.url), 'utf8');
const privateNames = [
  'UNVENDORED_NOTO_SERIF', 'WIDTH_MATCHED_SUBSTITUTES',
  'LIGHTER_THAN_REGULAR', 'HEAVIER_THAN_REGULAR', 'looksSerif', 'looksMono',
];
const compiled = ts.transpileModule(`${source}\nexport { ${privateNames.join(', ')} };`, {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, esModuleInterop: true },
}).outputText;
const tables = {} as {
  UNVENDORED_NOTO_SERIF: Record<string, string>;
  WIDTH_MATCHED_SUBSTITUTES: Map<string, { family: string; digitAdvances: Record<string, number> }>;
  LIGHTER_THAN_REGULAR: Set<string>;
  HEAVIER_THAN_REGULAR: Set<string>;
  looksSerif: (family: string) => boolean;
  looksMono: (family: string) => boolean;
};
new Function('exports', 'require', compiled)(tables, (path: string) =>
  JSON.parse(readFileSync(new URL(path, new URL('../src/', import.meta.url)), 'utf8')));

function markers(fn: (family: string) => boolean): string[] {
  return [...fn.toString().matchAll(/\.includes\(['"]([^'"]+)['"]\)/g)].map((match) => match[1]!);
}

function pattern(fn: (family: string) => boolean): string {
  const value = fn.toString().match(/\/(.+)\/\.test\(lower\)/)?.[1];
  if (value === undefined) throw new Error('Missing family classification pattern');
  return value;
}

function faceFile(face: manifest.BundledFontFace | undefined): string | null {
  return face?.file ?? null;
}

export function exportCatalog(): Record<string, string> {
  const faces = manifest.BUNDLED_FONTS.map((face) => {
    const pack = packs.find((pack) => pack.faces.some((entry) => entry.file === face.file));
    if (pack === undefined) throw new Error(`Missing pack for ${face.file}`);
    const entry = pack.faces.find((entry) => entry.file === face.file)!;
    return { pack: pack.pack, ...face, version: pack.version, codec: entry.codec,
      dict: entry.dict, encodedLength: entry.encodedLength };
  });
  const scripts = [...new Set(faces.flatMap((face) => face.script === undefined ? [] : [face.script]))];
  const catalog = {
    faces,
    wordFamilyAliases: manifest.WORD_FAMILY_ALIASES,
    unvendoredNotoSerif: tables.UNVENDORED_NOTO_SERIF,
    lighterThanRegular: [...tables.LIGHTER_THAN_REGULAR],
    heavierThanRegular: [...tables.HEAVIER_THAN_REGULAR],
    serifMarkers: markers(tables.looksSerif),
    serifExclusionPattern: pattern(tables.looksSerif),
    monoMarkers: markers(tables.looksMono),
    monoPattern: pattern(tables.looksMono),
    widthMatchedSubstitutes: Object.fromEntries(tables.WIDTH_MATCHED_SUBSTITUTES),
    scriptFallbackOrder: Object.fromEntries(scripts.map((script) => [
      script, faces.filter((face) => face.script === script).map((face) => face.file),
    ])),
    lastResortFamilies: Object.fromEntries((['powerpoint', 'word'] as const).map((office) => [office, {
      sans: manifest.resolveLastResortFace('Unknown', false, false, office).family,
      serif: manifest.resolveLastResortFace('Unknown Serif', false, false, office).family,
      mono: manifest.resolveLastResortFace('Unknown Mono', false, false, office).family,
    }])),
    glyphFallbackFamilies: [...new Set(faces.filter((face) =>
      manifest.resolveGlyphFallbackFace(face.metricCompatWith ?? face.family, false, false) !== undefined,
    ).map((face) => face.family))],
  };
  const families = [...new Set([
    ...faces.map((face) => face.metricCompatWith ?? face.family),
    ...Object.keys(manifest.WORD_FAMILY_ALIASES),
    ...faces.map((face) => face.family),
    ...Object.keys(tables.UNVENDORED_NOTO_SERIF),
    ...[...tables.LIGHTER_THAN_REGULAR].map((weight) => `Calibri ${weight}`),
    ...[...tables.HEAVIER_THAN_REGULAR].flatMap((weight) => [`Calibri ${weight}`, `Segoe UI ${weight}`]),
    ...['Calibri Light', 'Inter-Thin', 'Segoe UI Light', 'Segoe UI Semilight', 'Calibri Black',
      'Arial Heavy', 'Segoe UI Heavy', 'Unknown', 'Unknown Serif', 'Unknown Mono',
      'Consolas', 'Baskerville', 'Sans-serif', 'Verdana', 'Calibri Unknown', ''],
  ])];
  const cases: object[] = [];
  for (const family of [...families, '  CALIBRI  ', '  ＭＳ ゴシック ', 'not-a-family', 'Calibri\tLight',
    'mono\nface', 'Calibri-Light-', 'Calibri--Light', '\ufeffCalibri\ufeff', '\u0085Calibri\u0085']) {
    for (const bold of [false, true]) for (const italic of [false, true]) {
      const offices = (resolve: (office: manifest.LastResortOffice) => unknown) =>
        Object.fromEntries((['powerpoint', 'word'] as const).map((office) => [office, resolve(office)]));
      cases.push({ query: { family, bold, italic }, answer: {
        metricCompatFamily: manifest.resolveMetricCompatFamily(family) ?? null,
        metricCompatFace: faceFile(manifest.resolveMetricCompatFace(family, bold, italic)),
        bundledFamilyFace: faceFile(manifest.resolveBundledFamilyFace(family, bold, italic)),
        lightVariantFace: faceFile(manifest.resolveLightVariantFace(family, bold, italic)),
        digitAdvance: manifest.resolveDigitAdvance(family, bold, italic) ?? null,
        lastResortChoice: offices((office) => {
          const choice = manifest.resolveLastResortChoice(family, bold, italic, office);
          return { face: choice.face.file, onlyIfPrimaryFails: choice.onlyIfPrimaryFails };
        }),
        lastResortFace: offices((office) => faceFile(manifest.resolveLastResortFace(family, bold, italic, office))),
        glyphFallbackFace: offices((office) => faceFile(manifest.resolveGlyphFallbackFace(family, bold, italic, office))),
      } });
    }
  }
  const samples = Object.fromEntries(SCRIPT_PATTERNS.map(([script, regex]) => {
    const match = regex.source.match(/\\u(?:\{([0-9a-f]+)\}|([0-9a-f]{4}))/i)!;
    return [script, parseInt(match[1] ?? match[2]!, 16)];
  }));
  Object.assign(samples, { 'cjk-sc': 0x4e2d, 'cjk-tc': 0x570b, 'cjk-jp': 0x3042, 'cjk-kr': 0xac00 });
  for (const script of scripts) for (const bold of [false, true]) for (const italic of [false, true]) {
    cases.push({ query: { script, codePoint: samples[script], bold, italic }, answer: {
      scriptFallbackFace: faceFile(manifest.resolveScriptFallbackFace(script, bold, italic)),
      glyphFallbackFace: faceFile(manifest.resolveGlyphFallbackFace('Segoe UI', bold, italic)),
    } });
  }
  return {
    'catalog.json': `${JSON.stringify(catalog, null, 2)}\n`,
    'resolve-cases.json': `[\n${cases.map((row) => JSON.stringify(row)).join(',\n')}\n]\n`,
  };
}

if (import.meta.main) {
  mkdirSync(output, { recursive: true });
  for (const [file, data] of Object.entries(exportCatalog())) writeFileSync(new URL(file, output), data);
}

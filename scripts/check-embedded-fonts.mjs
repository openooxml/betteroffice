import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const ALLOWLIST = [
  { source: 'crates/xlsx-raster/src/font.rs', asset: '../assets/Carlito-Regular.ttf', reason: 'Migrates in PR B' },
  { source: 'crates/vsdx-raster/src/lib.rs', asset: '../../xlsx-raster/assets/Carlito-Regular.ttf', reason: 'Migrates in PR B' },
  { source: 'apps/native-viewer/src/xlsx_scene.rs', asset: '../../../crates/xlsx-raster/assets/Carlito-Regular.ttf', reason: 'Uses the xlsx-raster asset until PR B' },
  { source: 'apps/native-viewer/src/pptx_scene.rs', asset: '../../../packages/fonts/assets/LiberationSans-Regular.ttf', reason: 'Full Arial substitutes until the pptx wiring PR' },
  { source: 'apps/native-viewer/src/pptx_scene.rs', asset: '../../../packages/fonts/assets/LiberationSans-Bold.ttf', reason: 'Full Arial substitutes until the pptx wiring PR' },
  { source: 'apps/native-viewer/src/pptx_scene.rs', asset: '../../../packages/fonts/assets/LiberationSans-Italic.ttf', reason: 'Full Arial substitutes until the pptx wiring PR' },
  { source: 'apps/native-viewer/src/pptx_scene.rs', asset: '../../../packages/fonts/assets/LiberationSans-BoldItalic.ttf', reason: 'Full Arial substitutes until the pptx wiring PR' },
];

export function embeddedFontViolations(file, source) {
  const path = file.replaceAll('\\', '/');
  if (/(^|\/)(tests|benches|examples)\//.test(path) || /(^|\/|_)tests\.rs$/.test(path)) return [];
  const production = source.split(/#\s*\[\s*cfg\s*\(\s*(?:test\b|all\s*\([^\]]*\btest\b)/, 1)[0];
  const violations = [];
  for (const match of production.matchAll(/include_bytes!\s*\(\s*(?:r#*)?"([^"\n]+\.(?:ttf|otf|woff2?))"/gi)) {
    const asset = match[1];
    const target = relative(ROOT, resolve(ROOT, dirname(path), asset)).replaceAll('\\', '/');
    if (path.startsWith('crates/base-latin/') && target.startsWith('crates/base-latin/fonts/')) continue;
    if (ALLOWLIST.some(entry => entry.source === path && entry.asset === asset)) continue;
    violations.push(`${path}: ${asset}`);
  }
  return violations;
}

export function lastResortFamilyMatches(rust, manifest) {
  const family = rust.match(/pub const FAMILY: &str = "([^"]+)";/)?.[1];
  const lastResort = manifest.match(/export const BASE_LATIN_FAMILY = ['"]([^'"]+)['"];/)?.[1];
  return family !== undefined && family === lastResort;
}

function rustFiles(path) {
  return readdirSync(path, { withFileTypes: true }).flatMap(entry => {
    if (['target', 'node_modules', 'vendor', '.git'].includes(entry.name)) return [];
    const child = join(path, entry.name);
    if (entry.isDirectory()) return rustFiles(child);
    return entry.isFile() && entry.name.endsWith('.rs') ? [child] : [];
  });
}

export function checkEmbeddedFonts(root = ROOT) {
  const violations = ['crates', 'apps'].flatMap(directory =>
    rustFiles(join(root, directory)).flatMap(file =>
      embeddedFontViolations(relative(root, file), readFileSync(file, 'utf8'))));
  if (!lastResortFamilyMatches(
    readFileSync(join(root, 'crates/base-latin/src/lib.rs'), 'utf8'),
    readFileSync(join(root, 'packages/fonts/src/manifest.ts'), 'utf8'),
  )) violations.push('The font manifest last-resort family differs from base_latin::FAMILY');
  return violations;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const violations = checkEmbeddedFonts();
  for (const violation of violations) console.error(violation);
  process.exitCode = violations.length === 0 ? 0 : 1;
}

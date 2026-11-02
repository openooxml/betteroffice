import { expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FONT_PACK_CRATES, prepareFontPacks } from './rust-font-packs.mjs';

const root = new URL('../', import.meta.url);

test('stages self-contained byte crates with only their own faces and licenses', () => {
  const temp = mkdtempSync(join(tmpdir(), 'font-packs-'));
  try {
    for (const pack of ['latin', 'intl', 'cjk']) {
      const manifest = JSON.parse(readFileSync(new URL(`packages/fonts-${pack}/assets/manifest.json`, root), 'utf8'));
      const source = join(temp, 'packages', `fonts-${pack}`);
      mkdirSync(join(source, 'assets'), { recursive: true });
      mkdirSync(join(source, 'LICENSES'));
      writeFileSync(join(source, 'LICENSE'), `${pack} license`);
      writeFileSync(join(source, 'LICENSES', 'OFL.txt'), `${pack} font license`);
      writeFileSync(join(source, 'assets', 'manifest.json'), JSON.stringify(manifest));
      for (const face of manifest.faces) writeFileSync(join(source, 'assets', face.file), face.file);
    }
    prepareFontPacks(temp);
    for (const crate of FONT_PACK_CRATES) {
      const target = join(temp, 'crates', crate.name);
      const manifest = JSON.parse(readFileSync(join(target, 'assets', 'manifest.json'), 'utf8'));
      expect(manifest.pack).toBe(crate.pack);
      expect(readFileSync(join(target, 'LICENSE'), 'utf8')).toBe(`${crate.pack} license`);
      expect(readFileSync(join(target, 'LICENSES', 'OFL.txt'), 'utf8')).toBe(`${crate.pack} font license`);
      if (crate.family) {
        expect(manifest.faces).toHaveLength(1);
        expect(manifest.faces[0].family).toBe(crate.family);
      }
      expect(readdirSync(join(target, 'assets')).sort()).toEqual([
        'manifest.json', ...(crate.group ? [] : manifest.faces.map((face: { file: string }) => face.file)),
      ].sort());
      if (!crate.group) for (const face of manifest.faces) {
        expect(readFileSync(join(target, 'assets', face.file), 'utf8')).toBe(face.file);
      }
      writeFileSync(join(target, 'assets', 'stale.ttf.zst'), 'stale');
    }
    prepareFontPacks(temp);
    for (const crate of FONT_PACK_CRATES) {
      expect(readdirSync(join(temp, 'crates', crate.name, 'assets'))).not.toContain('stale.ttf.zst');
    }
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
});

test('embeds each face once and keeps byte crates under the registry size limit', () => {
  for (const crate of FONT_PACK_CRATES) {
    const manifest = JSON.parse(readFileSync(new URL(`packages/fonts-${crate.pack}/assets/manifest.json`, root), 'utf8'));
    const faces = manifest.faces.filter((face: { family: string }) => !crate.family || face.family === crate.family);
    const source = readFileSync(new URL(`crates/${crate.name}/src/lib.rs`, root), 'utf8');
    const embedded = [...source.matchAll(/static FACE_\d+: &\[u8\] =\s*include_bytes!\(concat!\(\s*env!\("FONT_PACK_ASSETS"\),\s*"\/([^\"]+)"\s*,?\s*\)\);/g)]
      .map((match) => match[1]).sort();
    expect(embedded).toEqual(crate.group ? [] : faces.map((face: { file: string }) => face.file).sort());
    if (!crate.group) {
      expect(faces.reduce((sum: number, face: { encodedLength: number }) => sum + face.encodedLength, 0))
        .toBeLessThan(10_000_000);
    }
    const cargo = Bun.TOML.parse(readFileSync(new URL(`crates/${crate.name}/Cargo.toml`, root), 'utf8'));
    expect(cargo.package.include).toContain('assets/**');
    expect(cargo.package.include).toContain('LICENSES/**');
    expect(cargo.package.license).toBe('Apache-2.0 AND OFL-1.1');
  }
});

test('stages assets before release metadata and packaging', () => {
  const publish = readFileSync(new URL('scripts/publish-crates.mjs', root), 'utf8');
  const start = publish.indexOf('async function publish()');
  expect(publish.indexOf('prepareFontPacks();', start)).toBeLessThan(publish.indexOf('cargoMetadata(', start));
});

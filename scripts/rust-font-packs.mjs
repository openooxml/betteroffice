import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const FONT_PACK_CRATES = [
  { name: 'betteroffice-fonts-latin', pack: 'latin' },
  { name: 'betteroffice-fonts-intl', pack: 'intl' },
  { name: 'betteroffice-fonts-cjk-sc', pack: 'cjk', family: 'NotoSansSC' },
  { name: 'betteroffice-fonts-cjk-tc', pack: 'cjk', family: 'NotoSansTC' },
  { name: 'betteroffice-fonts-cjk-jp', pack: 'cjk', family: 'NotoSansJP' },
  { name: 'betteroffice-fonts-cjk-kr', pack: 'cjk', family: 'NotoSansKR' },
  { name: 'betteroffice-fonts-cjk-serif-sc', pack: 'cjk', family: 'NotoSerifSC' },
  { name: 'betteroffice-fonts-cjk', pack: 'cjk', group: true }
];

export function prepareFontPacks(root = fileURLToPath(new URL('..', import.meta.url))) {
  for (const crate of FONT_PACK_CRATES) {
    const source = join(root, 'packages', `fonts-${crate.pack}`);
    const target = join(root, 'crates', crate.name);
    const assets = join(target, 'assets');
    const manifest = JSON.parse(readFileSync(join(source, 'assets', 'manifest.json'), 'utf8'));
    if (crate.family) manifest.faces = manifest.faces.filter((face) => face.family === crate.family);
    if (manifest.faces.length === 0) throw new Error(`No font faces for ${crate.name}`);
    rmSync(assets, { recursive: true, force: true });
    mkdirSync(assets, { recursive: true });
    writeFileSync(join(assets, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
    if (!crate.group) {
      for (const face of manifest.faces) {
        cpSync(join(source, 'assets', face.file), join(assets, face.file));
      }
    }
    cpSync(join(source, 'LICENSE'), join(target, 'LICENSE'));
    cpSync(join(source, 'LICENSES'), join(target, 'LICENSES'), { recursive: true });
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) prepareFontPacks();

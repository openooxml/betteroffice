import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { BUNDLED_FONTS, WORD_FAMILY_ALIASES } from '../../packages/fonts/src/manifest';
import { loadPackedFontBytes } from '../../packages/fonts/src/packed';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const output = resolve(process.argv[2] ?? '.source/office-quality/native-fonts');
await mkdir(output, { recursive: true });
const faces = await Promise.all(BUNDLED_FONTS.map(async (face) => {
  const source = resolve(root, 'packages', face.script?.startsWith('cjk-') ? 'fonts-cjk' : 'fonts', 'assets', face.file);
  const bytes = face.file.endsWith('.ttf.zst')
    ? new Uint8Array(await loadPackedFontBytes(
        face,
        async (dependency) => pathToFileURL(resolve(root, 'packages/fonts-intl/assets', dependency.file)),
      ))
    : await readFile(source);
  if (bytes.length !== face.byteLength) throw new Error(`Font size mismatch: ${face.file}`);
  const file = face.file.replace(/\.zst$/, '');
  await writeFile(resolve(output, file), bytes);
  return { ...face, file, sha256: createHash('sha256').update(bytes).digest('hex') };
}));
await writeFile(resolve(output, 'manifest.json'), JSON.stringify({ faces, aliases: WORD_FAMILY_ALIASES }));

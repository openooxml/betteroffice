import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { BUNDLED_FONTS, WORD_FAMILY_ALIASES } from '../../packages/fonts/src/manifest';
import { loadPackedFontBytes, isLatinFont } from '../../packages/fonts/src/packed';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const output = resolve(process.argv[2] ?? '.source/office-quality/native-fonts');
await mkdir(output, { recursive: true });
const faces = await Promise.all(BUNDLED_FONTS.map(async (face) => {
  const latin = isLatinFont(face.file);
  const file = face.file.replace(/\.zst$/, '');
  const source = resolve(root, 'packages', face.script?.startsWith('cjk-') ? 'fonts-cjk' : 'fonts', 'assets', file);
  const bytes = !latin && face.file.endsWith('.ttf.zst')
    ? new Uint8Array(await loadPackedFontBytes(
        face,
        async (dependency) => pathToFileURL(resolve(root, 'packages/fonts-intl/assets', dependency.file)),
      ))
    : await readFile(source);
  if (!latin && bytes.length !== face.byteLength) throw new Error(`Font size mismatch: ${face.file}`);
  await writeFile(resolve(output, file), bytes);
  return { ...face, file, byteLength: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') };
}));
await writeFile(resolve(output, 'manifest.json'), JSON.stringify({ faces, aliases: WORD_FAMILY_ALIASES }));

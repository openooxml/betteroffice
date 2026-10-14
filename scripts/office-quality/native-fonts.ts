import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { BUNDLED_FONTS, WORD_FAMILY_ALIASES } from '../../packages/fonts/src/manifest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const output = resolve(process.argv[2] ?? '.source/office-quality/native-fonts');
await mkdir(output, { recursive: true });
const faces = await Promise.all(BUNDLED_FONTS.map(async (face) => {
  const source = resolve(root, 'packages', face.script?.startsWith('cjk-') ? 'fonts-cjk' : 'fonts', 'assets', face.file);
  const bytes = await readFile(source);
  if (bytes.length !== face.byteLength) throw new Error(`Font size mismatch: ${face.file}`);
  await copyFile(source, resolve(output, face.file));
  return { ...face, sha256: createHash('sha256').update(bytes).digest('hex') };
}));
await writeFile(resolve(output, 'manifest.json'), JSON.stringify({ faces, aliases: WORD_FAMILY_ALIASES }));

import { mkdir } from 'node:fs/promises';
import tailwindcss from '@tailwindcss/postcss';
import postcss from 'postcss';

const from = `${import.meta.dir}/src/styles.css`;
const to = `${import.meta.dir}/dist/styles.css`;
const result = await postcss([tailwindcss({ optimize: { minify: true } })]).process(await Bun.file(from).text(), {
  from,
  to,
});
await Bun.write(to, result.css);
console.log(`styles.css ${(result.css.length / 1024).toFixed(1)} KB`);

const fontsDir = `${import.meta.dir}/dist/fonts`;
await mkdir(fontsDir, { recursive: true });
for (const specifier of [
  '@fontsource-variable/geist/files/geist-latin-wght-normal.woff2',
  '@fontsource-variable/geist-mono/files/geist-mono-latin-wght-normal.woff2',
]) {
  const source = Bun.resolveSync(specifier, import.meta.dir);
  await Bun.write(`${fontsDir}/${specifier.split('/').pop()}`, Bun.file(source));
}

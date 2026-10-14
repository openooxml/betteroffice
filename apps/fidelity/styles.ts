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

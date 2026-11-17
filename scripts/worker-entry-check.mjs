import { build } from 'esbuild';

const dynamicImport = /(?<![\w$.])import\s*\(/g;

/** Parses `source` with esbuild; a problem for any static import or require, and unless it has exactly one dynamic import, whose argument is a string literal; `target` is that literal, else null. */
export async function validatePublishedWorker(source) {
  const imports = [];
  let output;
  try {
    const result = await build({
      stdin: { contents: source, loader: 'js' },
      bundle: true, write: false, format: 'esm', platform: 'browser',
      minifyWhitespace: true, legalComments: 'none', logLevel: 'silent',
      plugins: [{ name: 'record-imports', setup(stub) {
        stub.onResolve({ filter: /.*/ }, ({ path, kind }) => {
          imports.push({ path, kind });
          return { path, external: true };
        });
      } }],
    });
    output = result.outputFiles[0].text;
  } catch (error) {
    return { problems: [`does not parse: ${error.message}`], target: null };
  }
  const problems = imports.filter(({ kind }) => kind !== 'dynamic-import')
    .map(({ path, kind }) => `statically imports ${path} (${kind}); it must only import its code dynamically`);
  const literal = imports.filter(({ kind }) => kind === 'dynamic-import').map(({ path }) => path);
  const dynamic = output.match(dynamicImport)?.length ?? 0;
  if (dynamic !== 1) problems.push(`has ${dynamic} dynamic imports; it must load its code with exactly one`);
  else if (literal.length !== 1) problems.push('its dynamic import target is not a string literal');
  return { problems, target: dynamic === 1 && literal.length === 1 ? literal[0] : null };
}

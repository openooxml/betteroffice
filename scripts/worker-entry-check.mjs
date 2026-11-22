import { build } from 'esbuild';

async function workerImports(source) {
  let imports = [];
  let output;
  try {
    const result = await build({
      stdin: { contents: source, loader: 'js' },
      bundle: true, write: false, format: 'esm', platform: 'browser',
      minifyWhitespace: true, legalComments: 'none', logLevel: 'silent', metafile: true,
      plugins: [{ name: 'record-imports', setup(stub) {
        stub.onResolve({ filter: /.*/ }, ({ path }) => ({ path, external: true }));
      } }],
    });
    output = result.outputFiles[0].text;
    imports = Object.values(result.metafile.inputs).flatMap((entry) => entry.imports);
  } catch (error) {
    return { imports, output, problems: [`does not parse: ${error.message}`] };
  }
  return { imports, output, problems: [] };
}

/** Validates that a published browser worker is self-contained. */
export async function validatePublishedWorker(source, sibling = null) {
  const { imports, output, problems } = await workerImports(source);
  if (problems.length) return { problems, target: null };
  for (const { path, kind } of imports) {
    if (kind === 'dynamic-import') problems.push(`dynamically imports ${path}`);
    else if (kind !== 'import-statement' || path !== sibling) {
      problems.push(`imports ${path} (${kind})`);
    }
  }
  if (!imports.some(({ kind }) => kind === 'dynamic-import')) {
    for (const [, keyword] of output.matchAll(/"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|(?<![\w$.])\b(import)\s*\(/gs)) {
      if (!keyword) continue;
      problems.push('contains a dynamic import');
      break;
    }
  }
  for (const [, path] of output.matchAll(/\bnew URL\(["'`]([^"'`]+)["'`],\s*import\.meta\.url\)/g)) {
    if (!path.startsWith('./generated/') || /fontpack\.wasm|\.(?:zst|[ot]tf|woff2?)$|slices\.json/.test(path)) {
      problems.push(`resolves ${path} relative to the worker module`);
    }
  }
  const statics = imports.filter(({ kind }) => kind === 'import-statement');
  if (statics.length !== Number(sibling !== null)) problems.push(`has ${statics.length} static imports; expected ${Number(sibling !== null)}`);
  const target = sibling && statics.some(({ path }) => path === sibling) ? sibling : null;
  if (sibling && target === null) problems.push('sibling entry must statically import its canonical bundle');
  if (target && output.trim() !== `import${JSON.stringify(target)};`) problems.push('sibling entry must contain only its static import');
  return { problems, target };
}

export async function validateNativeWorker(source, sibling = null) {
  if (sibling) return validatePublishedWorker(source, sibling);
  const { imports, problems } = await workerImports(source);
  for (const { path, kind } of imports) {
    if (/^(?:\.{1,2}(?:\/|$)|\/|[A-Za-z]:[\\/])/.test(path)) problems.push(`imports relative or absolute path ${path} (${kind})`);
    if (path === '@betteroffice/fonts/worker') problems.push('imports the browser worker font entry');
  }
  return { problems, target: null };
}

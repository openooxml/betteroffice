import { fileURLToPath } from 'node:url';

// The Python release train. Adding a distribution here enrols it in versioning,
// CI, and wheel builds; see RELEASING.md for the PyPI side.
// `publish: false` holds it out of the PyPI matrix until its project is ready.
const REGISTRY = [
  { path: 'bindings/python-docx', publish: true },
  { path: 'bindings/python-pptx', publish: true },
  { path: 'bindings/python-xlsx', publish: true }
];

function bindingName(path) {
  return path.replace('bindings/python-', '');
}

export const PYTHON_BINDINGS = REGISTRY.map((entry) => entry.path);

export const PYTHON_BINDING_NAMES = PYTHON_BINDINGS.map(bindingName);

export const PYTHON_PUBLISH_NAMES = REGISTRY.filter((entry) => entry.publish).map((entry) =>
  bindingName(entry.path)
);

/** The PyPI projects that exist, so a `publish: false` binding is absent. */
export const PYPI_DISTRIBUTIONS = PYTHON_PUBLISH_NAMES.map((name) => `betteroffice-${name}`);

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  // Falling through to the full list would publish bindings deliberately held back.
  const unknown = args.find((arg) => arg !== '--paths' && arg !== '--publish');
  if (unknown) {
    console.error(`python-bindings.mjs: unknown argument ${unknown}; expected --paths or --publish`);
    process.exit(1);
  }
  if (args.includes('--paths')) {
    console.log(PYTHON_BINDINGS.join('\n'));
  } else if (args.includes('--publish')) {
    console.log(JSON.stringify(PYTHON_PUBLISH_NAMES));
  } else {
    console.log(JSON.stringify(PYTHON_BINDING_NAMES));
  }
}

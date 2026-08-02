import { fileURLToPath } from 'node:url';

// The Python release train. Adding a distribution here enrols it in versioning,
// CI, and wheel builds; see RELEASING.md for the PyPI side.
// `publish: false` holds it out of the PyPI matrix until its project is ready.
const REGISTRY = [
  { path: 'bindings/python-pptx', publish: false },
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

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  if (process.argv.includes('--paths')) {
    console.log(PYTHON_BINDINGS.join('\n'));
  } else if (process.argv.includes('--publish')) {
    console.log(JSON.stringify(PYTHON_PUBLISH_NAMES));
  } else {
    console.log(JSON.stringify(PYTHON_BINDING_NAMES));
  }
}

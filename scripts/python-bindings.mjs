import { fileURLToPath } from 'node:url';

// The Python release train. Adding a distribution here enrols it in versioning,
// CI, and publishing; see RELEASING.md for the PyPI side.
export const PYTHON_BINDINGS = ['bindings/python-pptx', 'bindings/python-xlsx'];

export const PYTHON_BINDING_NAMES = PYTHON_BINDINGS.map((binding) =>
  binding.replace('bindings/python-', '')
);

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const paths = process.argv.includes('--paths');
  console.log(paths ? PYTHON_BINDINGS.join('\n') : JSON.stringify(PYTHON_BINDING_NAMES));
}

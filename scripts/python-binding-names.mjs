// The Python release train. Adding a distribution here enrols it in versioning,
// CI, and wheel builds; see RELEASING.md for the PyPI side.
// `publish: false` holds it out of the PyPI matrix until its project is ready.
// Kept free of Node imports so the web worker can bundle the name list.
export const REGISTRY = [
  { path: 'bindings/python-docx', publish: true },
  { path: 'bindings/python-pptx', publish: true },
  { path: 'bindings/python-xlsx', publish: true }
];

export function bindingName(path) {
  return path.replace('bindings/python-', '');
}

export const PYTHON_BINDINGS = REGISTRY.map((entry) => entry.path);

export const PYTHON_BINDING_NAMES = PYTHON_BINDINGS.map(bindingName);

export const PYTHON_PUBLISH_NAMES = REGISTRY.filter((entry) => entry.publish).map((entry) =>
  bindingName(entry.path)
);

/** The PyPI projects that exist, so a `publish: false` binding is absent. */
export const PYPI_DISTRIBUTIONS = PYTHON_PUBLISH_NAMES.map((name) => `betteroffice-${name}`);

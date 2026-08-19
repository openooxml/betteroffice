import { readFileSync } from 'node:fs';
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

/** The crate version maturin stamps on the wheel; bindings version independently of the workspace. */
export function bindingVersion(path) {
  const manifest = readFileSync(new URL(`../${path}/Cargo.toml`, import.meta.url), 'utf8');
  const version = manifest.match(/^version = "([^"]+)"$/m);
  if (!version) throw new Error(`${path}/Cargo.toml declares no version`);
  return version[1];
}

/** Publish-enabled bindings whose crate version PyPI does not serve yet. */
export async function pendingPublishNames({
  fetchImpl = fetch,
  attempts = 3,
  retryDelayMs = 2000
} = {}) {
  const pending = [];
  for (const entry of REGISTRY) {
    if (!entry.publish) continue;
    const name = bindingName(entry.path);
    const released = await pypiVersions(`betteroffice-${name}`, { fetchImpl, attempts, retryDelayMs });
    if (released === null || !released.includes(bindingVersion(entry.path))) pending.push(name);
  }
  return pending;
}

/** null when the project does not exist yet; throws when PyPI cannot be read at all. */
async function pypiVersions(project, { fetchImpl, attempts, retryDelayMs }) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetchImpl(`https://pypi.org/pypi/${project}/json`, {
        headers: { Accept: 'application/json' },
        signal: AbortSignal.timeout(15_000)
      });
      if (response.status === 404) return null;
      if (!response.ok) throw new Error(`PyPI answered ${response.status} for ${project}`);
      return Object.keys((await response.json()).releases ?? {});
    } catch (error) {
      lastError = error;
      if (attempt < attempts) {
        await new Promise((resolve) => setTimeout(resolve, retryDelayMs * attempt));
      }
    }
  }
  throw new Error(`cannot read PyPI for ${project}: ${lastError}`, { cause: lastError });
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  // Falling through to the full list would publish bindings deliberately held back.
  const unknown = args.find((arg) => !['--paths', '--publish', '--pending'].includes(arg));
  if (unknown) {
    console.error(
      `python-bindings.mjs: unknown argument ${unknown}; expected --paths, --publish or --pending`
    );
    process.exit(1);
  }
  if (args.includes('--paths')) {
    console.log(PYTHON_BINDINGS.join('\n'));
  } else if (args.includes('--pending')) {
    console.log(JSON.stringify(await pendingPublishNames()));
  } else if (args.includes('--publish')) {
    console.log(JSON.stringify(PYTHON_PUBLISH_NAMES));
  } else {
    console.log(JSON.stringify(PYTHON_BINDING_NAMES));
  }
}

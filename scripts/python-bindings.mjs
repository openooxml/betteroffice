import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import {
  REGISTRY,
  FONT_PACK_NAMES,
  PYPI_DISTRIBUTIONS,
  PYTHON_BINDINGS,
  PYTHON_BINDING_NAMES,
  PYTHON_PUBLISH_NAMES,
  bindingName
} from './python-binding-names.mjs';

export { FONT_PACK_NAMES, PYPI_DISTRIBUTIONS, PYTHON_BINDINGS, PYTHON_BINDING_NAMES, PYTHON_PUBLISH_NAMES };

/** The crate version maturin stamps on the wheel; bindings version independently of the workspace. */
export function bindingVersion(path) {
  const manifest = readFileSync(new URL(`../${path}/Cargo.toml`, import.meta.url), 'utf8');
  const table = manifest.split(/^\[/m).find((section) => section.startsWith('package]'));
  const version = table?.match(/^\s*version\s*=\s*"([^"]+)"/m);
  if (!version) throw new Error(`${path}/Cargo.toml declares no literal [package] version`);
  return version[1];
}

/** The font pack version a wheel carries: its manifest's, which the Rust font catalog reads too. */
export function fontPackVersion(name) {
  const manifest = JSON.parse(
    readFileSync(new URL(`../packages/${name}/assets/manifest.json`, import.meta.url), 'utf8')
  );
  if (typeof manifest.version !== 'string') {
    throw new Error(`packages/${name}/assets/manifest.json declares no version`);
  }
  return manifest.version;
}

/** Font packs, then publish-enabled bindings, whose version PyPI does not serve yet. */
export async function pendingPublishNames({
  fetchImpl = fetch,
  attempts = 3,
  retryDelayMs = 2000
} = {}) {
  const pending = [];
  const candidates = [
    ...FONT_PACK_NAMES.map((name) => ({ name, version: () => fontPackVersion(name) })),
    ...REGISTRY.filter((entry) => entry.publish).map((entry) => ({
      name: bindingName(entry.path),
      version: () => bindingVersion(entry.path)
    }))
  ];
  for (const candidate of candidates) {
    const name = candidate.name;
    const project = `betteroffice-${name}`;
    const releases = await pypiReleases(project, { fetchImpl, attempts, retryDelayMs });
    const version = candidate.version();
    const files = releases?.[version];
    if (releases === null || files === undefined) {
      pending.push(name);
      continue;
    }
    if (!files.some((file) => file?.yanked !== true)) {
      throw new Error(
        `${project} ${version} exists on PyPI with no installable file; bump the binding version`
      );
    }
  }
  return pending;
}

/** null when the project does not exist yet; throws when PyPI cannot be read at all. */
async function pypiReleases(project, { fetchImpl, attempts, retryDelayMs }) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetchImpl(`https://pypi.org/pypi/${project}/json`, {
        headers: { Accept: 'application/json' },
        signal: AbortSignal.timeout(15_000)
      });
      if (response.status === 404) return null;
      if (!response.ok) throw new Error(`PyPI answered ${response.status} for ${project}`);
      return (await response.json()).releases ?? {};
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

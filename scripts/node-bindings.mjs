import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export const NODE_BINDINGS = ['bindings/node-docx', 'bindings/node-pptx', 'bindings/node-xlsx'];

function bindingName(path) {
  return path.replace('bindings/node-', '');
}

export const NODE_BINDING_NAMES = NODE_BINDINGS.map(bindingName);

function manifest(path) {
  return JSON.parse(readFileSync(new URL(`../${path}/package.json`, import.meta.url), 'utf8'));
}

export function bindingVersion(path) {
  return manifest(path).version;
}

export function platformPackageVersions() {
  return NODE_BINDINGS.flatMap((path) =>
    readdirSync(new URL(`../${path}/npm`, import.meta.url)).map((identity) => {
      const platform = manifest(`${path}/npm/${identity}`);
      return { name: platform.name, version: platform.version };
    })
  );
}

export async function pendingPublishNames({ fetchImpl = fetch } = {}) {
  const pending = [];
  for (const path of NODE_BINDINGS) {
    const { name, version } = manifest(path);
    const response = await fetchImpl(`https://registry.npmjs.org/${encodeURIComponent(name)}`, {
      headers: { Accept: 'application/vnd.npm.install-v1+json' },
      signal: AbortSignal.timeout(15_000)
    });
    if (response.status === 404) {
      pending.push(bindingName(path));
      continue;
    }
    if (!response.ok) throw new Error(`npm answered ${response.status} for ${name}`);
    const metadata = await response.json();
    if (!(version in (metadata.versions ?? {}))) pending.push(bindingName(path));
  }
  return pending;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [mode, ...rest] = process.argv.slice(2);
  if (rest.length > 0 || !['--paths', '--pending'].includes(mode)) {
    console.error('node-bindings.mjs: expected exactly one of --paths or --pending');
    process.exit(2);
  }
  console.log(
    mode === '--paths' ? NODE_BINDINGS.join('\n') : JSON.stringify(await pendingPublishNames())
  );
}

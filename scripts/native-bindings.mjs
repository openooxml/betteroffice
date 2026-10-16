import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export const NATIVE_PACKAGES = [
  'bindings/docx-native',
  'bindings/pptx-native',
  'bindings/xlsx-native'
];

function bindingName(path) {
  return path.replace(/^bindings\//, '').replace(/-native$/, '');
}

export const NATIVE_FORMATS = NATIVE_PACKAGES.map(bindingName);

function manifest(path) {
  return JSON.parse(readFileSync(new URL(`../${path}/package.json`, import.meta.url), 'utf8'));
}

export function bindingVersion(path) {
  return manifest(path).version;
}

export function validateNativeVersions() {
  for (const path of NATIVE_PACKAGES) {
    const root = manifest(path);
    const version = root.version;
    if (root.private !== true) {
      throw new Error(`${path} must be private`);
    }
    if (bindingVersion(`packages/${bindingName(path)}`) !== version) {
      throw new Error(`${path} must match its core package version`);
    }
    for (const identity of readdirSync(new URL(`../${path}/npm`, import.meta.url))) {
      const platform = manifest(`${path}/npm/${identity}`);
      if (platform.private === true) {
        throw new Error(`${path}/npm/${identity} must be public`);
      }
      if (platform.version !== version) {
        throw new Error(`${path}/npm/${identity} must match ${version}`);
      }
    }
  }
}

export function platformPackageVersions() {
  return NATIVE_PACKAGES.flatMap((path) =>
    readdirSync(new URL(`../${path}/npm`, import.meta.url)).map((identity) => {
      const platform = manifest(`${path}/npm/${identity}`);
      return { name: platform.name, version: platform.version };
    })
  );
}

export async function pendingPublishNames({ fetchImpl = fetch } = {}) {
  const pending = [];
  const missing = [];
  for (const path of NATIVE_PACKAGES) {
    const root = manifest(path);
    const packages = [
      root,
      ...readdirSync(new URL(`../${path}/npm`, import.meta.url)).map((identity) =>
        manifest(`${path}/npm/${identity}`)
      )
    ];
    for (const { name } of packages) {
      const response = await fetchImpl(`https://registry.npmjs.org/${encodeURIComponent(name)}`, {
        headers: { Accept: 'application/vnd.npm.install-v1+json' },
        signal: AbortSignal.timeout(15_000)
      });
      if (response.status === 404) {
        missing.push(name);
        continue;
      }
      if (!response.ok) throw new Error(`npm answered ${response.status} for ${name}`);
      if (name === root.name) {
        const metadata = await response.json();
        if (!(root.version in (metadata.versions ?? {}))) pending.push(bindingName(path));
      }
    }
  }
  if (missing.length > 0) {
    throw new Error(
      `Missing npm packages: ${missing.join(', ')}. They must be published by hand first ` +
        '(RELEASING.md, "Native packages"), because Trusted Publishing cannot create a package.'
    );
  }
  return pending;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [mode, ...rest] = process.argv.slice(2);
  if (rest.length > 0 || !['--paths', '--pending'].includes(mode)) {
    console.error('native-bindings.mjs: expected exactly one of --paths or --pending');
    process.exit(2);
  }
  console.log(
    mode === '--paths' ? NATIVE_PACKAGES.join('\n') : JSON.stringify(await pendingPublishNames())
  );
}

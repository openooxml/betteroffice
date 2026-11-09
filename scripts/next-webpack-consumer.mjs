// Builds a Next.js app with `next build --webpack` that imports every export of the editor
// packages, statically on one route and through next/dynamic on another. Run under node after
// `bun run build:packages`.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { exportSpecifiers, linkPackages, packageManifests } from './check-webpack-exports.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const FORMATS = ['docx', 'xlsx', 'pptx', 'vsdx'];
const EDITOR_PACKAGES = FORMATS.flatMap((format) => ['', '-react', '-i18n'].map((suffix) => `@betteroffice/${format}${suffix}`));

const demo = createRequire(join(ROOT, 'apps/demo/package.json'));
const packages = packageManifests().filter(({ manifest }) => EDITOR_PACKAGES.includes(manifest.name));
assert.deepEqual(packages.map(({ manifest }) => manifest.name).sort(), [...EDITOR_PACKAGES].sort());
const specifiers = packages
  .flatMap(({ manifest }) => exportSpecifiers(manifest))
  .filter((specifier) => !specifier.endsWith('/package.json'));
const modules = specifiers.filter((specifier) => !specifier.endsWith('.css'));
const stylesheets = specifiers.filter((specifier) => specifier.endsWith('.css'));

const page = `import Client from './client';
export const dynamic = 'force-dynamic';
export default function Page() {
  return <Client />;
}
`;

const files = {
  'package.json': JSON.stringify({ name: 'betteroffice-webpack-consumer', private: true }),
  'next.config.mjs': `export default {
  outputFileTracingRoot: ${JSON.stringify(ROOT)},
  webpack(config) {
    config.optimization.minimize = false;
    return config;
  },
};
`,
  'app/layout.jsx': `export default function RootLayout({ children }) {
  return <html lang="en"><body>{children}</body></html>;
}
`,
  'app/static/page.jsx': page,
  'app/static/client.jsx': `'use client';
${stylesheets.map((specifier) => `import ${JSON.stringify(specifier)};`).join('\n')}
${modules.map((specifier, index) => `import * as entry${index} from ${JSON.stringify(specifier)};`).join('\n')}
const entries = [${modules.map((_, index) => `entry${index}`).join(', ')}];
export default function Client() {
  return <p>{entries.length}</p>;
}
`,
  'app/dynamic/page.jsx': page,
  'app/dynamic/client.jsx': `'use client';
import dynamic from 'next/dynamic';
${modules
  .map(
    (specifier, index) =>
      `const Entry${index} = dynamic(() => import(${JSON.stringify(specifier)}).then(() => () => null), { ssr: false });`
  )
  .join('\n')}
export default function Client() {
  return <>${modules.map((_, index) => `<Entry${index} />`).join('')}</>;
}
`,
};

const directory = await mkdtemp(join(tmpdir(), 'next-webpack-consumer-'));
try {
  for (const [path, contents] of Object.entries(files)) {
    await mkdir(dirname(join(directory, path)), { recursive: true });
    await writeFile(join(directory, path), contents);
  }
  await linkPackages(directory, packages);
  for (const name of ['next', 'react', 'react-dom']) {
    await symlink(dirname(demo.resolve(`${name}/package.json`)), join(directory, 'node_modules', name), 'dir');
  }
  const build = spawnSync(process.execPath, [demo.resolve('next/dist/bin/next'), 'build', '--webpack', directory], {
    cwd: directory,
    stdio: 'inherit',
    env: { ...process.env, NEXT_TELEMETRY_DISABLED: '1' },
  });
  assert.equal(build.status, 0, `next build --webpack exited with ${build.status ?? build.signal}`);
  for (const route of ['static', 'dynamic']) {
    assert.ok(existsSync(join(directory, '.next/server/app', route, 'page.js')), `the ${route} route was not built`);
  }
  console.log(
    `next build --webpack bundles ${modules.length} entries and ${stylesheets.length} stylesheets of ` +
      `${packages.length} editor packages, statically and through next/dynamic.`
  );
} finally {
  await rm(directory, { recursive: true, force: true });
}

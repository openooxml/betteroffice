import { expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import type { BunPlugin } from 'bun';
import type { packageBoundary } from './package-boundary';

const packages = resolve(import.meta.dir, '../packages');
const reactPackages = readdirSync(packages, { withFileTypes: true })
  .filter((entry) => entry.isDirectory() && entry.name.endsWith('-react'))
  .map((entry) => entry.name);
type BoundaryPlugin = ReturnType<typeof packageBoundary>;

for (const name of reactPackages) {
  test(`${name} bundles without sibling package sources`, async () => {
    const directory = join(packages, name);
    const manifest = JSON.parse(readFileSync(join(directory, 'package.json'), 'utf8'));
    const { default: config } = await import(pathToFileURL(join(directory, 'tsup.config.ts')).href);
    const external: string[] = [...new Set<string>([
      ...Object.keys(manifest.dependencies ?? {}),
      ...Object.keys(manifest.peerDependencies ?? {}),
      ...config.external,
    ])];
    const isExternal = (path: string) => external.some((name) => path === name || path.startsWith(`${name}/`));
    const plugins: BunPlugin[] = config.esbuildPlugins.map((plugin: BoundaryPlugin): BunPlugin => ({
      name: plugin.name,
      setup(build) {
        plugin.setup({
          async resolve(path, options) {
            if (isExternal(path)) return { path, external: true };
            return { path: Bun.resolveSync(path, options.resolveDir), namespace: 'file' };
          },
          onResolve(options, callback) {
            build.onResolve(options, async (args) => {
              const result = await callback({
                path: args.path,
                importer: args.importer,
                resolveDir: args.importer ? dirname(args.importer) : directory,
                namespace: args.namespace,
                kind: 'import-statement',
              });
              if (result?.errors?.length) throw new Error(result.errors.map((error) => error.text).join('\n'));
              if (!result?.path) return;
              return { path: result.path, namespace: result.namespace, external: result.external };
            });
          },
        });
      },
    }));
    plugins.push({
      name: 'empty-assets',
      setup(build) {
        build.onLoad({ filter: /\.(css|svg|png|jpe?g|gif|webp|ico|woff2?|ttf|otf|wasm|txt|html)(\?.*)?$/ }, () => ({
          contents: '', loader: 'text',
        }));
      },
    });
    const result = await Bun.build({
      entrypoints: (Object.values(config.entry) as string[]).map((entry) => resolve(directory, entry)),
      format: 'esm',
      target: 'browser',
      external: external.flatMap((name) => [name, `${name}/*`]),
      plugins,
      write: false,
    });
    expect({ success: result.success, errors: result.logs.filter((log) => log.level === 'error').map((log) => log.message) })
      .toEqual({ success: true, errors: [] });
  });
}

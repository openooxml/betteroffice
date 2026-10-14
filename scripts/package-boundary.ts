import { basename, dirname, isAbsolute, relative, resolve, sep } from 'node:path';

type ResolveArgs = {
  path: string;
  importer: string;
  resolveDir: string;
  namespace: string;
  kind: 'entry-point' | 'import-statement' | 'require-call' | 'dynamic-import' |
    'require-resolve' | 'import-rule' | 'url-token' | 'composes-from';
  pluginData?: unknown;
};
type ResolveResult = {
  path?: string;
  namespace?: string;
  external?: boolean;
  errors?: { text: string }[];
  pluginData?: unknown;
};
type PluginBuild = {
  resolve(path: string, options: Omit<ResolveArgs, 'path'>): Promise<ResolveResult>;
  onResolve(options: { filter: RegExp }, callback: (args: ResolveArgs) => Promise<ResolveResult | undefined>): void;
};

export function packageBoundary(packageDirectory: string) {
  const directory = resolve(packageDirectory);
  const packages = dirname(directory);
  const root = dirname(packages);
  const resolving = Symbol('package-boundary');

  return {
    name: 'package-boundary',
    setup(build: PluginBuild) {
      build.onResolve({ filter: /.*/ }, async ({ path, ...args }) => {
        if (args.pluginData === resolving) return;
        const result = await build.resolve(path, { ...args, pluginData: resolving });
        if (result.pluginData === resolving) result.pluginData = args.pluginData;
        if (result.path && !result.external && result.namespace === 'file') {
          const target = relative(packages, result.path);
          const owner = target.split(sep)[0];
          if (!isAbsolute(target) && owner !== '..' && owner !== basename(directory) && target.includes(sep)) {
            throw new Error(
              `Package boundary violation: ${relative(root, args.importer) || '<entry>'} imports ${relative(root, result.path)} from another workspace package while bundling ${basename(directory)}.`
            );
          }
        }
        return result;
      });
    },
  };
}

// Rebuilds every workspace package whose inputs changed since it was last
// built. Apps resolve @betteroffice/* to dist/, so a missing or outdated dist
// silently serves a stale bundle; hashing inputs (rather than comparing mtimes)
// keeps the skip trustworthy even though the wasm build rewrites its output on
// every run.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readdir, readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const IGNORED_DIRS = new Set(['dist', 'node_modules']);
const STAMP = '.build-hash';

interface Target {
  name: string;
  dir: string;
  deps: string[];
}

export async function readTargets(): Promise<Target[]> {
  const packages = resolve(root, 'packages');
  const entries = await readdir(packages, { withFileTypes: true });
  const targets: Target[] = [];
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const dir = resolve(packages, entry.name);
    const packaged = await readFile(resolve(dir, 'package.json'), 'utf8').catch(() => null);
    if (!packaged) continue;
    const manifest = JSON.parse(packaged);
    if (!manifest.scripts?.build) continue;
    // A workspace peer still has to be built first: its types resolve to dist/.
    const deps = [manifest.dependencies, manifest.peerDependencies]
      .flatMap((group) => Object.entries<string>(group ?? {}))
      .filter(([, range]) => range.startsWith('workspace:'))
      .map(([name]) => name);
    targets.push({ name: manifest.name, dir, deps });
  }
  return targets;
}

export function order(targets: Target[]): Target[] {
  const byName = new Map(targets.map((target) => [target.name, target]));
  const sorted: Target[] = [];
  const seen = new Set<string>();
  const visit = (target: Target, stack: string[]) => {
    if (seen.has(target.name)) return;
    if (stack.includes(target.name)) {
      throw new Error(`workspace dependency cycle: ${[...stack, target.name].join(' -> ')}`);
    }
    for (const dep of target.deps) {
      const next = byName.get(dep);
      if (next) visit(next, [...stack, target.name]);
    }
    seen.add(target.name);
    sorted.push(target);
  };
  for (const target of targets) visit(target, []);
  return sorted;
}

async function hashInputs(dir: string): Promise<string> {
  const hash = createHash('sha256');
  const walk = async (current: string, prefix: string) => {
    const entries = await readdir(current, { withFileTypes: true });
    entries.sort((left, right) => (left.name < right.name ? -1 : 1));
    for (const entry of entries) {
      if (entry.name.startsWith('.') || entry.name.endsWith('.tsbuildinfo')) continue;
      if (entry.isDirectory()) {
        if (IGNORED_DIRS.has(entry.name)) continue;
        await walk(join(current, entry.name), `${prefix}${entry.name}/`);
        continue;
      }
      hash.update(`${prefix}${entry.name}\0`);
      hash.update(await readFile(join(current, entry.name)));
    }
  };
  await walk(dir, '');
  return hash.digest('hex');
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const targets = order(await readTargets());
  const keys = new Map<string, string>();

  async function keyOf(target: Target): Promise<string> {
    return createHash('sha256')
      .update(await hashInputs(target.dir))
      .update(target.deps.map((dep) => keys.get(dep) ?? '').join('\0'))
      .digest('hex');
  }

  for (const target of targets) {
    const key = await keyOf(target);
    keys.set(target.name, key);

    const stamp = resolve(target.dir, 'dist', STAMP);
    const current = await readFile(stamp, 'utf8').catch(() => null);
    if (current === key) {
      console.log(`${target.name}: up to date`);
      continue;
    }

    console.log(`${target.name}: building`);
    const build = spawnSync('bun', ['run', '--filter', target.name, 'build'], {
      cwd: root,
      stdio: 'inherit',
    });
    if (build.status !== 0) process.exit(build.status ?? 1);

    // The wasm builds vendor their output back into the hashed tree, so a key
    // taken beforehand never matches again; stamp the tree as it settled.
    const settled = await keyOf(target);
    keys.set(target.name, settled);
    await writeFile(stamp, settled);
  }
}

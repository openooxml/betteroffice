import { createHash } from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

export const BINDINGS = ['xlsx', 'docx', 'pptx'] as const;
const DEFAULT_VENV = path.join(
  os.homedir(),
  '.cache',
  'betteroffice',
  'e2e-venv'
);
const ROOT = path.resolve(import.meta.dir, '../..');

let sourceHash: string | undefined;

function bindingInputsHash(): string {
  if (sourceHash) return sourceHash;
  const tracked = Bun.spawnSync(
    [
      'git',
      'ls-files',
      '-z',
      '--cached',
      '--others',
      '--exclude-standard',
      '--',
      'Cargo.toml',
      'Cargo.lock',
      'crates',
      'bindings',
    ],
    { cwd: ROOT }
  );
  if (!tracked.success)
    throw new Error('cannot fingerprint Python binding sources');
  const hash = createHash('sha256');
  for (const name of tracked.stdout
    .toString()
    .split('\0')
    .filter(
      (name) =>
        name.endsWith('.rs') ||
        name.endsWith('.py') ||
        name.endsWith('pyproject.toml') ||
        name.endsWith('Cargo.toml') ||
        name.endsWith('Cargo.lock')
    )
    .sort()) {
    hash
      .update(name)
      .update('\0')
      .update(fs.readFileSync(path.join(ROOT, name)));
  }
  return (sourceHash = hash.digest('hex'));
}

export function venvDir(): string {
  return process.env.BETTEROFFICE_E2E_VENV ?? DEFAULT_VENV;
}

export function venvPython(dir = venvDir()): string {
  return process.platform === 'win32'
    ? path.join(dir, 'Scripts', 'python.exe')
    : path.join(dir, 'bin', 'python');
}

/** The interpreter with all three bindings importable, or the reason there is none. */
export function pythonWithBindings(): { python: string } | { missing: string } {
  const python = venvPython();
  if (!fs.existsSync(python))
    return {
      missing: `no interpreter at ${python}; run bun scripts/e2e/python-env.ts`,
    };
  const stamp = path.join(venvDir(), 'betteroffice-source.sha256');
  if (
    !fs.existsSync(stamp) ||
    fs.readFileSync(stamp, 'utf8').trim() !== bindingInputsHash()
  )
    return {
      missing:
        'Python bindings do not match this checkout; run bun scripts/e2e/python-env.ts',
    };
  const probe = Bun.spawnSync([
    python,
    '-c',
    BINDINGS.map((b) => `import betteroffice_${b}`).join('; '),
  ]);
  if (!probe.success)
    return {
      missing: `bindings not importable from ${python}: ${probe.stderr
        .toString()
        .trim()}`,
    };
  return { python };
}

function run(
  command: string[],
  env: Record<string, string | undefined> = {}
): void {
  console.log(`$ ${command.join(' ')}`);
  const result = Bun.spawnSync(command, {
    cwd: ROOT,
    env: { ...process.env, ...env },
    stdout: 'inherit',
    stderr: 'inherit',
  });
  if (!result.success)
    throw new Error(`${command[0]} exited with ${result.exitCode}`);
}

if (import.meta.main) {
  const dir = venvDir();
  if (!fs.existsSync(venvPython(dir)))
    run(['uv', 'venv', '--python', '3.13', dir]);
  const target =
    process.env.CARGO_TARGET_DIR ?? path.join(ROOT, 'bindings', 'target');
  for (const binding of BINDINGS) {
    run(
      [
        'maturin',
        'develop',
        '--release',
        '--locked',
        '--uv',
        '--manifest-path',
        path.join(ROOT, 'bindings', `python-${binding}`, 'Cargo.toml'),
      ],
      { VIRTUAL_ENV: dir, CARGO_TARGET_DIR: target }
    );
  }
  fs.writeFileSync(
    path.join(dir, 'betteroffice-source.sha256'),
    bindingInputsHash() + '\n'
  );
  const ready = pythonWithBindings();
  if ('missing' in ready) throw new Error(ready.missing);
  run([
    ready.python,
    '-c',
    BINDINGS.map(
      (b) =>
        `import betteroffice_${b} as m; print("betteroffice_${b}", getattr(m, "__version__", "?"))`
    ).join('; '),
  ]);
}

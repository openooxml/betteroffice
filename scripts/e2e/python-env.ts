/**
 * The Python side of the cross-SDK scenarios: a virtualenv holding release
 * builds of the three bindings. `bun scripts/e2e/python-env.ts` creates or
 * refreshes it; the suites find it through `BETTEROFFICE_E2E_VENV` or the
 * default location and skip the cross-SDK scenarios when it is missing.
 */

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

export const BINDINGS = ['xlsx', 'docx', 'pptx'] as const;
const DEFAULT_VENV = path.join(os.homedir(), '.cache', 'betteroffice', 'e2e-venv');
const ROOT = path.resolve(import.meta.dir, '../..');

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
  if (!fs.existsSync(python)) return { missing: `no interpreter at ${python}; run bun scripts/e2e/python-env.ts` };
  const probe = Bun.spawnSync([python, '-c', BINDINGS.map((b) => `import betteroffice_${b}`).join('; ')]);
  if (!probe.success) return { missing: `bindings not importable from ${python}: ${probe.stderr.toString().trim()}` };
  return { python };
}

function run(command: string[], env: Record<string, string | undefined> = {}): void {
  console.log(`$ ${command.join(' ')}`);
  const result = Bun.spawnSync(command, { cwd: ROOT, env: { ...process.env, ...env }, stdout: 'inherit', stderr: 'inherit' });
  if (!result.success) throw new Error(`${command[0]} exited with ${result.exitCode}`);
}

if (import.meta.main) {
  const dir = venvDir();
  if (!fs.existsSync(venvPython(dir))) run(['uv', 'venv', dir]);
  const target = process.env.CARGO_TARGET_DIR ?? path.join(ROOT, 'bindings', 'target');
  for (const binding of BINDINGS) {
    run(
      ['maturin', 'develop', '--release', '--uv', '--manifest-path', path.join(ROOT, 'bindings', `python-${binding}`, 'Cargo.toml')],
      { VIRTUAL_ENV: dir, CARGO_TARGET_DIR: target }
    );
  }
  const ready = pythonWithBindings();
  if ('missing' in ready) throw new Error(ready.missing);
  run([ready.python, '-c', BINDINGS.map((b) => `import betteroffice_${b} as m; print("betteroffice_${b}", getattr(m, "__version__", "?"))`).join('; ')]);
}

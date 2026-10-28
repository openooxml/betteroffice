import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const script = resolve(import.meta.dir, '../apps/native-viewer/macos/package-dmg.sh');
const APPS = ['BetterOffice Docs.app', 'BetterOffice Sheets.app', 'BetterOffice Slides.app'];

function packageWith(failures: string[]) {
  const dir = mkdtempSync(join(tmpdir(), 'package-dmg-'));
  try {
    const bin = join(dir, 'bin');
    mkdirSync(bin);
    writeFileSync(join(dir, 'failures'), failures.map((failure) => failure + '\n').join(''));
    const stubs: Record<string, string> = {
      ditto: 'cp -R "$1" "$2"',
      sleep: `echo "$1" >> "${dir}/sleeps"`,
      hdiutil: `echo "$1" >> "${dir}/calls"
ls "$TMPDIR" | grep '^betteroffice-dmg' >> "${dir}/scratch"
failure=$(head -n 1 "${dir}/failures")
sed -i.bak 1d "${dir}/failures"
if [ -n "$failure" ]; then echo "hdiutil: create failed - $failure" >&2; exit 1; fi
for last; do :; done
touch "$last"`,
    };
    for (const [name, body] of Object.entries(stubs)) {
      writeFileSync(join(bin, name), `#!/usr/bin/env bash\n${body}\n`);
      chmodSync(join(bin, name), 0o755);
    }
    for (const app of APPS) mkdirSync(join(dir, app));
    const dmg = join(dir, 'out', 'BetterOffice.dmg');
    const run = spawnSync('bash', [script, ...APPS.map((app) => join(dir, app)), dmg], {
      env: { ...process.env, PATH: `${bin}:${process.env.PATH}`, TMPDIR: dir },
      encoding: 'utf8',
    });
    const lines = (name: string) => {
      try { return readFileSync(join(dir, name), 'utf8').trim().split('\n').filter(Boolean); } catch { return []; }
    };
    return { status: run.status, stderr: run.stderr, calls: lines('calls'), sleeps: lines('sleeps'), scratch: lines('scratch') };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

test('package-dmg retries hdiutil create while the scratch volume is busy', () => {
  const run = packageWith(['Resource busy', 'Resource busy']);
  expect(run.status).toBe(0);
  expect(run.calls).toEqual(['create', 'create', 'create']);
  expect(run.sleeps).toEqual(['2', '4']);
});

test('package-dmg gives up after five busy attempts', () => {
  const run = packageWith(Array(5).fill('Resource busy'));
  expect(run.status).toBe(1);
  expect(run.calls).toHaveLength(5);
  expect(run.sleeps).toEqual(['2', '4', '8', '16']);
});

test('package-dmg fails at once on any other hdiutil error', () => {
  const run = packageWith(['No space left on device']);
  expect(run.status).toBe(1);
  expect(run.calls).toEqual(['create']);
  expect(run.sleeps).toEqual([]);
  expect(run.stderr).toContain('No space left on device');
});

test('package-dmg keeps hdiutil stderr inside its private scratch dir', () => {
  const run = packageWith(['Resource busy']);
  expect(run.status).toBe(0);
  expect(run.scratch).toHaveLength(2);
  expect(new Set(run.scratch).size).toBe(1);
});

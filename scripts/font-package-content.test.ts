import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const packages = resolve(import.meta.dir, '../packages');

function packedFiles(name: string): string[] {
  const output = mkdtempSync(join(tmpdir(), 'font-package-'));
  try {
    const archive = join(output, `${name}.tgz`);
    const pack = spawnSync(process.execPath, ['pm', 'pack', '--ignore-scripts', '--filename', archive], {
      cwd: join(packages, name), encoding: 'utf8',
    });
    expect({ status: pack.status, error: pack.error?.message, stderr: pack.stderr })
      .toMatchObject({ status: 0, error: undefined });
    const list = spawnSync('tar', ['-tzf', archive], { encoding: 'utf8' });
    expect({ status: list.status, error: list.error?.message, stderr: list.stderr })
      .toMatchObject({ status: 0, error: undefined });
    return list.stdout.trim().split('\n').map((file) => file.replace(/^package\//, '')).sort();
  } finally {
    rmSync(output, { recursive: true, force: true });
  }
}

test('@betteroffice/fonts publishes no raw TTFs and retains its licenses', () => {
  const files = packedFiles('fonts');
  expect(readdirSync(join(packages, 'fonts/assets')).filter((file) => file.endsWith('.ttf'))).toHaveLength(70);
  expect(files.filter((file) => /\.ttf$/i.test(file))).toEqual([]);
  expect(files).toContain('LICENSE');
  for (const file of readdirSync(join(packages, 'fonts/LICENSES'))) {
    expect(files).toContain(`LICENSES/${file}`);
  }
});

test('@betteroffice/fonts-latin publishes every packed face and its manifest', () => {
  const files = packedFiles('fonts-latin');
  const manifest = JSON.parse(readFileSync(join(packages, 'fonts-latin/assets/manifest.json'), 'utf8'));
  expect(files).toContain('assets/manifest.json');
  const frames = manifest.faces.map((face: { file: string }) => `assets/${face.file}`).sort();
  expect(frames).toHaveLength(70);
  expect(files.filter((file) => file.endsWith('.zst'))).toEqual(frames);
  expect(files.filter((file) => /\.ttf$/i.test(file))).toEqual([]);
});

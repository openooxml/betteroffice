import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const fonts = resolve(import.meta.dir, '../packages/fonts');

test('rebuilding the fonts package keeps the decoder and still clears stale output', () => {
  const out = mkdtempSync(join(tmpdir(), 'betteroffice-fonts-build-'));
  try {
    const decoder = join(out, 'decoder/fontpack.wasm');
    mkdirSync(join(out, 'decoder'));
    writeFileSync(decoder, 'decoder');
    writeFileSync(join(out, 'stale.js'), 'stale');

    const build = spawnSync('bun', ['x', 'tsup', '--out-dir', out], { cwd: fonts, encoding: 'utf8' });
    if (build.status !== 0) throw new Error(build.stdout + build.stderr);

    expect(readFileSync(decoder, 'utf8')).toBe('decoder');
    expect(existsSync(join(out, 'stale.js'))).toBe(false);
    expect(existsSync(join(out, 'progressive.js'))).toBe(true);
  } finally {
    rmSync(out, { recursive: true, force: true });
  }
}, 180_000);

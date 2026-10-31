import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { RUST_CRATES, RUST_PUBLISH_CRATES } from './rust-crates.mjs';

const root = new URL('../', import.meta.url);
const manifest = Bun.TOML.parse(readFileSync(new URL('Cargo.toml', root), 'utf8'));
const dependencies = manifest.workspace.dependencies as Record<
  string,
  { package?: string; path?: string; version?: string }
>;

test('the Rust release train includes every versioned workspace path dependency', () => {
  const expected = Object.entries(dependencies)
    .filter(([, dependency]) => dependency.path && dependency.version)
    .map(([key, dependency]) => `${key}:${dependency.package ?? key}`)
    .sort();

  expect(RUST_CRATES.map((crate) => `${crate.dependency}:${crate.name}`).sort()).toEqual(expected);
});

test('the Rust publish set respects each crate manifest', () => {
  for (const crate of RUST_CRATES) {
    const path = dependencies[crate.dependency]!.path!;
    const source = readFileSync(new URL(`${path}/Cargo.toml`, root), 'utf8');
    const manifest = Bun.TOML.parse(source);
    const published = RUST_PUBLISH_CRATES.some((entry) => entry.name === crate.name);

    expect(manifest.package.publish).toEqual(published ? ['crates-io'] : false);
  }
});

test('font preflight remains an unpublished path-only development dependency', () => {
  const edit = Bun.TOML.parse(readFileSync(new URL('crates/docx-edit/Cargo.toml', root), 'utf8'));
  const preflight = Bun.TOML.parse(readFileSync(new URL('crates/ooxml-font-preflight/Cargo.toml', root), 'utf8'));
  expect(edit['dev-dependencies']['ooxml-font-preflight']).toEqual({
    package: 'betteroffice-ooxml-font-preflight',
    path: '../ooxml-font-preflight',
  });
  expect(dependencies['ooxml-font-preflight']).toBeUndefined();
  expect(preflight.package.publish).toBe(false);
});

test('font byte crates publish before the manager and every engine', () => {
  const positions = new Map(RUST_PUBLISH_CRATES.map((crate, index) => [crate.name, index]));
  const manager = positions.get('betteroffice-fonts')!;
  for (const name of ['latin', 'intl', 'cjk', 'cjk-sc', 'cjk-tc', 'cjk-jp', 'cjk-kr', 'cjk-serif-sc']) {
    expect(positions.get(`betteroffice-fonts-${name}`)!).toBeLessThan(manager);
  }
  expect(manager).toBeLessThan(positions.get('betteroffice-xlsx')!);
  expect(manager).toBeLessThan(positions.get('betteroffice-docx')!);
  expect(manager).toBeLessThan(positions.get('betteroffice-pptx')!);
});

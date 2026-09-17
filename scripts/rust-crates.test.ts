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

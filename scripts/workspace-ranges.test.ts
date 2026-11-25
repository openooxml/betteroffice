import { afterEach, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { explicitWorkspaceRanges, restoreWorkspaceRanges } from './workspace-ranges.mjs';

const roots: string[] = [];
const dependency = { manifest: 'packages/b/package.json', section: 'dependencies', name: 'a', range: '^0' };
const peer = { manifest: 'packages/c/package.json', section: 'peerDependencies', name: 'a', range: '^0' };

afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true });
});

function writeManifest(root: string, manifest: string, json: object) {
  const file = join(root, manifest);
  mkdirSync(dirname(file), { recursive: true });
  writeFileSync(file, JSON.stringify(json, null, 2) + '\n');
  return file;
}

function workspace(peerRange = '^0') {
  const root = mkdtempSync(join(tmpdir(), 'workspace-ranges-'));
  roots.push(root);
  writeManifest(root, 'package.json', { workspaces: ['packages/*'] });
  const a = writeManifest(root, 'packages/a/package.json', { name: 'a', version: '0.4.2' });
  const b = writeManifest(root, dependency.manifest, {
    name: 'b',
    version: '1.0.0',
    config: { dependencies: { a: 'workspace:^0.5.0' } },
    dependencies: { a: 'workspace:^0' },
    devDependencies: { a: 'workspace:^', tilde: 'workspace:~', invalid: 'workspace:invalid' },
    optionalDependencies: { a: 'workspace:*' }
  });
  const c = writeManifest(root, peer.manifest, {
    name: 'c',
    version: '1.0.0',
    peerDependencies: { a: `workspace:${peerRange}` }
  });
  return { root, a, b, c };
}

test('records only explicit semver workspace ranges', () => {
  const { root } = workspace();
  expect(explicitWorkspaceRanges(root)).toEqual([dependency, peer]);
});

test('restores the in-range dependency value and leaves unchanged manifests untouched', () => {
  const { root, a, b, c } = workspace();
  const before = readFileSync(b, 'utf8');
  const peerSource = readFileSync(c, 'utf8');
  const past = new Date('2000-01-01T00:00:00Z');
  utimesSync(c, past, past);
  const peerMtime = statSync(c).mtimeMs;
  const recorded = explicitWorkspaceRanges(root);

  const bumped = before.replace('"1.0.0"', '"1.0.1"');
  writeFileSync(a, readFileSync(a, 'utf8').replace('0.4.2', '0.5.0'));
  writeFileSync(b, bumped.replace('"workspace:^0"', '"workspace:^0.5.0"'));

  expect(restoreWorkspaceRanges(recorded, root)).toEqual([dependency]);
  expect(readFileSync(b, 'utf8')).toBe(bumped);
  expect(readFileSync(c, 'utf8')).toBe(peerSource);
  expect(statSync(c).mtimeMs).toBe(peerMtime);
});

test('rejects an out-of-range version before writing any restorations', () => {
  const { root, a, b, c } = workspace('~0.4.0');
  const recorded = explicitWorkspaceRanges(root);
  const nextB = readFileSync(b, 'utf8').replace('"workspace:^0"', '"workspace:^0.5.0"');
  const nextC = readFileSync(c, 'utf8').replace('"workspace:~0.4.0"', '"workspace:~0.5.0"');
  writeFileSync(a, readFileSync(a, 'utf8').replace('0.4.2', '0.5.0'));
  writeFileSync(b, nextB);
  writeFileSync(c, nextC);

  expect(() => restoreWorkspaceRanges(recorded, root)).toThrow(
    'packages/c/package.json peerDependencies a declares workspace:~0.4.0, which does not include the new version 0.5.0.'
  );
  expect(readFileSync(b, 'utf8')).toBe(nextB);
  expect(readFileSync(c, 'utf8')).toBe(nextC);
});

test('rejects an out-of-range version for an unchanged workspace range', () => {
  const { root, a, b } = workspace();
  const recorded = explicitWorkspaceRanges(root);
  const before = readFileSync(b, 'utf8');
  const past = new Date('2000-01-01T00:00:00Z');
  utimesSync(b, past, past);
  const mtime = statSync(b).mtimeMs;
  writeFileSync(a, readFileSync(a, 'utf8').replace('0.4.2', '1.0.0'));

  expect(() => restoreWorkspaceRanges(recorded, root)).toThrow(
    'packages/b/package.json dependencies a declares workspace:^0, which does not include the new version 1.0.0.'
  );
  expect(readFileSync(b, 'utf8')).toBe(before);
  expect(statSync(b).mtimeMs).toBe(mtime);
});

test('rejects an invalid workspace range without rewriting the manifest', () => {
  const { root, b } = workspace();
  const recorded = explicitWorkspaceRanges(root);
  const invalid = Buffer.from(readFileSync(b, 'utf8').replace('"workspace:^0"', '"workspace:invalid"'));
  writeFileSync(b, invalid);

  expect(() => restoreWorkspaceRanges(recorded, root)).toThrow(
    'packages/b/package.json dependencies a changed to workspace:invalid, not a workspace range'
  );
  expect(readFileSync(b)).toEqual(invalid);
});

test('restores ranges while preserving four-space and tab indentation', () => {
  const { root, b, c } = workspace();
  const recorded = explicitWorkspaceRanges(root);
  const fourSpaces = JSON.stringify(JSON.parse(readFileSync(b, 'utf8')), null, 4) + '\n';
  const tabs = JSON.stringify(JSON.parse(readFileSync(c, 'utf8')), null, '\t');
  writeFileSync(b, fourSpaces.replace('"workspace:^0"', '"workspace:^0.4.2"'));
  writeFileSync(c, tabs.replace('"workspace:^0"', '"workspace:^0.4.2"'));

  expect(restoreWorkspaceRanges(recorded, root)).toEqual([dependency, peer]);
  expect(readFileSync(b, 'utf8')).toBe(fourSpaces);
  expect(readFileSync(c, 'utf8')).toBe(tabs);
});

test('records the explicit font pack ranges in docx-native', () => {
  const root = fileURLToPath(new URL('../', import.meta.url));
  const recorded = explicitWorkspaceRanges(root);
  for (const name of ['@betteroffice/fonts-cjk', '@betteroffice/fonts-intl', '@betteroffice/fonts-latin']) {
    expect(recorded).toContainEqual({
      manifest: 'bindings/docx-native/package.json',
      section: 'dependencies',
      name,
      range: '^0'
    });
  }
});

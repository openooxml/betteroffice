import { describe, expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PYTHON_BINDINGS } from './python-bindings.mjs';
import {
  PYPI_DISTRIBUTIONS,
  publishedCrates,
  publishedPackages,
  workspacePackages
} from './published-packages.mjs';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const CONTENT = join(ROOT, 'apps/docs/content/docs');
const DOCS_SITE = 'https://docs.betteroffice.dev/docs/';

function mdx(directory: string): string[] {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) return mdx(path);
    return entry.name.endsWith('.mdx') ? [readFileSync(path, 'utf8')] : [];
  });
}

const DOCS = mdx(CONTENT).join('\n');

describe('docs cover every published package', () => {
  test('names every published npm package', () => {
    expect(publishedPackages().filter((name) => !DOCS.includes(name))).toEqual([]);
  });

  test('names every published crate', () => {
    expect(publishedCrates().filter((name) => !DOCS.includes(name))).toEqual([]);
  });

  test('links every published PyPI distribution', () => {
    // A bare name would already match the same-named crate.
    const missing = PYPI_DISTRIBUTIONS.filter(
      (name) => !DOCS.includes(`https://pypi.org/project/${name}`)
    );
    expect(missing).toEqual([]);
  });
});

describe('docs name nothing that does not exist', () => {
  test('every npm package named is a package in this repository', () => {
    const real = new Set(workspacePackages());
    const claimed = [...DOCS.matchAll(/@betteroffice\/[a-z0-9-]+/g)].map((m) => m[0]);
    expect([...new Set(claimed)].filter((name) => !real.has(name))).toEqual([]);
  });

  test('every crate named is published, or is a PyPI distribution', () => {
    const real = new Set([...publishedCrates(), ...PYPI_DISTRIBUTIONS]);
    const claimed = [...DOCS.matchAll(/\bbetteroffice(?:-[a-z0-9]+)+\b/g)].map((m) => m[0]);
    expect([...new Set(claimed)].filter((name) => !real.has(name))).toEqual([]);
  });

  test('tells no one to install a distribution that is not on PyPI', () => {
    const real = new Set(PYPI_DISTRIBUTIONS);
    const claimed = [...DOCS.matchAll(/pip install (betteroffice-[a-z0-9-]+)/g)].map((m) => m[1]);
    expect(claimed.filter((name) => !real.has(name))).toEqual([]);
  });
});

describe('the PyPI sidebar link', () => {
  test('sends every binding at a docs page that exists', () => {
    for (const binding of PYTHON_BINDINGS) {
      const manifest = readFileSync(join(ROOT, binding, 'pyproject.toml'), 'utf8');
      const url = manifest.match(/^Documentation = "([^"]+)"/m)?.[1] ?? '';
      expect(url.startsWith(DOCS_SITE)).toBe(true);
      expect(existsSync(join(CONTENT, `${url.slice(DOCS_SITE.length)}.mdx`))).toBe(true);
    }
  });
});

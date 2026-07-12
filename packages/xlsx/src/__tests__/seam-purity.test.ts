/**
 * Guards the purity of the headless compute seam: the module closure reachable
 * from `src/headless.ts` must not import react, vue, the Canvas2D backend, the
 * wasm loader, or any DOM code. Keeping this graph plain-data is what lets the
 * seam be golden-tested and, later, crossed by a Rust/WASM implementation. If
 * this test fails, an impure import crept in — move the offending dependency to
 * a DOM-free home instead.
 */

import { describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

const ENTRY = resolve(import.meta.dir, '../headless.ts');

// import specifiers that must never appear anywhere in the seam's closure.
const FORBIDDEN = [/react/i, /vue/i, /canvas2d/, /wasm\/loader/, /\bdom\b/i];

// pull the module specifier out of every static import/export-from line.
function importSpecifiers(source: string): string[] {
  const specs: string[] = [];
  const re = /(?:import|export)[^'"]*?from\s*['"]([^'"]+)['"]/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(source)) !== null) specs.push(m[1]);
  return specs;
}

function resolveRelative(spec: string, fromFile: string): string | null {
  if (!spec.startsWith('.')) return null;
  const base = resolve(dirname(fromFile), spec);
  for (const candidate of [`${base}.ts`, `${base}/index.ts`]) {
    try {
      readFileSync(candidate, 'utf8');
      return candidate;
    } catch {
      // try next candidate
    }
  }
  return null;
}

// breadth-first walk of the relative-import graph rooted at the seam entry.
function collectClosure(entry: string): Map<string, string[]> {
  const seen = new Map<string, string[]>();
  const queue = [entry];
  while (queue.length > 0) {
    const file = queue.shift() as string;
    if (seen.has(file)) continue;
    const source = readFileSync(file, 'utf8');
    const specs = importSpecifiers(source);
    seen.set(file, specs);
    for (const spec of specs) {
      const next = resolveRelative(spec, file);
      if (next && !seen.has(next)) queue.push(next);
    }
  }
  return seen;
}

describe('headless seam purity', () => {
  it('imports no react, vue, canvas, wasm loader, or DOM code transitively', () => {
    const closure = collectClosure(ENTRY);
    const violations: string[] = [];
    for (const [file, specs] of closure) {
      for (const spec of specs) {
        if (FORBIDDEN.some((p) => p.test(spec))) {
          violations.push(`${file.split('/packages/')[1]} imports "${spec}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });

  it('reaches a non-trivial closure (sanity: the walk actually ran)', () => {
    const closure = collectClosure(ENTRY);
    expect(closure.size).toBeGreaterThan(3);
  });
});

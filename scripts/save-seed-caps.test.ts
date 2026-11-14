import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const NAMES = ['SAVE_SEED_MAX_PARAGRAPHS', 'SAVE_SEED_MAX_INFLATED_BYTES'] as const;

function caps(path: string): string[] {
  const source = readFileSync(join(root, path), 'utf8');
  return NAMES.map((name) => {
    const match = source.match(new RegExp(`^const ${name}(?::\\s*u32)?\\s*=\\s*([^;]+);`, 'm'));
    expect(match, `${path}: ${name}`).not.toBeNull();
    return match![1]!.replace(/\s+/g, '');
  });
}

test('the crate save seed caps match the product worker', () => {
  const product = caps('packages/docx/src/yrs/residentEngineWorker.ts');
  expect(product).toEqual(['4096', '16*1024*1024']);
  expect(caps('crates/docx-edit/src/identity.rs')).toEqual(product);
  expect(caps('crates/betteroffice-docx/src/document.rs')).toEqual(product);
});

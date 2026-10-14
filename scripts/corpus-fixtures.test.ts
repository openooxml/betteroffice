import { describe, expect, test } from 'bun:test';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

const ROOT = path.resolve(import.meta.dir, '..');
const FIXTURES = path.join(ROOT, 'crates/betteroffice-docx/tests/corpus/fixtures');

const GENERATED = [
  { builder: 'scripts/create-wordprocessingml-comprehensive.ts', fixture: 'wordprocessingml-comprehensive.docx' },
  { builder: 'scripts/create-demo-doc.ts', fixture: 'betteroffice-demo.docx' },
];

const digest = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

describe('generated corpus fixtures are provably current', () => {
  for (const { builder, fixture } of GENERATED) {
    test(`${fixture} equals the output of ${builder}`, () => {
      const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'corpus-fixture-'));
      const built = path.join(directory, fixture);
      try {
        execFileSync('bun', [path.join(ROOT, builder), built], { cwd: ROOT, stdio: 'pipe' });
        expect(digest(fs.readFileSync(built))).toBe(
          digest(fs.readFileSync(path.join(FIXTURES, fixture))),
        );
      } finally {
        fs.rmSync(directory, { recursive: true, force: true });
      }
    });
  }

  test('the demo fixture and the demo app ship the same bytes', () => {
    expect(digest(fs.readFileSync(path.join(FIXTURES, 'betteroffice-demo.docx')))).toBe(
      digest(fs.readFileSync(path.join(ROOT, 'apps/demo/public/betteroffice-demo.docx'))),
    );
  });

  test('the manifest hashes the bytes that are checked in', () => {
    const manifest = JSON.parse(
      fs.readFileSync(path.join(FIXTURES, '..', 'manifest.json'), 'utf8'),
    ) as { fixtures: { file: string; sha256: string }[] };
    for (const entry of manifest.fixtures) {
      expect(digest(fs.readFileSync(path.join(FIXTURES, entry.file)))).toBe(entry.sha256);
    }
  });
});

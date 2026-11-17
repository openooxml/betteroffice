import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test } from 'bun:test';

const TABLE = join(import.meta.dir, 'TABLE-COMMANDS.md');
const PROTOCOL = join(import.meta.dir, 'residentCommandProtocol.ts');
const CLASSES = new Set(['none', 'table-op', 'reads-payload']);

/** The editor command kinds: the keys of `DocxCommandPayloads`. */
function commandKinds(): string[] {
  const source = readFileSync(PROTOCOL, 'utf8');
  const body = /export interface DocxCommandPayloads \{\n([\s\S]*?)\n\}/.exec(source)![1]!;
  return [...body.matchAll(/^ {2}([a-zA-Z]+):/gm)].map((match) => match[1]!).sort();
}

const TARGETS = new Set(['selection', 'explicit', 'explicit-root', 'explicit-layout', 'named-limit-pr2', 'none']);

function rows(): Array<{ kind: string; reaches: string; mutatesBefore: string; target: string }> {
  return readFileSync(TABLE, 'utf8')
    .split('\n')
    .filter((line) => line.startsWith('| ') && !line.startsWith('| kind ') && !line.startsWith('|---'))
    .map((line) => {
      const [kind, reaches, , mutatesBefore, , target] = line.split('|').slice(1, -1).map((cell) => cell.trim());
      return { kind: kind!, reaches: reaches!, mutatesBefore: mutatesBefore!, target: (target ?? '').split(':')[0]!.trim() };
    });
}

test('every command kind says where its target comes from', () => {
  for (const row of rows()) expect(TARGETS.has(row.target), `${row.kind}: target ${row.target || 'missing'}`).toBe(true);
});

test('every editor command kind says whether it reaches a table operation, once', () => {
  const listed = rows().map((row) => row.kind);
  expect(listed.length).toBe(new Set(listed).size);
  expect([...listed].sort()).toEqual(commandKinds());
});

test('no command reaching a table operation writes the document before it', () => {
  for (const row of rows()) {
    expect(CLASSES.has(row.reaches), `${row.kind}: ${row.reaches}`).toBe(true);
    if (row.reaches === 'table-op') expect(row.mutatesBefore, row.kind).toBe('no');
  }
});

import { deepStrictEqual } from 'node:assert';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';

import { pptx } from './pptx';
import { xlsx } from './xlsx';

export interface StepResult {
  op: string;
  result: unknown;
}

export interface Script {
  fixture: string;
  steps: { op: string; args: Record<string, unknown>; targets?: ('ts' | 'rust' | 'python')[] }[];
  [key: string]: unknown;
}

export interface Target {
  operations: Record<string, (args: Record<string, unknown>) => unknown>;
  dispose(): void;
}

export interface Format {
  engineKind: string;
  open(bytes: Uint8Array, script: Script): Promise<Target>;
  versionPattern?: RegExp;
}

export const formats: Record<string, Format> = { xlsx, pptx };
const DEFAULT_VERSION_PATTERN = /^[0-9a-f]{32}-[0-9]+$/;
const root = new URL('../../', import.meta.url);

function expectedPath(format: string): URL {
  return new URL(`parity/${format}/expected.json`, root);
}

export function readExpected(format: string): StepResult[] {
  return JSON.parse(readFileSync(expectedPath(format), 'utf8'));
}

export function compareOrWrite(format: string, results: StepResult[], write = false): void {
  if (write) writeFileSync(expectedPath(format), `${JSON.stringify(results, null, 2)}\n`);
  deepStrictEqual(results, readExpected(format));
}

function mapStrings(value: unknown, transform: (value: string) => string): unknown {
  if (typeof value === 'string') return transform(value);
  if (Array.isArray(value)) return value.map((item) => mapStrings(item, transform));
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)
        .map(([key, item]) => [key, mapStrings(item, transform)])
    );
  }
  return value;
}

export async function run(format: string): Promise<StepResult[]> {
  if (!Object.hasOwn(formats, format)) throw new Error(`Unknown parity format ${format}`);
  const script = JSON.parse(
    readFileSync(new URL(`parity/${format}/script.json`, root), 'utf8')
  ) as Script;
  const target = await formats[format].open(readFileSync(new URL(script.fixture, root)), script);
  const versionPattern = formats[format].versionPattern ?? DEFAULT_VERSION_PATTERN;
  const versions: string[] = [];
  const normalize = (value: string): string => {
    if (versionPattern.exec(value)?.[0] !== value) return value;
    let index = versions.indexOf(value);
    if (index === -1) index = versions.push(value) - 1;
    return `<version:${index + 1}>`;
  };
  const resolve = (value: string): string => {
    const match = /^<version:([1-9][0-9]*)>$/.exec(value);
    if (!match || match[0] !== value) return value;
    const version = versions[Number(match[1]) - 1];
    if (version === undefined) throw new Error(`Unknown version reference ${value}`);
    return version;
  };
  const results: StepResult[] = [];
  try {
    for (const [index, { op, args }] of script.steps.entries()) {
      if (!Object.hasOwn(target.operations, op)) {
        throw new Error(`Unknown parity op ${op} at step ${index}`);
      }
      try {
        let result = await target.operations[op](mapStrings(args, resolve) as Record<string, unknown>);
        if (op === 'save') {
          if (!(result instanceof Uint8Array)) throw new TypeError('save must return bytes');
          result = {
            sha256: createHash('sha256').update(result).digest('hex'),
            byteLength: result.byteLength
          };
        }
        results.push({ op, result: mapStrings(result, normalize) });
      } catch (error) {
        throw new Error(`Parity step ${index} (${op}) failed`, { cause: error });
      }
    }
    return results;
  } finally {
    target.dispose();
  }
}

if (import.meta.main) {
  const format = process.argv[2];
  if (!format || format.startsWith('--')) {
    throw new Error('Usage: bun scripts/parity/harness.ts <format> [--write]');
  }
  const results = await run(format);
  compareOrWrite(
    format,
    results,
    process.argv.includes('--write') || process.env.BETTEROFFICE_PARITY_WRITE === '1'
  );
}

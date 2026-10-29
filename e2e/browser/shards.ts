import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const dirs = ['e2e/browser', 'e2e/browser/plugin-host'];

export function browserFiles(): string[] {
  return dirs
    .flatMap((dir) =>
      readdirSync(resolve(root, dir))
        .filter((name) => name.endsWith('.browser.ts'))
        .map((name) => `${dir}/${name}`),
    )
    .sort();
}

export function shardWeights(): Record<string, number> {
  return JSON.parse(readFileSync(resolve(import.meta.dirname, 'shard-weights.json'), 'utf8'));
}

export function assignShards(files: string[], weights: Record<string, number>, total: number): string[][] {
  const measured = files.map((file) => weights[file]).filter((weight) => weight !== undefined);
  const fallback = measured.length ? measured.reduce((sum, weight) => sum + weight, 0) / measured.length : 1;
  const weight = (file: string) => weights[file] ?? fallback;
  const shards = Array.from({ length: total }, () => ({ files: [] as string[], load: 0 }));
  for (const file of [...files].sort((a, b) => weight(b) - weight(a) || (a < b ? -1 : 1))) {
    const target = shards.reduce((best, shard) => (shard.load < best.load ? shard : best));
    target.files.push(file);
    target.load += weight(file);
  }
  return shards.map((shard) => shard.files.sort());
}

export function parseShard(spec: string): { current: number; total: number } {
  const match = /^([1-9]\d*)\/([1-9]\d*)$/.exec(spec);
  const current = Number(match?.[1]);
  const total = Number(match?.[2]);
  if (!match || current > total) throw new Error(`BETTEROFFICE_BROWSER_SHARD must be current/total, got ${spec}`);
  return { current, total };
}

export function shardTestMatch(spec = process.env.BETTEROFFICE_BROWSER_SHARD): RegExp[] | undefined {
  if (!spec) return undefined;
  const { current, total } = parseShard(spec);
  return assignShards(browserFiles(), shardWeights(), total)[current - 1].map(
    (file) => new RegExp(`${resolve(root, file).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`),
  );
}

import { S3Client } from 'bun';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { r2Options } from '../scripts/office-quality/r2.mjs';

export const FORMATS = ['docx', 'pptx', 'xlsx'] as const;
export const PREFIX = 'e2e';

export interface Upload {
  format: (typeof FORMATS)[number];
  key: string;
  file: string;
}

export interface Publication {
  sha: string;
  uploads: Upload[];
  manifest: {
    schema_version: 1;
    sha: string;
    recorded_at_utc: string;
    published_at_utc: string;
    run_number: number | null;
    formats: string[];
  };
}

function fail(message: string): never {
  throw new Error(`Refusing to publish e2e results: ${message}`);
}

/** Validates one complete, clean, green run of every format and names its bucket keys. */
export function planPublication(
  files: { file: string; run: unknown }[],
  publishedAt: string,
  expectedSha?: string,
  runNumber: number | null = null
): Publication {
  const seen = new Map<string, Upload>();
  let sha = '';
  let recordedAt = '';
  for (const { file, run } of files) {
    const result = run as Record<string, unknown> | null;
    const format = FORMATS.find((entry) => entry === result?.format);
    if (!result || !format || seen.has(format)) fail('expected one result file per format');
    if (result.schemaVersion !== 3) fail(`${format} has an unsupported schema`);
    if (!/^[a-f0-9]{40}$/.test(String(result.commit))) fail(`${format} has no commit`);
    if (sha && result.commit !== sha) fail('formats were recorded at different commits');
    if (result.dirty !== false) fail(`${format} was recorded from a dirty tree`);
    const scenarios = Array.isArray(result.scenarios) ? result.scenarios : [];
    if (!scenarios.length || scenarios.some((scenario) => scenario?.status !== 'passed'))
      fail(`${format} has failed or skipped scenarios`);
    if (typeof result.recordedAt !== 'string') fail(`${format} has no timestamp`);
    sha = String(result.commit);
    if (result.recordedAt > recordedAt) recordedAt = result.recordedAt;
    seen.set(format, { format, key: `${PREFIX}/${sha}/${format}.json`, file });
  }
  if (seen.size !== FORMATS.length) fail('expected one result file per format');
  if (expectedSha && sha !== expectedSha) fail(`results are for ${sha}, not ${expectedSha}`);
  return {
    sha,
    uploads: FORMATS.map((format) => seen.get(format)!),
    manifest: {
      schema_version: 1,
      sha,
      recorded_at_utc: recordedAt,
      published_at_utc: publishedAt,
      run_number: runNumber,
      formats: [...FORMATS],
    },
  };
}

/**
 * A re-run of an older commit uploads its files but must not move the pointer backwards.
 * Run numbers survive re-runs and only grow for newer pushes; record times do not.
 */
export function advances(current: unknown, next: Publication['manifest']): boolean {
  const pointer = current as { run_number?: unknown; recorded_at_utc?: unknown } | null;
  if (Number.isInteger(pointer?.run_number) && next.run_number !== null)
    return next.run_number >= (pointer!.run_number as number);
  const recorded = pointer?.recorded_at_utc;
  return typeof recorded !== 'string' || recorded <= next.recorded_at_utc;
}

if (import.meta.main) {
  const directory = resolve(process.argv[2] ?? '');
  const bucket = process.env.QUALITY_RENDER_BUCKET ?? '';
  if (!/^[a-z0-9][a-z0-9-]{1,62}$/.test(bucket))
    throw new Error('QUALITY_RENDER_BUCKET must name an R2 bucket');
  const files = await Promise.all(
    FORMATS.map(async (format) => {
      const file = resolve(directory, `${format}.json`);
      return { file, run: JSON.parse(await readFile(file, 'utf8')) as unknown };
    })
  );
  const run = Number(process.env.GITHUB_RUN_NUMBER);
  const plan = planPublication(files, new Date().toISOString(), process.env.GITHUB_SHA, Number.isInteger(run) ? run : null);
  const client = new S3Client(await r2Options(process.env));
  for (const upload of plan.uploads)
    await client.file(upload.key).write(Bun.file(upload.file), { type: 'application/json' });
  const pointer = client.file(`${PREFIX}/latest.json`);
  const current: unknown = (await pointer.exists()) ? await pointer.json().catch(() => null) : null;
  if (advances(current, plan.manifest)) {
    await pointer.write(JSON.stringify(plan.manifest, null, 2) + '\n', { type: 'application/json' });
    console.log(`Published e2e results for ${plan.sha}`);
  } else console.log(`Uploaded e2e results for ${plan.sha}; the pointer already names a newer run`);
}

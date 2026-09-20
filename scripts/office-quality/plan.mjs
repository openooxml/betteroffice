import { execFile } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { promisify } from 'node:util';
import { FORMATS } from './readme.mjs';
import { validateReferenceMetadata } from './reference.mjs';
import { CORPUS_ORIGIN, MAX_SAMPLES, selectSamples } from './samples.mjs';
import { download } from './download.mjs';

export const PLAN_SCHEMA_VERSION = 1;
export const METADATA_POOL_SIZE = 8;
const execute = promisify(execFile);

function isSha(value) {
  return typeof value === 'string' && /^[a-f0-9]{40}$/.test(value);
}

function isVersion(value) {
  return typeof value === 'string' && /^\d+\.\d+\.\d+(?:[-+][\w.-]+)?$/.test(value);
}

function validateSample(sample) {
  if (
    !sample ||
    typeof sample.id !== 'string' ||
    !/^[a-z0-9-]+$/.test(sample.id) ||
    !FORMATS.includes(sample.format) ||
    !sample.metadata ||
    sample.metadata.format !== sample.format ||
    (sample.metadata.id !== undefined && sample.metadata.id !== sample.id)
  )
    throw new Error('Invalid fidelity plan sample');
  validateReferenceMetadata(sample.metadata, sample.id);
}

export function createPlan({ source_sha, commit, versions, react_version, samples, docx_published_source_sha }) {
  if (!isSha(source_sha) || !isSha(commit))
    throw new Error('Invalid fidelity plan revision');
  if (
    !versions ||
    !FORMATS.every((format) => isVersion(versions[format])) ||
    Object.keys(versions).length !== FORMATS.length ||
    !isVersion(react_version)
  )
    throw new Error('Invalid fidelity plan versions');
  if (!Array.isArray(samples) || !samples.length || samples.length > MAX_SAMPLES)
    throw new Error(`Fidelity plan requires 1–${MAX_SAMPLES} samples`);
  const ids = new Set();
  for (const sample of samples) {
    validateSample(sample);
    if (ids.has(sample.id)) throw new Error('Fidelity plan has duplicate samples');
    ids.add(sample.id);
  }
  const formats = FORMATS.filter((format) =>
    samples.some((sample) => sample.format === format)
  );
  if (docx_published_source_sha !== undefined && !isSha(docx_published_source_sha))
    throw new Error('Invalid published DOCX source revision');
  return {
    schema_version: PLAN_SCHEMA_VERSION,
    source_sha,
    commit,
    versions,
    react_version,
    formats,
    samples,
    ...(docx_published_source_sha === undefined ? {} : { docx_published_source_sha }),
  };
}

export function validatePlan(plan) {
  if (plan?.schema_version !== PLAN_SCHEMA_VERSION || !Array.isArray(plan.formats))
    throw new Error('Invalid fidelity plan schema');
  const normalized = createPlan(plan);
  if (
    plan.formats.length !== normalized.formats.length ||
    plan.formats.some((format, index) => format !== normalized.formats[index])
  )
    throw new Error('Invalid fidelity plan formats');
  return normalized;
}

export function selectPlanSamples(plan, format = null) {
  const normalized = validatePlan(plan);
  if (format !== null && !FORMATS.includes(format))
    throw new Error('QUALITY_FORMAT must be docx, pptx, or xlsx');
  const formats = format ? [format] : normalized.formats;
  const samples = normalized.samples.filter((sample) => formats.includes(sample.format));
  if (!samples.length)
    throw new Error(
      `Fidelity plan has no samples for ${format ?? 'the selected formats'}`
    );
  return { formats, samples };
}

export async function mapPool(values, limit, task) {
  if (!Number.isInteger(limit) || limit < 1)
    throw new Error('Invalid fidelity metadata pool size');
  const results = new Array(values.length);
  let cursor = 0;
  await Promise.all(
    Array.from({ length: Math.min(limit, values.length) }, async () => {
      for (;;) {
        const index = cursor++;
        if (index >= values.length) return;
        results[index] = await task(values[index], index);
      }
    })
  );
  return results;
}

export async function preparePlan(environment, dependencies) {
  const {
    download: downloadMetadata,
    command,
    registry,
    poolSize = METADATA_POOL_SIZE,
  } = dependencies;
  const ids = await selectSamples(environment, downloadMetadata);
  const samples = await mapPool(ids, poolSize, async (id) => {
    const metadata = JSON.parse(
      await downloadMetadata(`${CORPUS_ORIGIN}/${id}/metadata.json`, 2 * 1024 * 1024)
    );
    if (!FORMATS.includes(metadata?.format))
      throw new Error(`Unsupported sample format: ${id}`);
    validateReferenceMetadata(metadata, id);
    return { id, format: metadata.format, metadata };
  });
  const [source_sha, commit, ...published] = await Promise.all([
    command('git', ['rev-parse', 'HEAD']),
    command('git', ['log', '-1', '--format=%H', '--', '.', ':!README.md']),
    ...FORMATS.map((format) => registry(`@betteroffice/${format}`)),
    registry('@betteroffice/docx-react'),
  ]);
  return createPlan({
    source_sha: source_sha.trim(),
    commit: commit.trim(),
    versions: Object.fromEntries(
      FORMATS.map((format, index) => [format, published[index].version])
    ),
    react_version: published.at(-1).version,
    samples,
    ...(samples.some((sample) => sample.format === 'docx')
      ? { docx_published_source_sha: published[0].gitHead ?? 'missing' }
      : {}),
  });
}

export async function writePlan(output, plan) {
  const directory = resolve(output);
  await mkdir(directory, { recursive: true });
  const path = resolve(directory, 'plan.json');
  await writeFile(path, JSON.stringify(validatePlan(plan), null, 2) + '\n', {
    flag: 'wx',
  });
  return path;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const command = async (program, args) => {
    const result = await execute(program, args);
    if (result.stderr) process.stderr.write(result.stderr);
    return result.stdout;
  };
  const registry = async (name) =>
    JSON.parse(
      await download(
        `https://registry.npmjs.org/${encodeURIComponent(name)}/latest`,
        2 * 1024 * 1024
      )
    );
  const plan = await preparePlan(process.env, { download, command, registry });
  const path = await writePlan(
    process.env.QUALITY_OUTPUT ?? '.source/office-quality/plan',
    plan
  );
  process.stdout.write(JSON.stringify({ path, formats: plan.formats }) + '\n');
}

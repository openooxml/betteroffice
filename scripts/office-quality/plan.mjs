import { execFile } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { promisify } from 'node:util';
import { FORMATS } from './readme.mjs';
import { validateReferenceMetadata } from './reference.mjs';
import { MAX_SAMPLES, corpusOrigin, selectSamples } from './samples.mjs';
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

export function createPlan({ source_sha, commit, versions, react_version, published_fonts, fonts_basis, samples, docx_published_source_sha, xlsx_published_source_sha, pptx_published_source_sha }) {
  if (!isSha(source_sha) || !isSha(commit))
    throw new Error('Invalid fidelity plan revision');
  if (
    !versions ||
    !FORMATS.every((format) => isVersion(versions[format])) ||
    Object.keys(versions).length !== FORMATS.length ||
    !isVersion(react_version) ||
    (published_fonts !== undefined && (typeof published_fonts !== 'object' || published_fonts === null ||
      !Object.entries(published_fonts).every(([format, version]) => FONT_FORMATS.includes(format) && isVersion(version)))) ||
    (fonts_basis !== undefined && (typeof fonts_basis !== 'object' || fonts_basis === null ||
      !Object.entries(fonts_basis).every(([format, basis]) => published_fonts?.[format] && ['release-commit', 'latest-before-release'].includes(basis))))
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
  if (xlsx_published_source_sha !== undefined && !isSha(xlsx_published_source_sha))
    throw new Error('Invalid published XLSX source revision');
  if (pptx_published_source_sha !== undefined && !isSha(pptx_published_source_sha))
    throw new Error('Invalid published PPTX source revision');
  return {
    schema_version: PLAN_SCHEMA_VERSION,
    source_sha,
    commit,
    versions,
    react_version,
    ...(published_fonts === undefined ? {} : { published_fonts }),
    ...(fonts_basis === undefined ? {} : { fonts_basis }),
    formats,
    samples,
    ...(docx_published_source_sha === undefined ? {} : { docx_published_source_sha }),
    ...(pptx_published_source_sha === undefined ? {} : { pptx_published_source_sha }),
    ...(xlsx_published_source_sha === undefined ? {} : { xlsx_published_source_sha }),
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

export const FONT_FORMATS = ['pptx', 'xlsx'];

const STABLE = /^\d+\.\d+\.\d+$/;
const released = (version) => STABLE.test(version) && version !== '0.0.0';
const semverOrder = (left, right) => {
  const a = left.split('.').map(Number);
  const b = right.split('.').map(Number);
  return a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
};

export function releasedFonts(fonts, releases) {
  const versions = Object.entries(fonts?.versions ?? {});
  const fonts_basis = {};
  const published_fonts = {};
  for (const [format, release] of Object.entries(releases)) {
    const label = `@betteroffice/${format}@${release?.version}`;
    const sameCommit = versions
      .filter(([version, manifest]) => released(version) && /^[a-f0-9]{40}$/.test(release?.gitHead ?? '') && manifest?.gitHead === release.gitHead)
      .map(([version]) => version);
    if (sameCommit.length > 1)
      throw new Error(`Expected at most one @betteroffice/fonts release from the release commit of ${label}, found ${sameCommit.length}`);
    if (sameCommit.length === 1) {
      published_fonts[format] = sameCommit[0];
      fonts_basis[format] = 'release-commit';
      continue;
    }
    const releasedAt = Date.parse(release?.time ?? '');
    const earlier = versions
      .map(([version]) => version)
      .filter((version) => released(version) && Date.parse(fonts?.time?.[version] ?? '') < releasedAt)
      .sort(semverOrder);
    if (!Number.isFinite(releasedAt) || !earlier.length)
      throw new Error(`No @betteroffice/fonts release from the release commit of ${label} or published before it`);
    published_fonts[format] = earlier.at(-1);
    fonts_basis[format] = 'latest-before-release';
  }
  return { published_fonts, fonts_basis };
}

export function registryLoaders(download) {
  const fetchJson = async (url, maximum) => JSON.parse(await download(url, maximum));
  return {
    registry: (name) => fetchJson(`https://registry.npmjs.org/${encodeURIComponent(name)}/latest`, 2 * 1024 * 1024),
    packument: (name) => fetchJson(`https://registry.npmjs.org/${encodeURIComponent(name)}`, 32 * 1024 * 1024),
  };
}

export async function preparePlan(environment, dependencies) {
  const {
    download: downloadMetadata,
    command,
    registry,
    packument,
    poolSize = METADATA_POOL_SIZE,
  } = dependencies;
  const origin = corpusOrigin(environment);
  const ids = await selectSamples(environment, downloadMetadata);
  const samples = await mapPool(ids, poolSize, async (id) => {
    const metadata = JSON.parse(
      await downloadMetadata(`${origin}/${id}/metadata.json`, 2 * 1024 * 1024)
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
  const fontFormats = FONT_FORMATS.filter((format) => samples.some((sample) => sample.format === format));
  const fontsChoice = fontFormats.length
    ? releasedFonts(await packument('@betteroffice/fonts'), Object.fromEntries(await Promise.all(fontFormats.map(async (format) => {
        const release = published[FORMATS.indexOf(format)];
        const document = await packument(`@betteroffice/${format}`);
        return [format, { ...release, time: document?.time?.[release.version] }];
      }))))
    : undefined;
  return createPlan({
    source_sha: source_sha.trim(),
    commit: commit.trim(),
    versions: Object.fromEntries(
      FORMATS.map((format, index) => [format, published[index].version])
    ),
    react_version: published.at(-1).version,
    ...(fontsChoice ? fontsChoice : {}),
    samples,
    ...(samples.some((sample) => sample.format === 'docx')
      ? { docx_published_source_sha: published[0].gitHead ?? 'missing' }
      : {}),
    ...(samples.some((sample) => sample.format === 'pptx')
      ? { pptx_published_source_sha: published[1].gitHead ?? 'missing' }
      : {}),
    ...(samples.some((sample) => sample.format === 'xlsx')
      ? { xlsx_published_source_sha: published[2].gitHead ?? 'missing' }
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
  const plan = await preparePlan(process.env, { download, command, ...registryLoaders(download) });
  const path = await writePlan(
    process.env.QUALITY_OUTPUT ?? '.source/office-quality/plan',
    plan
  );
  process.stdout.write(JSON.stringify({ path, formats: plan.formats }) + '\n');
}

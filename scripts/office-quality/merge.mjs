import {
  copyFile,
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rename,
  rm,
  writeFile,
} from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { FORMATS, renderSection } from './readme.mjs';
import { validatePlan as validateFidelityPlan } from './plan.mjs';
import { validateComparison } from './results.mjs';
import { digest, docxShards, mergeDocxBenchmarks } from './docx-benchmark.mjs';
import { mergeXlsxFidelity, xlsxFidelityShards } from './xlsx-fidelity.mjs';
import { mergeRoundtrips, roundtripShards } from './roundtrip.mjs';
import { mergePptxBenchmark } from './pptx-benchmark.mjs';
import { mergeXlsxBenchmarks, xlsxShards } from './xlsx-benchmark.mjs';

function fail(message) {
  throw new Error(`Invalid fidelity merge: ${message}`);
}

function object(value, name) {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    fail(`${name} must be an object`);
  return value;
}

function sameVersions(left, right) {
  return (
    left &&
    typeof left === 'object' &&
    !Array.isArray(left) &&
    Object.keys(left).length === FORMATS.length &&
    FORMATS.every((format) => left[format] === right[format])
  );
}

function failed(result, sample, channel) {
  if (
    result.status !== 'failed' ||
    !['capture', 'compare'].includes(result.stage) ||
    typeof result.error !== 'string' ||
    !result.error.trim() ||
    'penalized_ssim' in result
  )
    fail(`${sample.id} has an invalid failed ${channel} comparison`);
}

function comparison(result, sample, plan, channel) {
  object(result, `${sample.id} ${channel} comparison`);
  if (result.channel !== channel) fail(`${sample.id} has an invalid ${channel} channel`);
  if (
    (channel === 'published' && result.version !== plan.versions[sample.format]) ||
    (channel === 'published' && 'renderer_source_commit' in result) ||
    (channel === 'commit' && result.renderer_source_commit !== plan.commit) ||
    (channel === 'commit' && 'version' in result)
  )
    fail(`${sample.id} has mismatched ${channel} identity`);
  if (result.status === 'failed') {
    failed(result, sample, channel);
    return;
  }
  if (result.status !== undefined && result.status !== 'ok')
    fail(`${sample.id} has an invalid successful ${channel} status`);
  try {
    validateComparison(result);
  } catch {
    fail(`${sample.id} has an invalid successful ${channel} comparison`);
  }
  if (result.reference.sha256 !== sample.source_sha256)
    fail(`${sample.id} ${channel} reference hash does not match the planned source`);
  if (result.reference_pages !== sample.metadata.reference.pages)
    fail(`${sample.id} ${channel} reference pages do not match the planned reference`);
}

function normalizePlan(input) {
  const plan = validateFidelityPlan(input);
  return {
    ...plan,
    samples: plan.samples.map((sample) => ({
      ...sample,
      source_sha256: sample.metadata.source.sha256,
    })),
  };
}

export function mergeReports(planInput, reportInputs) {
  const plan = normalizePlan(planInput);
  if (!Array.isArray(reportInputs)) fail('reports must be an array');
  const byFormat = new Map();
  for (const input of reportInputs) {
    const report = object(input, 'report');
    const format = report.format;
    if (!plan.formats.includes(format) || byFormat.has(format))
      fail(`duplicate or unexpected report format: ${format}`);
    if (
      report.source_sha !== plan.source_sha ||
      report.commit !== plan.commit ||
      !sameVersions(report.versions, plan.versions) ||
      report.react_version !== plan.react_version
    )
      fail(`${format} report identity does not match plan`);
    if (!Array.isArray(report.samples)) fail(`${format} samples must be an array`);
    byFormat.set(format, report);
  }
  for (const format of plan.formats)
    if (!byFormat.has(format)) fail(`missing ${format} report`);

  const reportSamples = new Map();
  for (const [format, report] of byFormat) {
    const expected = plan.samples.filter((sample) => sample.format === format);
    const expectedIds = new Set(expected.map((sample) => sample.id));
    for (const sample of report.samples) {
      object(sample, `${format} sample`);
      if (
        typeof sample.id !== 'string' ||
        !expectedIds.has(sample.id) ||
        reportSamples.has(sample.id)
      )
        fail(`${format} report has an unknown or duplicate sample`);
      if (sample.format !== format) fail(`${sample.id} has a cross-format report`);
      const planned = expected.find((entry) => entry.id === sample.id);
      if (sample.source_sha256 !== planned.source_sha256)
        fail(`${sample.id} source hash does not match plan`);
      if (!Array.isArray(sample.comparisons) || sample.comparisons.length !== 2)
        fail(`${sample.id} must have exactly two comparisons`);
      const channels = new Set(sample.comparisons.map((entry) => entry?.channel));
      if (channels.size !== 2 || !channels.has('published') || !channels.has('commit'))
        fail(`${sample.id} must have one published and one commit comparison`);
      for (const channel of ['published', 'commit'])
        comparison(
          sample.comparisons.find((entry) => entry.channel === channel),
          planned,
          plan,
          channel
        );
      reportSamples.set(sample.id, sample);
    }
    if (report.samples.length !== expected.length)
      fail(`${format} report has missing samples`);
  }
  const samples = plan.samples.map((sample) => reportSamples.get(sample.id));
  return {
    source_sha: plan.source_sha,
    commit: plan.commit,
    versions: plan.versions,
    react_version: plan.react_version,
    samples,
  };
}

async function regular(path, label) {
  const info = await lstat(path).catch(() => null);
  if (!info?.isFile() || info.isSymbolicLink() || !info.size)
    fail(`${label} is missing, empty, or unsafe`);
}

async function directory(path, label) {
  const info = await lstat(path).catch(() => null);
  if (!info?.isDirectory() || info.isSymbolicLink())
    fail(`${label} is missing or unsafe`);
}

async function copyRenders(plan, report, parts, stage) {
  for (const sample of report.samples) {
    const comparison = sample.comparisons.find(
      (result) =>
        result.channel === 'commit' && result.renderer_source_commit === plan.commit
    );
    if (comparison.status === 'failed') continue;
    if (!Number.isInteger(comparison.actual_pages) || comparison.actual_pages < 1)
      continue;
    const root = resolve(
      parts,
      `visual-fidelity-renders-${sample.format}`,
      sample.id,
      'commit'
    );
    const destination = resolve(stage, sample.id, 'commit');
    await directory(
      resolve(parts, `visual-fidelity-renders-${sample.format}`),
      `${sample.format} render artifact`
    );
    await directory(
      resolve(parts, `visual-fidelity-renders-${sample.format}`, sample.id),
      `${sample.id} render directory`
    );
    await directory(root, `${sample.id} commit render directory`);
    await mkdir(destination, { recursive: true });
    for (let page = 1; page <= comparison.actual_pages; page += 1) {
      const name = `page_${String(page).padStart(4, '0')}.png`;
      const source = resolve(root, name);
      if (!source.startsWith(`${root}/`)) fail(`${sample.id} render path is unsafe`);
      await regular(source, `${sample.id}/${name}`);
      await copyFile(source, resolve(destination, name));
    }
  }
}

async function emptyOutput(output) {
  const info = await lstat(output).catch(() => null);
  if (!info) return;
  if (!info.isDirectory() || info.isSymbolicLink() || (await readdir(output)).length)
    fail('QUALITY_OUTPUT must be an empty regular directory or not exist');
  await rm(output, { recursive: true });
}

export async function mergeFromPaths({
  planPath,
  parts,
  output,
  requireRenders = false,
  requireDocxBenchmark = false,
  requirePptxBenchmark = false,
  requireXlsxBenchmark = false,
  requireXlsxFidelity = false,
  requireRoundtrip = false,
}) {
  const planBytes = await readFile(planPath);
  const plan = normalizePlan(JSON.parse(planBytes));
  await directory(resolve(parts), 'QUALITY_PARTS');
  const reports = await Promise.all(
    plan.formats.map(async (format) => {
      const path = resolve(parts, `visual-fidelity-report-${format}`, 'report.json');
      try {
        return JSON.parse(await readFile(path, 'utf8'));
      } catch {
        fail(`missing or invalid ${format} report`);
      }
    })
  );
  let report = mergeReports(plan, reports);
  if (requireDocxBenchmark && plan.formats.includes('docx')) {
    const benchmarks = await Promise.all(docxShards(plan).map(async (_, shard) =>
      JSON.parse(await readFile(resolve(parts, `docx-benchmark-report-${shard}`, 'report.json'), 'utf8'))));
    report = mergeDocxBenchmarks(plan, report, benchmarks, digest(planBytes));
  }
  if (requirePptxBenchmark && plan.formats.includes('pptx')) {
    const benchmark = JSON.parse(await readFile(resolve(parts, 'pptx-benchmark-report', 'report.json'), 'utf8'));
    report = mergePptxBenchmark(plan, report, benchmark, digest(planBytes));
  }
  if (requireXlsxBenchmark && xlsxShards(plan).length) {
    const benchmarks = await Promise.all(xlsxShards(plan).map(async (_, shard) =>
      JSON.parse(await readFile(resolve(parts, `xlsx-benchmark-report-${shard}`, 'report.json'), 'utf8'))));
    report = mergeXlsxBenchmarks(plan, report, benchmarks, digest(planBytes));
  }
  if (requireXlsxFidelity && plan.formats.includes('xlsx')) {
    const benchmarks = await Promise.all(xlsxFidelityShards(plan).map(async (_, shard) =>
      JSON.parse(await readFile(resolve(parts, `xlsx-fidelity-report-${shard}`, 'report.json'), 'utf8'))));
    report = mergeXlsxFidelity(plan, report, benchmarks, digest(planBytes));
  }
  if (requireRoundtrip) {
    const benchmarks = await Promise.all(roundtripShards(plan).map(async ({ format, shard }) =>
      JSON.parse(await readFile(resolve(parts, `roundtrip-report-${format}-${shard}`, 'report.json'), 'utf8'))));
    report = mergeRoundtrips(plan, report, benchmarks, digest(planBytes));
  }
  const section = renderSection(report);
  await emptyOutput(output);
  await mkdir(dirname(output), { recursive: true });
  const stage = await mkdtemp(join(dirname(output), `.${basename(output)}-`));
  try {
    if (requireRenders) await copyRenders(plan, report, resolve(parts), stage);
    await writeFile(join(stage, 'report.json'), JSON.stringify(report, null, 2) + '\n');
    await writeFile(join(stage, 'section.md'), section);
    await rename(stage, output);
  } catch (error) {
    await rm(stage, { recursive: true, force: true });
    throw error;
  }
  return report;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { QUALITY_PLAN, QUALITY_PARTS, QUALITY_OUTPUT, QUALITY_REQUIRE_RENDERS } =
    process.env;
  if (!QUALITY_PLAN || !QUALITY_PARTS || !QUALITY_OUTPUT)
    throw new Error('QUALITY_PLAN, QUALITY_PARTS, and QUALITY_OUTPUT are required');
  await mergeFromPaths({
    planPath: resolve(QUALITY_PLAN),
    parts: resolve(QUALITY_PARTS),
    output: resolve(QUALITY_OUTPUT),
    requireRenders: QUALITY_REQUIRE_RENDERS === 'true',
    requireDocxBenchmark: process.env.QUALITY_REQUIRE_DOCX_BENCHMARK === 'true',
    requirePptxBenchmark: process.env.QUALITY_REQUIRE_PPTX_BENCHMARK === 'true',
    requireXlsxBenchmark: process.env.QUALITY_REQUIRE_XLSX_BENCHMARK === 'true',
    requireXlsxFidelity: process.env.QUALITY_REQUIRE_XLSX_FIDELITY === 'true',
    requireRoundtrip: process.env.QUALITY_REQUIRE_ROUNDTRIP === 'true',
  });
}

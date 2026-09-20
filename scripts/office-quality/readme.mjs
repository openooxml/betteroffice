import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { validateComparison } from './results.mjs';
import { timingSummary } from './docx-benchmark.mjs';
import { roundtripSummary } from './roundtrip.mjs';
import { calculationSummary } from './xlsx-benchmark.mjs';

export const BEGIN = '<!-- BEGIN GENERATED VISUAL FIDELITY -->';
export const END = '<!-- END GENERATED VISUAL FIDELITY -->';
export const FORMATS = ['docx', 'pptx', 'xlsx'];

const LABEL_PX = 180;
const VALUE_PX = 230;

const link = (text, href) => `<a href="${href}">${text}</a>`;

function score(samples, channel, revision) {
  const scores = [];
  for (const sample of samples) {
    const results = sample.comparisons.filter(
      (result) =>
        result.channel === channel &&
        (channel === 'commit' ? result.renderer_source_commit : result.version) ===
          revision
    );
    if (results.length > 1) throw new Error('Duplicate sample/channel comparison');
    const result = results[0];
    if (!result) continue;
    if (result.status === 'failed') {
      if (
        !['capture', 'compare'].includes(result.stage) ||
        typeof result.error !== 'string' ||
        !result.error.trim() ||
        'penalized_ssim' in result
      )
        throw new Error('Invalid failed comparison');
    } else {
      if (result.status !== undefined && result.status !== 'ok')
        throw new Error('Invalid comparison status');
      validateComparison(result);
      scores.push(result);
    }
  }
  const paged = scores.filter(
    (result) =>
      Number.isInteger(result.reference_pages) && Number.isInteger(result.actual_pages)
  );
  return {
    count: scores.length,
    total: samples.length,
    value: scores.length
      ? (
          scores.reduce((sum, result) => sum + result.penalized_ssim, 0) / scores.length
        ).toFixed(4)
      : '—',
    exact: paged.filter((result) => result.reference_pages === result.actual_pages).length,
    paged: paged.length,
    pageError: paged.reduce(
      (sum, result) => sum + Math.abs(result.reference_pages - result.actual_pages),
      0
    ),
  };
}

export function renderSection(report) {
  if (!/^[a-f0-9]{40}$/.test(report.commit))
    throw new Error('Expected a full commit SHA');
  const short = report.commit.slice(0, 8);
  const commitLink = link(
    short,
    `https://github.com/openooxml/betteroffice/commit/${report.commit}`
  );
  const measure = (format) => {
    const version = report.versions[format];
    if (!/^\d+\.\d+\.\d+(?:[-+][\w.-]+)?$/.test(version))
      throw new Error(`Invalid ${format} version`);
    const samples = report.samples.filter((sample) => sample.format === format);
    return {
      version,
      versionLink: link(
        version,
        `https://www.npmjs.com/package/@betteroffice/${format}/v/${version}`
      ),
      published: score(samples, 'published', version),
      current: score(samples, 'commit', report.commit),
    };
  };
  const pages = (side) => (side.paged ? `${side.exact}/${side.paged}` : '—');
  const pageError = (side) => (side.paged ? String(side.pageError) : '—');
  for (const format of FORMATS) {
    const config = report[`${format}_benchmark`];
    if (config && !/^\d+(?:\.\d+){2,3}$/.test(config.libreoffice_version ?? ''))
      throw new Error('Invalid LibreOffice version');
    if (config && (config.published_version !== report.versions[format] ||
        config.builds?.commit?.source_sha !== report.source_sha))
      throw new Error('Benchmark results do not match the report revision');
  }
  const xlsxFidelity = report.xlsx_fidelity_benchmark;
  if (xlsxFidelity && (!/^\d+(?:\.\d+){2,3}$/.test(xlsxFidelity.libreoffice_version ?? '') ||
      xlsxFidelity.source_sha !== report.source_sha ||
      (report.xlsx_benchmark && report.xlsx_benchmark.libreoffice_version !== xlsxFidelity.libreoffice_version)))
    throw new Error('XLSX fidelity does not match the report revision or LibreOffice version');
  const timings = Object.fromEntries(['docx', 'pptx'].map(format =>
    [format, report[`${format}_benchmark`] ? timingSummary(report.samples, format) : null]));
  const calculation = report.xlsx_benchmark ? calculationSummary(report.samples) : null;
  const table = (format, extraRows = []) => {
    const { versionLink, published, current } = measure(format);
    const sides = [published, current];
    const headings = [`BetterOffice (${versionLink})`, `BetterOffice (${commitLink})`];
    const office = (format === 'xlsx' && xlsxFidelity) || report[`${format}_benchmark`];
    if (office) {
      sides.push(format === 'xlsx' && !xlsxFidelity ? { value: '—', not_measured: true } :
        score(report.samples.filter(sample => sample.format === format), 'libreoffice', office.libreoffice_version));
      headings.push(`LibreOffice (${office.libreoffice_version})`);
    }
    const rows = [
      ...extraRows.map(([label, cell]) => [label, ...sides.map(cell)]),
      ['SSIM', ...sides.map((side) => side.value)],
      ['Scored/total', ...sides.map((side) => side.not_measured ? '—' : `${side.count}/${side.total}`)],
    ];
    if (timings[format]) {
      const channels = ['published', 'commit', 'libreoffice'].map((channel) => timings[format].channels[channel]);
      rows.push(['Render time (avg)', ...channels.map((channel) => channel.mean_ms === null ? '—' : `${channel.mean_ms.toFixed(0)} ms`)]);
    }
    if (format === 'xlsx' && calculation) {
      const channels = ['published', 'commit', 'libreoffice'].map(channel => calculation.channels[channel]);
      rows.push(['Recalc time (avg)', ...channels.map(channel => channel.mean_ms === null ? '—' : `${channel.mean_ms.toFixed(0)} ms`)]);
    }
    const preservation = report.roundtrip_benchmark?.[format];
    if (preservation) {
      if (preservation.published_version !== report.versions[format] ||
          preservation.builds?.commit?.source_sha !== report.source_sha)
        throw new Error('Roundtrip results do not match the report revision');
      const summary = roundtripSummary(report.samples, format);
      for (const [label, field] of [['Parse success', 'parsed'], ['Lossless roundtrip', 'preserved']]) {
        const values = ['published', 'commit'].map(channel => summary[channel].total
          ? `${(100 * summary[channel][field] / summary[channel].total).toFixed(2)}%` : '—');
        rows.push([label, ...values, ...(office ? ['—'] : [])]);
      }
    }
    const value = (text) => `<td align="right">${text}</td>`;
    const head = (text) => `<th width="${VALUE_PX}" align="right">${text}</th>`;
    return [
      '<table>',
      `<tr><th width="${LABEL_PX}"></th>${headings.map(head).join('')}</tr>`,
      ...rows.map(([label, ...cells]) => `<tr><td>${label}</td>${cells.map(value).join('')}</tr>`),
      '</table>',
    ].join('\n');
  };
  const docxTable = table('docx', [
    ['Exact page counts', pages],
    ['Absolute page error', pageError],
  ]);
  return `${BEGIN}
## Benchmarks

*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml).*

### DOCX

Page agreement is reported on its own because a document either paginates as Word does or it does not; SSIM cannot express that.

${docxTable}


### PPTX

${table('pptx')}

### XLSX

${table('xlsx')}

For scoring, timing, coverage, and limitations, see the [benchmark methodology](scripts/office-quality/README.md).

${END}`;
}

export function updateReadme(readme, section) {
  if (!section.startsWith(BEGIN) || !section.endsWith(END))
    throw new Error('Invalid generated section');
  const begin = readme.indexOf(BEGIN);
  const end = readme.indexOf(END);
  if (begin < 0 !== end < 0 || (begin >= 0 && end < begin))
    throw new Error('Broken generated markers');
  const stripped =
    begin >= 0
      ? readme.slice(0, begin) + readme.slice(end + END.length)
      : readme.replace(/^## Visual fidelity\n[\s\S]*?(?=^## |$(?![\s\S]))/m, '');
  const packages = /^## Packages\r?\n/m.exec(stripped);
  if (!packages) throw new Error('README has no Packages section');
  const afterPackages = packages.index + packages[0].length;
  const nextSection = /^## /m.exec(stripped.slice(afterPackages));
  const position = nextSection ? afterPackages + nextSection.index : stripped.length;
  return `${stripped.slice(0, position).trimEnd()}\n\n${section}\n\n${stripped.slice(
    position
  )}`;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const [reportPath, readmePath = 'README.md'] = process.argv.slice(2);
  if (!reportPath) throw new Error('usage: readme.mjs report.json [README.md]');
  const section = renderSection(JSON.parse(await readFile(reportPath, 'utf8')));
  await writeFile(readmePath, updateReadme(await readFile(readmePath, 'utf8'), section));
}

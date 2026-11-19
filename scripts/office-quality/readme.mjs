import { lstat, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';
import { privateMode } from './private-step.mjs';
import { validateComparison } from './results.mjs';
import { timingSummary } from './docx-benchmark.mjs';
import { roundtripSummary } from './roundtrip.mjs';
import { calculationSummary } from './xlsx-benchmark.mjs';
import { failureCounts } from './failures.mjs';

export const BEGIN = '<!-- BEGIN GENERATED VISUAL FIDELITY -->';
export const END = '<!-- END GENERATED VISUAL FIDELITY -->';
export const LARGE_BEGIN = '<!-- BEGIN GENERATED LARGE SET -->';
export const LARGE_END = '<!-- END GENERATED LARGE SET -->';
const PAIRS = [[BEGIN, END], [LARGE_BEGIN, LARGE_END]];
export const FORMATS = ['docx', 'pptx', 'xlsx'];
export const RUN_KINDS = ['public', 'private'];

const RELEASE = String.raw`\d+\.\d+\.\d+`;
const FONT_PACKAGE = new RegExp(`^@betteroffice/fonts@${RELEASE}$`);
const FONT_CJK = new RegExp(`^@betteroffice/fonts-cjk@${RELEASE}$`);
const FONT_COPIES = ['local release copy', 'workspace'];
const FONT_SOURCE = new RegExp(`^@betteroffice/fonts@${RELEASE}(?: and @betteroffice/fonts-cjk@${RELEASE})? \\((?:${FONT_COPIES.join('|')})\\)$`);

const CHECKS = {
  count: [(value) => Number.isSafeInteger(value) && value >= 0, 'a non-negative integer'],
  mean: [(value) => value === null || Number.isFinite(value), 'a finite number or null'],
  fontPackage: [(value) => typeof value === 'string' && FONT_PACKAGE.test(value), 'an @betteroffice/fonts release'],
  fontCjk: [(value) => typeof value === 'string' && FONT_CJK.test(value), 'an @betteroffice/fonts-cjk release'],
  fontCopy: [(value) => FONT_COPIES.includes(value), `one of: ${FONT_COPIES.join(', ')}`],
  fontSource: [(value) => typeof value === 'string' && FONT_SOURCE.test(value), 'a font source of @betteroffice/fonts releases'],
  runKind: [(value) => RUN_KINDS.includes(value), `one of: ${RUN_KINDS.join(', ')}`],
  commit: [(value) => typeof value === 'string' && /^[a-f0-9]{40}$/.test(value), 'a full commit SHA'],
  version: [(value) => typeof value === 'string' && /^\d+\.\d+\.\d+(?:[-+][\w.-]+)?$/.test(value), 'a release version'],
  libreoffice: [(value) => value === null || (typeof value === 'string' && /^\d+(?:\.\d+){2,3}$/.test(value)), 'a LibreOffice version or null'],
};

/** Refuses a stored value the README would print, naming the field and never the value. */
export function checked(kind, value, field) {
  const [accepts, expected] = CHECKS[kind];
  if (!accepts(value)) throw new Error(`Invalid ${field}: expected ${expected}`);
  return value;
}

function checkFields(root, prefix, fields) {
  for (const [path, kind, optional = false] of fields) {
    const value = path.split('.').reduce((node, key) => node?.[key], root);
    if (!(optional && value === undefined)) checked(kind, value, `${prefix}${path}`);
  }
}

const isDate = (value) =>
  typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value) &&
  new Date(Date.parse(`${value}T00:00:00Z`) || 0).toISOString().slice(0, 10) === value;

const fixed = (value, digits) => {
  const text = value.toFixed(digits);
  return /^-0(?:\.0+)?$/.test(text) ? text.slice(1) : text;
};
const ms = (value) => (value === null ? '—' : `${fixed(value, 0)} ms`);
const percent = (part, total) => (total ? `${fixed(100 * part / total, 2)}%` : '—');
const ENGINES = ['published', 'commit', 'libreoffice'];

const LABEL_PX = 180;
const VALUE_PX = 230;

export const CHART_ROWS = ['SSIM', 'Exact page counts', 'Recalc accuracy'];
export const PERF_ROWS = ['Render time (avg)', 'Recalc time (avg)'];
export const CHART_URL = 'https://betteroffice.dev/api/charts/benchmarks.svg';

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
    mean: scores.length
      ? scores.reduce((sum, result) => sum + result.penalized_ssim, 0) / scores.length
      : null,
    value: scores.length
      ? fixed(scores.reduce((sum, result) => sum + result.penalized_ssim, 0) / scores.length, 4)
      : '—',
    exact: paged.filter((result) => result.reference_pages === result.actual_pages).length,
    paged: paged.length,
    pageError: paged.reduce(
      (sum, result) => sum + Math.abs(result.reference_pages - result.actual_pages),
      0
    ),
  };
}

const fontSourceLabel = (fonts) => {
  checkFields(fonts, 'fonts.', [['package', 'fontPackage'], ['cjk', 'fontCjk', true], ['copy', 'fontCopy']]);
  return `${fonts.package}${fonts.cjk ? ` and ${fonts.cjk}` : ''} (${fonts.copy})`;
};

const FONT_REASON = /^(?:http \d{3}|blocked|error)$/;
const fontReason = (reason) => (typeof reason === 'string' && FONT_REASON.test(reason) ? reason : 'other');

/** One arm's font source and face load counts across its captures. */
export function fontSummary(samples, channel) {
  const sources = new Set();
  const loaded = new Set();
  const failed = new Map();
  let advertised = 0;
  let baseLatin = 0;
  let captures = 0;
  let missing = 0;
  const unattributed = { loaded: 0, failed: 0 };
  const notMeasured = { timed_out: 0, errored: 0 };
  for (const sample of samples) {
    const fonts = sample.comparisons?.find((result) => result.channel === channel)?.fonts;
    if (!fonts) continue;
    if (typeof fonts.package === 'string') sources.add(fontSourceLabel(fonts));
    const capture = fonts.capture ?? 'ok';
    if (capture !== 'ok') {
      if (capture === 'timeout') notMeasured.timed_out++;
      else notMeasured.errored++;
      continue;
    }
    if (!Number.isInteger(fonts.advertised)) {
      if (fonts.measured) missing++;
      continue;
    }
    captures++;
    const count = (value, field) => {
      if (value === undefined) return 0;
      if (!Number.isSafeInteger(value) || value < 0)
        throw new Error(`the ${sample.format} measure job's ${channel} comparison has fonts.${field} that is not a non-negative integer`);
      return value;
    };
    advertised = Math.max(advertised, count(fonts.advertised, 'advertised'));
    baseLatin = Math.max(baseLatin, count(fonts.base_latin, 'base_latin'));
    unattributed.loaded += count(fonts.unattributed?.loaded, 'unattributed.loaded');
    unattributed.failed += count(fonts.unattributed?.failed, 'unattributed.failed');
    for (const face of fonts.loaded_faces ?? []) loaded.add(face);
    for (const [face, reason] of Object.entries(fonts.failed_faces ?? {}))
      if (!failed.has(face)) failed.set(face, reason);
  }
  for (const face of failed.keys()) loaded.delete(face);
  const reasons = {};
  for (const reason of [...failed.values()].map(fontReason)) reasons[reason] = (reasons[reason] ?? 0) + 1;
  return {
    sources: [...sources].sort(),
    captures,
    missing,
    advertised,
    base_latin: baseLatin,
    requested: loaded.size + failed.size,
    loaded: loaded.size,
    failed: failed.size,
    reasons,
    unattributed,
    not_measured: notMeasured,
  };
}

const notMeasuredText = ({ not_measured: { timed_out, errored } }) =>
  timed_out || errored ? `; captures not measured: ${timed_out} timed out, ${errored} errored` : '';

/** Why a DOCX arm fails its font check, by name; its completed captures are always checked. */
export function fontArmFailures(samples, channel) {
  const summary = fontSummary(samples, channel);
  const line = `DOCX ${fontArmLine(channel, summary)}`;
  const failures = [];
  if (!summary.captures && !summary.missing) failures.push(`DOCX ${channel} arm: nothing was measured${notMeasuredText(summary)}`);
  if (summary.failed) failures.push(line);
  const unrecorded = samples
    .filter((sample) => {
      const fonts = sample.comparisons.find((entry) => entry.channel === channel)?.fonts;
      return (fonts?.capture ?? 'ok') === 'ok' && !Number.isInteger(fonts?.advertised);
    })
    .map((sample) => sample.id);
  if (unrecorded.length)
    failures.push(`DOCX ${channel} arm: ${unrecorded.length} completed captures recorded no font loads: ${unrecorded.slice(0, 10).join(', ')}${unrecorded.length > 10 ? ', …' : ''}`);
  return { line, failures };
}

const reasonText = (reasons) =>
  Object.entries(reasons).sort(([a], [b]) => a.localeCompare(b)).map(([reason, total]) => `${reason} ${total}`).join(', ') || 'none';

const requestedDefinition = (summary) =>
  `M = faces the documents requested that the arm's provider offers, ${summary.base_latin ? 'base latin included' : 'base latin: none in this release'}`;

const unattributedText = ({ unattributed }) =>
  unattributed.loaded || unattributed.failed
    ? `; unattributed loads (not named in the advertised set, never a failure): ${unattributed.loaded} loaded, ${unattributed.failed} failed`
    : '';

const slicedText =
  '; sliced faces are observed only through their whole-face fallback at open; a chunk that fails or times out after open paints blank glyphs and is not observed';

export function fontArmLine(channel, summary) {
  return `${channel} arm: ${summary.failed} of ${summary.requested} faces did not load (${requestedDefinition(summary)}), by reason: ${reasonText(summary.reasons)}${summary.missing ? `; ${summary.missing} captures recorded no font loads` : ''}${slicedText}${unattributedText(summary)}${notMeasuredText(summary)}`;
}

function fontSourcesNote(sources) {
  const lines = [];
  for (const [channel, label] of [['published', 'Published'], ['commit', 'Commit']])
    if (sources[channel]?.length) lines.push(`- ${label} fonts: ${sources[channel].join('; ')}`);
  return lines.length ? `\n\n${lines.join('\n')}` : '';
}

const fontNote = (samples) =>
  fontSourcesNote(Object.fromEntries(['published', 'commit'].map((channel) => [channel, fontSummary(samples, channel).sources])));

const count = (value) => value.toLocaleString('en-US');
const NOUNS = { docx: ['DOCX documents', 'pages'], pptx: ['PPTX presentations', 'slides'], xlsx: ['XLSX workbooks', 'printed pages'] };

export function corpusSummary(plan) {
  return Object.fromEntries(FORMATS.filter((format) => plan.formats.includes(format)).map((format) => {
    const samples = plan.samples.filter((sample) => sample.format === format);
    return [format, { documents: samples.length, pages: samples.reduce((sum, sample) => sum + sample.metadata.reference.pages, 0) }];
  }));
}

function corpusNote(corpus) {
  if (!corpus) return '';
  const parts = FORMATS.filter((format) => corpus[format]).map((format) => {
    const { documents, pages } = corpus[format];
    if (!Number.isSafeInteger(documents) || !Number.isSafeInteger(pages) || documents < 1 || pages < documents)
      throw new Error('Invalid corpus summary');
    return `${count(documents)} ${NOUNS[format][0]} (${count(pages)} ${NOUNS[format][1]})`;
  });
  if (!parts.length) return '';
  const list = parts.length > 1 ? `${parts.slice(0, -1).join(', ')} and ${parts.at(-1)}` : parts[0];
  return `Corpus: ${list}, each with a Microsoft Office reference. Every page is rendered and scored against Office's PDF at 150 DPI with page-penalized SSIM, and each mean covers the documents that engine scored.${
    corpus.xlsx ? ' XLSX is compared per Excel print range, one range per page, so its page counts always match and are not reported.' : ''}`;
}

function recalcNote(calculation, libreoffice, missing, version, short) {
  const { workbooks, channels } = calculation;
  if (!workbooks) return '';
  const matching = (channel) => `${channels[channel].fully_matching}/${workbooks}`;
  return `Recalc rows: ${count(workbooks)} workbooks with formulas (${count(channels.commit.total)} formula results), each recalculated from cleared caches and compared with the values Excel saved. Workbooks fully matching: BetterOffice ${version} ${matching('published')}, BetterOffice at ${short} ${matching('commit')}, LibreOffice ${libreoffice} ${matching('libreoffice')}.${
    missing ? ` ${count(missing)} further ${missing === 1 ? 'workbook has' : 'workbooks have'} no Excel calculation reference and ${missing === 1 ? 'is' : 'are'} not counted.` : ''}`;
}

function timingNote(sequential, versions, short, sizeField = (format) => `timing_benchmark.${format}.sample_size`) {
  const formats = FORMATS.filter((format) => sequential[format]);
  if (!formats.length) return '';
  const sampled = Math.max(1, ...formats.map((format) => sequential[format].summary.total));
  for (const format of formats)
    if (sequential[format].sample_size < sampled)
      throw new Error(`Invalid ${sizeField(format)}: below the documents a format sampled`);
  const { sample_size: size, rounds } = sequential[formats[0]];
  const timed = formats.map((format) => {
    const { summary } = sequential[format];
    const engines = [
      `BetterOffice ${versions[format]} ${summary.channels.published.successful}`,
      `BetterOffice at ${short} ${summary.channels.commit.successful}`,
      `LibreOffice ${summary.channels.libreoffice.successful}`,
    ].join(', ');
    const unsettled = summary.notes_not_converged
      ? `, ${summary.notes_not_converged} of them with notes whose layout did not converge, timed as the product renders them`
      : '';
    return `${format.toUpperCase()} ${summary.common} of ${summary.total} sampled (timed per engine: ${engines}${unsettled})`;
  }).join('; ');
  return `Times: mean per document over a seeded sample of up to ${size} documents per format. A document counts only when every engine timed it in every round: ${timed}. They run one after another on one machine per format with the engine order rotated per document, ${rounds} rounds, each after a warm-up. Every run is a fresh process, so process start is included. DOCX and PPTX time a first-page render, XLSX a full recalculation and save.`;
}

function picture(url, alt) {
  return `<picture>
<source media="(prefers-color-scheme: dark)" srcset="${url}?theme=dark">
<img src="${url}" alt="${alt}" width="100%">
</picture>`;
}

export function chart(headline, perf = []) {
  const quality = headline.length ? `Scored against Microsoft Office, BetterOffice vs LibreOffice: ${headline.join(', ')}` : '';
  const time = perf.length
    ? `${quality ? 'average' : 'Average'} time per file, lower is better${quality ? '' : ', BetterOffice vs LibreOffice'}: ${perf
        .map(([format, label, ours, theirs]) => `${format.toUpperCase()} ${label} ${ours} vs ${theirs}`)
        .join(', ')}`
    : '';
  const parts = [quality, time].filter(Boolean);
  return picture(
    CHART_URL,
    parts.length ? `${parts.join('; ')}.` : 'BetterOffice and LibreOffice scored against Microsoft Office.'
  );
}

export function renderSection(report, { privateCorpus = false, date = new Date().toISOString().slice(0, 10) } = {}) {
  if (!/^[a-f0-9]{40}$/.test(report.commit))
    throw new Error('Expected a full commit SHA');
  if (privateCorpus && !isDate(date))
    throw new Error('Invalid date');
  const short = report.commit.slice(0, 8);
  const commitLink = link(
    short,
    `https://github.com/openooxml/betteroffice/commit/${report.commit}`
  );
  const measure = (format) => {
    const version = checked('version', report.versions[format], `${format} version`);
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
  const sequential = report.timing_benchmark ?? {};
  for (const format of FORMATS.filter((name) => sequential[name]))
    checkFields(sequential[format], `timing_benchmark.${format}.`, [
      ['rounds', 'count'], ['sample_size', 'count'], ['summary.total', 'count'], ['summary.common', 'count'],
      ['summary.notes_not_converged', 'count', true],
      ...ENGINES.flatMap((channel) => [[`summary.channels.${channel}.successful`, 'count'], [`summary.channels.${channel}.mean_ms`, 'mean']]),
    ]);
  const timings = Object.fromEntries(['docx', 'pptx'].map(format =>
    [format, privateCorpus ? sequential[format]?.summary ?? null
      : report[`${format}_benchmark`] ? timingSummary(report.samples, format) : null]));
  const calculation = report.xlsx_benchmark ? calculationSummary(report.samples) : null;
  const recalcTiming = privateCorpus ? sequential.xlsx?.summary ?? null : calculation;
  const headline = [];
  const perf = [];
  const table = (format, extraRows = []) => {
    const { versionLink, published, current } = measure(format);
    const sides = [published, current];
    const headings = [`BetterOffice (${versionLink})`, `BetterOffice (${commitLink})`];
    const preservation = report.roundtrip_benchmark?.[format];
    const fidelity = (format === 'xlsx' && xlsxFidelity) || (format !== 'xlsx' && report[`${format}_benchmark`]);
    const office = fidelity || report[`${format}_benchmark`] || preservation;
    if (preservation && (!/^\d+(?:\.\d+){2,3}$/.test(preservation.libreoffice_version ?? '') ||
        office.libreoffice_version !== preservation.libreoffice_version))
      throw new Error('Roundtrip LibreOffice version does not match the report');
    if (office) {
      sides.push(!fidelity ? { value: '—', not_measured: true } :
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
      rows.push(['Render time (avg)', ...channels.map((channel) => ms(channel.mean_ms))]);
    }
    if (format === 'xlsx' && calculation) {
      const channels = ['published', 'commit', 'libreoffice'].map(channel => calculation.channels[channel]);
      rows.push(['Recalc accuracy', ...channels.map(channel => percent(channel.correct, channel.total))]);
      if (recalcTiming) {
        const times = ['published', 'commit', 'libreoffice'].map(channel => recalcTiming.channels[channel]);
        rows.push(['Recalc time (avg)', ...times.map(channel => ms(channel.mean_ms))]);
      }
    }
    if (preservation) {
      if (preservation.published_version !== report.versions[format] ||
          preservation.builds?.commit?.source_sha !== report.source_sha)
        throw new Error('Roundtrip results do not match the report revision');
      const summary = roundtripSummary(report.samples, format);
      const values = ['published', 'commit', 'libreoffice'].map(channel => percent(summary[channel].parsed, summary[channel].total));
      rows.push(['Parse success', ...values]);
    }
    for (const label of CHART_ROWS) {
      const row = rows.find((candidate) => candidate[0] === label);
      const [ours, theirs] = row ? [row[2], row[3]] : [];
      if (/^[\d.,/%]+$/.test(ours ?? '') && /^[\d.,/%]+$/.test(theirs ?? ''))
        headline.push(`${format.toUpperCase()} ${label} ${ours} vs ${theirs}`);
    }
    for (const label of PERF_ROWS) {
      const row = rows.find((candidate) => candidate[0] === label);
      if (row && /^[\d.,]+ ms$/.test(row[2] ?? '') && /^[\d.,]+ ms$/.test(row[3] ?? ''))
        perf.push([format, label, row[2], row[3]]);
    }
    return tableHtml(headings, rows);
  };
  const docxTable = table('docx', [
    ['Exact page counts', pages],
    ['Absolute page error', pageError],
  ]);
  const pptxTable = table('pptx', [
    ['Exact page counts', pages],
    ['Absolute page error', pageError],
  ]);
  const xlsxTable = table('xlsx');
  const versionOf = (format) => report.versions[format];
  const formatSamples = (format) => report.samples.filter((sample) => sample.format === format);
  const notes = {
    corpus: corpusNote(report.corpus),
    recalc: calculation ? recalcNote(calculation, report.xlsx_benchmark.libreoffice_version,
      report.xlsx_benchmark.reference_coverage?.missing_reference?.length ?? 0, versionOf('xlsx'), short) : '',
    timing: privateCorpus ? timingNote(sequential, report.versions, short) : '',
  };
  const generated = privateCorpus
    ? largeGenerated(Object.fromEntries(FORMATS.map((format) => [format, formatSamples(format).length])), commitLink, date)
    : '; per-document results are on the [benchmarks site](https://benchmarks.betteroffice.dev)';
  const lead = [notes.corpus, privateCorpus ? '' : chart(headline, perf)].filter(Boolean).join('\n\n');
  return layout({
    large: privateCorpus,
    generated,
    lead,
    docx: `${docxTable}${fontNote(formatSamples('docx'))}`,
    pptx: `${pptxTable}${fontNote(formatSamples('pptx'))}`,
    xlsx: `${xlsxTable}${fontNote(formatSamples('xlsx'))}${note(notes.recalc)}${note(notes.timing)}`,
  });
}

function tableHtml(headings, rows) {
  const value = (text) => `<td align="right">${text}</td>`;
  const head = (text) => `<th width="${VALUE_PX}" align="right">${text}</th>`;
  return [
    '<table>',
    `<tr><th width="${LABEL_PX}"></th>${headings.map(head).join('')}</tr>`,
    ...rows.map(([label, ...cells]) => `<tr><td>${label}</td>${cells.map(value).join('')}</tr>`),
    '</table>',
  ].join('\n');
}

const note = (text) => (text ? `\n\n${text}` : '');

function largeGenerated(documents, commitLink, date) {
  const counts = FORMATS.map((format) => `${count(documents[format])} ${format.toUpperCase()}`);
  return ` from ${counts.slice(0, -1).join(', ')} and ${counts.at(-1)} documents at ${commitLink} on ${date}; aggregates only; documents are not public`;
}

function layout({ large, generated, lead, docx, pptx, xlsx }) {
  return `${large ? LARGE_BEGIN : BEGIN}
## Benchmarks${large ? ': large set' : ''}

*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml)${generated}.*

${lead ? `${lead}\n\n` : ''}<details>
<summary>Data</summary>

### DOCX

Page agreement is reported separately from SSIM.

${docx}


### PPTX

${pptx}

### XLSX

${xlsx}

</details>

For how scores, timings and coverage are measured, see the [benchmark methodology](scripts/office-quality/README.md).

${large ? LARGE_END : END}`;
}

function checkAggregate(aggregate) {
  checkFields(aggregate, '', [['run_kind', 'runKind', true], ['commit', 'commit']]);
  if (Object.keys(aggregate.formats).sort().join() !== FORMATS.join())
    throw new Error(`Invalid formats: expected exactly ${FORMATS.join(', ')}`);
  for (const format of FORMATS) {
    const entry = aggregate.formats[format];
    const fields = [['samples', 'count'], ['recalc_missing_reference', 'count', true],
      ['published_version', 'version'], ['libreoffice_version', 'libreoffice']];
    if (entry.corpus !== undefined) fields.push(['corpus.documents', 'count'], ['corpus.pages', 'count']);
    for (const side of ENGINES)
      if (side !== 'libreoffice' || entry.fidelity?.libreoffice !== null)
        fields.push([`fidelity.${side}.ssim`, 'mean'],
          ...['scored', 'total', 'exact_pages', 'paged', 'page_error'].map((key) => [`fidelity.${side}.${key}`, 'count']));
    if (entry.timing !== undefined)
      fields.push(...['rounds', 'sample', 'timed_by_every_engine'].map((key) => [`timing.${key}`, 'count']),
        ['timing.sample_size', 'count', true], ['timing.notes_not_converged', 'count', true],
        ...ENGINES.flatMap((channel) => [[`timing.channels.${channel}.successful`, 'count'], [`timing.channels.${channel}.mean_ms`, 'mean']]));
    if (entry.recalc !== undefined)
      fields.push(...ENGINES.flatMap((channel) => [
        ...['workbooks', 'correct', 'total', 'fully_matching'].map((key) => [`recalc.${channel}.${key}`, 'count']),
        [`recalc.${channel}.mean_ms`, 'mean'],
      ]));
    if (entry.roundtrip !== undefined)
      fields.push(...ENGINES.flatMap((channel) => ['parsed', 'preserved', 'total'].map((key) => [`roundtrip.${channel}.${key}`, 'count'])));
    checkFields(entry, `${format} `, fields);
    for (const channel of ['published', 'commit']) {
      const fonts = entry.fonts?.[channel];
      if (fonts === undefined) continue;
      const field = `${format} fonts.${channel}.sources`;
      if (!Array.isArray(fonts?.sources)) throw new Error(`Invalid ${field}: expected a list`);
      for (const source of fonts.sources) checked('fontSource', source, field);
    }
  }
}

/** The large set block from a run's aggregate.json, byte-identical to renderSection(report, { privateCorpus: true }). */
export function renderLargeSection(aggregate, { date = new Date().toISOString().slice(0, 10), timingSampleSize, missingReference } = {}) {
  if (aggregate?.schema_version !== 1 || !FORMATS.every((format) => aggregate.formats?.[format]))
    throw new Error('The large set block is written only from an aggregate.json of schema version 1');
  if (!isDate(date))
    throw new Error('Invalid date');
  for (const [name, value] of [['--timing-sample-size', timingSampleSize], ['--missing-reference', missingReference]])
    if (value !== undefined && !(Number.isSafeInteger(value) && value >= 0))
      throw new Error(`${name} must be a non-negative integer`);
  checkAggregate(aggregate);
  const { formats } = aggregate;
  const given = (value, argument, name, field, needed = true) => {
    if (value === undefined) {
      if (argument === undefined && needed) throw new Error(`aggregate.json has no ${field}; pass ${name}`);
      return argument;
    }
    if (argument !== undefined && argument !== value)
      throw new Error(`${name} ${argument} disagrees with ${field} ${value} in aggregate.json`);
    return value;
  };
  const short = aggregate.commit.slice(0, 8);
  const commitLink = link(short, `https://github.com/openooxml/betteroffice/commit/${aggregate.commit}`);
  const versions = Object.fromEntries(FORMATS.map((format) => [format, formats[format].published_version]));
  const table = (format) => {
    const entry = formats[format];
    const version = entry.published_version;
    const sides = [entry.fidelity.published, entry.fidelity.commit];
    const headings = [
      `BetterOffice (${link(version, `https://www.npmjs.com/package/@betteroffice/${format}/v/${version}`)})`,
      `BetterOffice (${commitLink})`,
    ];
    if (entry.libreoffice_version !== null) {
      sides.push(entry.fidelity.libreoffice);
      headings.push(`LibreOffice (${entry.libreoffice_version})`);
    }
    const rows = format === 'xlsx' ? [] : [
      ['Exact page counts', ...sides.map((side) => (side?.paged ? `${side.exact_pages}/${side.paged}` : '—'))],
      ['Absolute page error', ...sides.map((side) => (side?.paged ? String(side.page_error) : '—'))],
    ];
    rows.push(
      ['SSIM', ...sides.map((side) => (side && side.ssim !== null ? fixed(side.ssim, 4) : '—'))],
      ['Scored/total', ...sides.map((side) => (side ? `${side.scored}/${side.total}` : '—'))],
    );
    if (format !== 'xlsx' && entry.timing)
      rows.push(['Render time (avg)', ...ENGINES.map((channel) => ms(entry.timing.channels[channel].mean_ms))]);
    if (format === 'xlsx' && entry.recalc) {
      rows.push(['Recalc accuracy', ...ENGINES.map((channel) => percent(entry.recalc[channel].correct, entry.recalc[channel].total))]);
      if (entry.timing)
        rows.push(['Recalc time (avg)', ...ENGINES.map((channel) => ms(entry.timing.channels[channel].mean_ms))]);
    }
    if (entry.roundtrip)
      rows.push(['Parse success', ...ENGINES.map((channel) => percent(entry.roundtrip[channel].parsed, entry.roundtrip[channel].total))]);
    return `${tableHtml(headings, rows)}${fontSourcesNote({ published: entry.fonts?.published?.sources, commit: entry.fonts?.commit?.sources })}`;
  };
  const sequential = Object.fromEntries(FORMATS.filter((format) => formats[format].timing).map((format) => {
    const timing = formats[format].timing;
    return [format, {
      sample_size: given(timing.sample_size, timingSampleSize, '--timing-sample-size', `${format} timing.sample_size`),
      rounds: timing.rounds,
      summary: { total: timing.sample, common: timing.timed_by_every_engine, notes_not_converged: timing.notes_not_converged, channels: timing.channels },
    }];
  }));
  const { recalc } = formats.xlsx;
  const workbooks = recalc?.commit.workbooks ?? 0;
  const missing = given(formats.xlsx.recalc_missing_reference, missingReference, '--missing-reference', 'xlsx recalc_missing_reference', workbooks > 0);
  const recalcText = workbooks
    ? recalcNote({ workbooks, channels: recalc }, formats.xlsx.libreoffice_version, missing, versions.xlsx, short)
    : '';
  const docx = table('docx');
  const pptx = table('pptx');
  const xlsx = table('xlsx');
  return layout({
    large: true,
    generated: largeGenerated(Object.fromEntries(FORMATS.map((format) => [format, formats[format].samples])), commitLink, date),
    lead: corpusNote(Object.fromEntries(FORMATS.filter((format) => formats[format].corpus).map((format) => [format, formats[format].corpus]))),
    docx,
    pptx,
    xlsx: `${xlsx}${note(recalcText)}${note(timingNote(sequential, versions, short,
      (format) => (formats[format].timing.sample_size === undefined ? '--timing-sample-size' : `${format} timing.sample_size`)))}`,
  });
}

export function aggregateReport(report) {
  renderSection(report);
  const side = (result) => ({
    ssim: result.mean === null ? null : Number(fixed(result.mean, 4)),
    scored: result.count,
    total: result.total,
    exact_pages: result.exact,
    paged: result.paged,
    page_error: result.pageError,
  });
  const meanMs = (value) => (value === null ? null : Math.round(value));
  const formats = {};
  for (const format of FORMATS) {
    const samples = report.samples.filter((sample) => sample.format === format);
    const fidelity = format === 'xlsx' ? report.xlsx_fidelity_benchmark : report[`${format}_benchmark`];
    const office = fidelity ?? report[`${format}_benchmark`] ?? report.roundtrip_benchmark?.[format];
    const entry = {
      samples: samples.length,
      published_version: report.versions[format],
      libreoffice_version: office?.libreoffice_version ?? null,
      fidelity: {
        published: side(score(samples, 'published', report.versions[format])),
        commit: side(score(samples, 'commit', report.commit)),
        libreoffice: fidelity ? side(score(samples, 'libreoffice', fidelity.libreoffice_version)) : null,
      },
    };
    if (report.corpus?.[format]) entry.corpus = { ...report.corpus[format] };
    const fonts = Object.fromEntries(['published', 'commit']
      .map((channel) => [channel, fontSummary(samples, channel)])
      .filter(([, summary]) => summary.sources.length || summary.captures || summary.not_measured.timed_out || summary.not_measured.errored));
    if (Object.keys(fonts).length) entry.fonts = fonts;
    if (format !== 'xlsx' && report[`${format}_benchmark`]) {
      const timing = timingSummary(report.samples, format);
      entry.render_ms = Object.fromEntries(['published', 'commit', 'libreoffice'].map((channel) =>
        [channel, meanMs(timing.channels[channel].mean_ms)]));
    }
    if (format === 'xlsx' && report.xlsx_benchmark) {
      const calculation = calculationSummary(report.samples);
      entry.recalc = Object.fromEntries(['published', 'commit', 'libreoffice'].map((channel) => {
        const { correct, total, fully_matching, mean_ms } = calculation.channels[channel];
        return [channel, { correct, total, workbooks: calculation.workbooks, fully_matching, mean_ms: meanMs(mean_ms) }];
      }));
      entry.recalc_missing_reference = report.xlsx_benchmark.reference_coverage?.missing_reference?.length ?? 0;
    }
    if (report.roundtrip_benchmark?.[format]) {
      const summary = roundtripSummary(report.samples, format);
      entry.roundtrip = Object.fromEntries(['published', 'commit', 'libreoffice'].map((channel) => {
        const { parsed, preserved, total } = summary[channel];
        return [channel, { parsed, preserved, total }];
      }));
    }
    const sequential = report.timing_benchmark?.[format];
    if (sequential) {
      const { summary } = sequential;
      entry.timing = {
        method: sequential.method, sample_size: sequential.sample_size, sample: summary.total, timed_by_every_engine: summary.common, rounds: sequential.rounds,
        notes_not_converged: summary.notes_not_converged ?? 0,
        processors: sequential.runner.processors,
        channels: Object.fromEntries(Object.entries(summary.channels).map(([channel, value]) => [channel, {
          successful: value.successful, mean_ms: meanMs(value.mean_ms), median_ms: meanMs(value.median_ms),
          round_means_ms: value.round_means_ms.map(meanMs),
          round_spread: value.round_spread === null ? null : Number(value.round_spread.toFixed(4)),
        }])),
        ratio_to_libreoffice: summary.ratio_to_libreoffice,
      };
    }
    entry.failures = failureCounts(report, format);
    formats[format] = entry;
  }
  return { schema_version: 1, commit: report.commit, source_sha: report.source_sha ?? null, formats };
}

function blockRange(readme, begin, end) {
  const [start, stop] = [begin, end].map((marker) => {
    const at = readme.indexOf(marker);
    if (at >= 0 && (readme.includes(marker, at + 1) || (at > 0 && readme[at - 1] !== '\n') ||
        !/^(?:\r?\n|$)/.test(readme.slice(at + marker.length, at + marker.length + 2))))
      throw new Error('Broken generated markers');
    return at;
  });
  if ((start < 0) !== (stop < 0) || stop < start) throw new Error('Broken generated markers');
  return start < 0 ? null : [start, stop + end.length];
}

const occurrences = (text, marker) => text.split(marker).length - 1;

export function updateReadme(readme, section) {
  const kind = PAIRS.findIndex(([begin, end]) => section.startsWith(begin) && section.endsWith(end));
  if (kind < 0 || !PAIRS.every((pair, index) => pair.every((marker) => occurrences(section, marker) === (index === kind ? 1 : 0))))
    throw new Error('Invalid generated section');
  const [visual, large] = PAIRS.map(([begin, end]) => blockRange(readme, begin, end));
  if (visual && large && visual[1] > large[0] && large[1] > visual[0])
    throw new Error('Broken generated markers');
  const own = kind ? large : visual;
  const eol = readme.includes('\r\n') ? '\r\n' : '\n';
  const block = section.replace(/\r?\n/g, eol);
  if (own) return readme.slice(0, own[0]) + block + readme.slice(own[1]);
  if (kind && visual) return `${readme.slice(0, visual[1])}${eol}${eol}${block}${readme.slice(visual[1])}`;
  throw new Error('README has no generated block');
}

const RUN_PATH = /^\/(?!\.+\/)[\w.-]+\/(?!\.+\/)[\w.-]+\/actions\/runs\/\d+$/;

function isRunUrl(runUrl) {
  let url;
  try {
    url = new URL(runUrl);
  } catch {
    return false;
  }
  return url.href === runUrl && url.protocol === 'https:' && url.host === 'github.com' &&
    !url.username && !url.password && !url.search && !url.hash && RUN_PATH.test(url.pathname);
}

const isBadgeUrl = (text) => {
  try {
    const url = new URL(text.replaceAll('&amp;', '&'), 'https://invalid.');
    return url.hostname === 'betteroffice.dev' && decodeURIComponent(url.pathname).toLowerCase().replace(/\/+$/, '') === '/api/badge';
  } catch {
    return false;
  }
};

/** The section with a badge linking the workflow run that wrote it, dated by the given run date. */
export function withRunBadge(section, runUrl, date = new Date().toISOString().slice(0, 10)) {
  if (typeof runUrl !== 'string' || !isRunUrl(runUrl))
    throw new Error('--run-url must be https://github.com/<owner>/<repository>/actions/runs/<run id>');
  if ([...section.matchAll(/[\w+.-]*:?\/\/[^\s"'<>()]+/g)].some(([text]) => isBadgeUrl(text)))
    throw new Error('Invalid generated section');
  if (!isDate(date)) throw new Error('Invalid date');
  const generated = /^\*Generated by .*\*$/m.exec(section);
  if (!generated) throw new Error('Invalid generated section');
  const at = generated.index + generated[0].length;
  const badge = `<a href="${runUrl}"><img src="https://betteroffice.dev/api/badge?label=benchmark%20run&amp;message=${date}&amp;color=0a0a0a" alt="benchmark run ${date}"></a>`;
  return `${section.slice(0, at)}\n\n${badge}${section.slice(at)}`;
}

const USAGE = 'usage: readme.mjs --section section.md [README.md] | readme.mjs --as large aggregate.json [--timing-sample-size N] [--missing-reference N] [README.md]; each also takes [--date YYYY-MM-DD] [--run-url https://github.com/<owner>/<repository>/actions/runs/<run id>]';

async function readText(path) {
  if (!(await lstat(path)).isFile()) throw new Error(`${path} must be a regular file`);
  const bytes = await readFile(path);
  const text = bytes.toString('utf8');
  if (!Buffer.from(text, 'utf8').equals(bytes)) throw new Error(`${path} is not valid UTF-8`);
  return text;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      section: { type: 'string' }, as: { type: 'string' }, date: { type: 'string' },
      'timing-sample-size': { type: 'string' }, 'missing-reference': { type: 'string' }, 'run-url': { type: 'string' },
    },
  });
  const sectionRun = values.section !== undefined;
  const counts = ['timing-sample-size', 'missing-reference'].filter((name) => values[name] !== undefined);
  if (values.as === 'public') throw new Error('the public block is written only by --section');
  if ((values.as !== undefined && values.as !== 'large') || (sectionRun && (values.as !== undefined || counts.length || positionals.length > 1)) ||
      (!sectionRun && (!positionals.length || positionals.length > 2)))
    throw new Error(USAGE);
  if (!sectionRun && values.as === undefined) throw new Error('the positional command requires --as large');
  if (sectionRun && privateMode()) throw new Error('a private run writes the large set block only from aggregate.json');
  const date = values.date ?? new Date().toISOString().slice(0, 10);
  if (!isDate(date)) throw new Error('Invalid date');
  const argument = (name) => {
    if (values[name] === undefined) return undefined;
    if (!/^\d+$/.test(values[name]) || !Number.isSafeInteger(Number(values[name])))
      throw new Error(`--${name} must be a non-negative integer`);
    return Number(values[name]);
  };
  const timingSampleSize = argument('timing-sample-size');
  const missingReference = argument('missing-reference');
  const readmePath = (sectionRun ? positionals[0] : positionals[1]) ?? 'README.md';
  const readme = await readText(readmePath);
  let section;
  if (sectionRun) {
    section = await readText(values.section);
    if (!section.startsWith(BEGIN) || !section.endsWith(END))
      throw new Error('The generated section does not match this run');
  } else section = renderLargeSection(JSON.parse(await readText(positionals[0])), { date, timingSampleSize, missingReference });
  if (values['run-url'] !== undefined) section = withRunBadge(section, values['run-url'], date);
  await writeFile(readmePath, updateReadme(readme, section));
}

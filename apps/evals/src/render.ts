import { aggregate, formats, models, newestFirst, percent, repetitions, type Run, type Suite } from './stats';

export type Fixtures = { origin: string; provenance: string; fixtures: Record<string, { bytes: number; sha256: string }> };
export type SourceTask = { id: string; fixture: string };
export type SiteData = { suites: Suite[]; fixtures: Fixtures; tasks: SourceTask[]; commands: string[]; frame: string };
export const escape = (value: string | number) => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
const link = (href: string, label: string) => `<a href="${escape(href)}">${escape(label)}</a>`;
const section = (title: string, body: string) => `<section class="wrap rule py-[clamp(3.5rem,7vw,6rem)]"><h2 class="section-title">${escape(title)}</h2>${body}</section>`;
const repo = 'https://github.com/openooxml/betteroffice/blob/main/';
const count = (suite: Suite, format: string) => suite.tasks.filter(t => t.format === format).length;
const location = (base: string, path: string) => path === '/' ? base || '/' : base + path;

function intro(s: Suite) {
  const counts = formats(s).map(f => `${count(s, f)} ${f.toUpperCase()}`).join(', ');
  const repetition = repetitions(aggregate(s).runsPerTask);
  const modelCount = models(s).length;
  const arms = s.arms.map(a => a.id === 'generic' ? 'generic file tools' : a.id === 'betteroffice' ? 'the BetterOffice MCP server' : a.label).join(' vs ');
  return `${s.fixtures.count} ${escape(s.fixtures.kind)} fixtures (${counts}), ${repetition}, complete agent workflows ${modelCount === 1 ? 'with the same model' : `across ${modelCount} models`}: ${escape(arms)}.`;
}

function metric(s: Suite, model: string, arm: string, format?: string) {
  const a = s.arms.find(x => x.id === arm)!, stats = aggregate(s, { model, arm, format });
  if (!stats.n) return '';
  const label = format?.toUpperCase() ?? 'Overall';
  const every = stats.runsPerTask.some(n => n > 1)
    ? `<p>All runs pass: ${stats.allPass}/${stats.tasks} tasks (${percent(stats.allPassRate!)})</p>` : '';
  return `<article class="metric"><span class="eyebrow">${escape(label)} · ${escape(a.label)}</span><strong>${stats.passes}/${stats.n}</strong><div class="rate-bar" aria-hidden="true"><span style="width:${percent(stats.rate!)}"></span></div><p>${percent(stats.rate!)} pass rate</p><p>95% Wilson interval: ${percent(stats.interval[0])} to ${percent(stats.interval[1])}</p><p>n = ${stats.n} runs · ${stats.tasks} tasks</p><p>${repetitions(stats.runsPerTask)}</p>${every}</article>`;
}

function overview(s: Suite, suites: Suite[], base: string) {
  const headline = `<p class="mt-6 max-w-[760px] text-ink">${intro(s)}</p><p class="mt-5 max-w-[760px]">A task passes only if the requested edit is correct, the rest of the file is unchanged, and the input is untouched.</p><p class="mt-5 max-w-[760px] text-ink">&quot;Nothing else in the file moved&quot; matters because a correct-looking edit can still erase formatting, formula results, paragraph ids or speaker notes. Whole-package checks catch changes a quick visual check can miss.</p>`;
  let content = hero('Agent editing evals', s, headline);
  for (const model of models(s)) {
    content += section(`Results · ${model}`, `<div class="grid grid-cols-1 gap-px border border-line-soft bg-line-soft md:grid-cols-2">${[undefined, ...formats(s)].flatMap(f => s.arms.map(a => metric(s, model, a.id, f))).join('')}</div><p class="mt-6 text-sm text-ink">Intervals describe uncertainty in this small sample. They do not establish performance on other documents. ${link(location(base, '/method'), 'Method and limits')}</p>`);
  }
  content += section('Suites', `<table><caption class="sr-only">Every published evaluation suite</caption><thead><tr>${['Version', 'Date', 'Fixtures', 'Tasks', 'Runs per task', 'Models', 'Arms and pass rates'].map(t => `<th scope="col">${t}</th>`).join('')}</tr></thead><tbody>${suites.map(suite => {
    const rates = models(suite).flatMap(model => suite.arms.map(arm => {
      const stats = aggregate(suite, { model, arm: arm.id });
      return stats.n ? `<p>${escape(model)} · ${escape(arm.label)}: ${stats.passes}/${stats.n} (${percent(stats.rate!)})</p>` : '';
    })).join('');
    const cells = [link(`/suites/${encodeURIComponent(suite.suite)}`, suite.suite), escape(suite.date), escape(suite.fixtures.kind), suite.tasks.length, repetitions(aggregate(suite).runsPerTask), models(suite).map(escape).join('<br>'), rates];
    return `<tr>${cells.map((cell, i) => `<td data-label="${['Version', 'Date', 'Fixtures', 'Tasks', 'Runs per task', 'Models', 'Arms and pass rates'][i]}">${cell}</td>`).join('')}</tr>`;
  }).join('')}</tbody></table><p class="mt-6">${link(`/results/${encodeURIComponent(s.suite)}.json`, 'Download this suite as JSON')}</p>`);
  return content;
}

function hero(title: string, s: Suite, body = '') {
  return `<section class="wrap pt-12 pb-[clamp(3.5rem,7vw,6rem)] md:pt-20"><p class="eyebrow mb-6">${escape(s.suite)} · ${escape(s.date)} · ${escape(s.fixtures.kind)}</p><h1 class="max-w-[820px] text-[clamp(2.5rem,6vw,4.5rem)] leading-[0.95] font-semibold tracking-[-0.045em] text-balance">${escape(title)}</h1>${body}</section>`;
}

function resultCell(runs: Run[]) {
  if (!runs.length) return '<p>No run recorded.</p>';
  const passes = runs.filter(r => r.pass).length;
  const verdict = runs.length === 1 ? (passes ? 'Pass' : 'Fail') : `${passes}/${runs.length} pass`;
  return `<strong class="${passes === runs.length ? 'pass' : ''}">${verdict}</strong>${runs.map(r => `<div class="mt-3"><p class="eyebrow">Run ${r.run} · ${escape(r.date)}</p><p>${r.tool_calls} tool calls · ${r.seconds.toFixed(2)} seconds</p>${r.reason ? `<p>${escape(r.reason)} <span class="text-dim">Cause: ${escape(r.cause ?? 'unclassified')}.</span></p>` : ''}<details><summary>Checker message</summary><p>${escape(r.check)}</p></details></div>`).join('')}`;
}

function footnotes(s: Suite) {
  return `<div class="prose text-sm text-ink"><p>Tool calls count different things across arms.</p><ul>${s.arms.map(a => `<li>${escape(a.label)}: ${escape(a.tool_calls)}</li>`).join('')}</ul><p>Times come from runs on different days and are not comparable across arms.</p></div>`;
}

function taskPage(s: Suite, format: string) {
  let body = hero(`${format.toUpperCase()} tasks`, s);
  for (const model of models(s)) {
    const rows = s.tasks.filter(t => t.format === format).map(t => `<tr id="${escape(model + '-' + t.id)}"><td><a id="${escape(t.id)}${model === models(s)[0] ? '' : '-' + escape(model)}"></a><strong>${escape(t.title)}</strong><p class="eyebrow">${escape(t.id)}</p><details><summary>Task prompt</summary><p>${escape(t.prompt)}</p></details></td>${s.arms.map(a => `<td data-label="${escape(a.label)}">${resultCell(s.runs.filter(r => r.task === t.id && r.model === model && r.arm === a.id).sort((a, b) => a.run - b.run))}</td>`).join('')}</tr>`).join('');
    body += section(model, `<table class="task-table"><caption class="sr-only">${escape(format.toUpperCase())} results for ${escape(model)}</caption><thead><tr><th scope="col">Task</th>${s.arms.map(a => `<th scope="col">${escape(a.label)}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table>`);
  }
  return body + section('Reading these results', footnotes(s));
}

function failures(s: Suite, base: string) {
  let content = hero('How generic tools break files', s, '<p class="mt-6 max-w-[760px] text-ink">Each failed run below links to its task. Library means the library API or save changed the file. Model means the editing code made the error.</p>');
  const groups = new Map<string, Run[]>();
  for (const r of s.runs.filter(r => !r.pass)) {
    const format = s.tasks.find(t => t.id === r.task)!.format;
    const library = r.arm === 'generic' ? ({ docx: 'python-docx', xlsx: 'openpyxl', pptx: 'python-pptx' } as Record<string, string>)[format] ?? 'Other tools' : s.arms.find(a => a.id === r.arm)!.label;
    const key = `${r.cause ?? 'Unclassified'} · ${library}`;
    const group = groups.get(key) ?? []; group.push(r); groups.set(key, group);
  }
  for (const [label, runs] of [...groups].sort(([a], [b]) => a.localeCompare(b))) {
    content += section(label, `<ul class="grid gap-6">${runs.map(r => {
      const task = s.tasks.find(t => t.id === r.task)!;
      return `<li>${link(location(base, '/' + task.format) + '#' + encodeURIComponent(r.model + '-' + r.task), task.title)}<p class="eyebrow mt-2">${escape(r.model)} · ${escape(r.arm)} · run ${r.run}</p><p class="mt-2 max-w-[760px]">${escape(r.reason ?? r.check)}</p></li>`;
    }).join('')}</ul>`);
  }
  return content;
}

function method(s: Suite) {
  const v = s.checker_validation;
  return hero('Method and limits', s) + section('Complete agent workflows', `<div class="prose"><p>${intro(s)}</p><p>Models: ${models(s).map(escape).join(', ')} via ${escape(s.harness)}. Each task starts in a fresh temporary directory containing only its input. User configuration, rules and persisted sessions are skipped. Runs have a time limit and no approval prompts.</p><h3>Generic file tools</h3><p>The agent uses shell commands and Python with python-docx, openpyxl and python-pptx in a workspace-write sandbox, with no network or MCP servers. Logged commands are audited for access to the repository or task metadata. Reads are not isolated by the filesystem sandbox.</p><h3>BetterOffice MCP</h3><p>The agent runs in a read-only sandbox. The required BetterOffice MCP server is rooted at the task directory and is the only writer. Other document access tools are disabled. Code mode can call tools but has no file or network access. Server startup and required tools are checked before tasks run.</p></div>`) + section('Whole-package checkers', `<div class="prose"><p>Checkers derive the expected document from the input and the task specification, then compare the whole package. They check content, formatting, formulas and cached results, order, notes, metadata and unrelated parts. The runner also checks tool access, completion and that the input bytes are unchanged.</p><p>Equivalent serialization, relationship ids and resolved style indexes are normalized. Calculation chains and the DOCX save timestamp have documented exceptions. ${link(repo + 'eval/README.md#scoring', 'Exact scoring rules')}</p>${v ? `<p>Checker validation: ${v.accepted}/${v.accepted} correct outputs accepted, ${v.untouched_rejected}/${v.untouched_rejected} untouched inputs rejected, and ${v.corrupted_rejected}/${v.corrupted_rejected} corrupted outputs rejected.</p>` : ''}</div>`) + section('Pinned synthetic fixtures', `<div class="prose"><p>The v2 generator creates synthetic files with fixed timestamps. DOCX fixtures start as hand-written OOXML, are opened and saved by BetterOffice, and carry paragraph ids. The manifest pins each published file by size and SHA-256. Downloads and local fixtures are verified against those pins.</p><p>Regenerated bytes can change when engine output changes. A changed fixture set needs a new version prefix and updated pins. Fixed timestamps do not guarantee identical bytes across engine versions.</p></div>`) + section('Limits of v2', `<div class="prose"><p>The v2 results cover small synthetic fixtures, one run per task and one model. They measure the agent and tools together. They do not establish performance on real-world documents, other models or repeated runs. The checkers validate defined edits rather than visual rendering quality.</p><p>Tool counts differ by arm. Timings were recorded on different days and cannot support a speed comparison. Wilson intervals do not account for task selection or predict results on other document sets.</p>${footnotes(s)}</div>`);
}

function reproduce(s: Suite, data: SiteData) {
  const fixtures = Object.entries(data.fixtures.fixtures).sort(([a], [b]) => a.localeCompare(b));
  return hero('Reproduce the evals', s) + section('Run both arms', `<div class="prose"><p>Install the harness prerequisites and build the packages before running the BetterOffice arm. The generic arm needs a Python environment with python-docx, openpyxl and python-pptx. Results directories must be new or empty and outside the repository.</p><p>${link(repo + 'eval/README.md#running', 'Harness instructions')}</p>${data.commands.map(c => `<pre><code>${escape(c)}</code></pre>`).join('')}</div>`) + section('Every v2 fixture', `<div class="prose"><p>${link(data.fixtures.origin + '/metadata.json', 'metadata.json')} · ${link(repo + 'eval/formats/fixtures.json', 'Pinned sizes and SHA-256')} · ${link(repo + 'eval/formats/generate-fixtures.ts', 'Fixture generator')}</p><ul>${fixtures.map(([name, pin]) => `<li>${link(data.fixtures.origin + '/' + name, name)}${data.tasks.find(t => t.fixture === name) ? ` · ${escape(data.tasks.find(t => t.fixture === name)!.id)}` : ''}<p class="eyebrow">${pin.bytes} bytes · SHA-256 <code>${escape(pin.sha256)}</code></p></li>`).join('')}</ul></div>`);
}

export function renderPages(data: SiteData): Map<string, string> {
  const suites = newestFirst(data.suites);
  if (!suites.length) throw new Error('No evaluation suites');
  const pages = new Map<string, string>();
  for (const [index, s] of suites.entries()) {
    for (const base of index === 0 ? ['', `/suites/${encodeURIComponent(s.suite)}`] : [`/suites/${encodeURIComponent(s.suite)}`]) {
      const routes: [string, string][] = [ ['/', 'Overview'], ...formats(s).map((f): [string, string] => ['/' + f, f.toUpperCase()]), ['/failures', 'Failures'], ['/method', 'Method'], ['/reproduce', 'Reproduce'] ];
      for (const [route, title] of routes) {
        const body = route === '/' ? overview(s, suites, base) : route === '/failures' ? failures(s, base) : route === '/method' ? method(s) : route === '/reproduce' ? reproduce(s, data) : taskPage(s, route!.slice(1));
        const nav = `<nav class="flex flex-wrap items-center gap-x-7 gap-y-2 max-[44rem]:gap-x-5 max-[36rem]:gap-x-4" aria-label="Sections">${routes.map(([path, label]) => `<a class="font-mono text-xs text-ink no-underline transition-colors hover:text-fg" href="${location(base, path!)}"${path === route ? ' aria-current="page"' : ''}>${escape(label)}</a>`).join('')}</nav>`;
        const html = data.frame.replace('{{TITLE}}', () => escape(title!)).replace('{{NAV}}', () => nav).replace('{{CONTENT}}', () => body);
        if (/[\u2014]|undefined|NaN/.test(html)) throw new Error(`Invalid generated page: ${route}`);
        const path = location(base, route!);
        pages.set(path === '/' ? 'index.html' : path.slice(1) + '/index.html', html);
      }
    }
  }
  return pages;
}

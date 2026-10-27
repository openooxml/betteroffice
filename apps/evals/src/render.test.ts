import { expect, test } from 'bun:test';
import { renderPages, taskAnchor } from './render';
import type { Suite } from './stats';
import data from '../../../eval/results/v2.json';
import v3 from '../../../eval/results/v3.json';

const suite = data as Suite;
const frame = '{{TITLE}}{{NAV}}{{CONTENT}}';

test('single-tier suites retain the same rendered pages', () => {
  const implicit = renderPages({ suites: [suite], frame });
  const explicit = renderPages({ suites: [{ ...suite, tasks: suite.tasks.map(task => ({ ...task, tier: 'basic' })) }], frame });
  expect([...explicit]).toEqual([...implicit]);
});

test('multi-tier suites separate arm pass rates and format task tables', () => {
  const tasks = suite.tasks.filter(task => task.format === 'docx').slice(0, 2).map((task, index) => ({ ...task, tier: index ? 'complex' as const : 'basic' as const }));
  const future: Suite = {
    ...suite, suite: 'v3', tasks,
    runs: tasks.flatMap(task => suite.arms.flatMap(arm => [1, 2, 3].map(run => ({
      ...suite.runs[0]!, model: 'model-a', task: task.id, arm: arm.id, run,
      pass: task.tier === 'basic', reason: task.tier === 'basic' ? null : 'Failed complex edit.',
    })))),
  };
  const pages = renderPages({ suites: [future], frame, publishedTiers: ['basic', 'complex'] });
  const overview = pages.get('index.html')!;
  expect(overview.match(/<figure class="comparison">/g)).toHaveLength(1);
  for (const arm of suite.arms) {
    expect(overview).toContain(`Overall · ${arm.label}</span><strong class="tiered">complex 0/3 · basic 3/3</strong>`);
    expect(overview).toContain(`<td>${arm.label} · Complex</td><td>0/3</td>`);
    expect(overview).toContain(`<td>${arm.label} · Basic</td><td>3/3</td>`);
  }
  const page = pages.get('docx/index.html')!;
  const groups = page.split('<h3 class="eyebrow my-6">');
  expect(groups).toHaveLength(3);
  expect(groups[1]).toStartWith('Basic</h3>');
  expect(groups[2]).toStartWith('Complex</h3>');
  expect(groups[1]).toContain(taskAnchor('model-a', tasks[0]!.id));
  expect(groups[1]).not.toContain(taskAnchor('model-a', tasks[1]!.id));
  expect(groups[2]).toContain(taskAnchor('model-a', tasks[1]!.id));
  expect(page.match(/<table class="task-table">/g)).toHaveLength(2);
});

test('a basic-only allowlist hides complex results and copy on every route', () => {
  const pages = renderPages({ suites: [v3 as Suite], frame, publishedTiers: ['basic'] });
  const hidden = v3.tasks.filter(task => task.tier === 'complex');
  for (const page of pages.values()) {
    expect(page).not.toMatch(/\bcomplex\b/i);
    expect(page).not.toContain('<h3 class="eyebrow my-6">Basic</h3>');
    for (const task of hidden) expect(page).not.toContain(taskAnchor('gpt-6-luna', task.id));
  }
  const overview = pages.get('index.html')!;
  expect(overview).toContain('262/288');
  expect(overview).toContain('176/288');
  expect(overview).not.toContain('/387');
  expect(pages.get('docx/index.html')).toContain(taskAnchor('gpt-6-luna', v3.tasks.find(task => task.tier === 'basic' && task.format === 'docx')!.id));
});

test('both tiers are published by default', () => {
  const pages = renderPages({ suites: [v3 as Suite], frame });
  const overview = pages.get('index.html')!;
  for (const arm of v3.arms) for (const tier of ['Complex', 'Basic']) expect(overview).toContain(`</span>${arm.label} · ${tier}</li>`);
  expect([...overview.matchAll(/<li><span class="chart-key ([^"]+)"/g)].map(match => match[1])).toEqual(['generic', 'ours', 'generic basic', 'ours basic']);
  for (const format of ['docx', 'xlsx', 'pptx']) {
    const complex = v3.tasks.find(task => task.tier === 'complex' && task.format === format)!;
    expect(pages.get(`${format}/index.html`)).toContain(taskAnchor('gpt-6-luna', complex.id));
  }
});

test('grouped bars pair both arms for each format and total on one percentage axis', () => {
  const overview = renderPages({ suites: [v3 as Suite], frame }).get('index.html')!;
  const chart = overview.match(/<figure class="comparison">[\s\S]*?<\/figure>/)![0];
  expect(chart.match(/class="chart-axis"/g)).toHaveLength(1);
  expect(chart).toContain('<span>100%</span><span>50%</span><span>0%</span>');
  const groups = [...chart.matchAll(/<div class="chart-pair">([\s\S]*?)<span class="chart-label">([^<]+)/g)];
  expect(groups.map(group => group[2])).toEqual(['DOCX', 'XLSX', 'PPTX', 'Total']);
  for (const group of groups) {
    expect(group[1]).toContain('data-arm="generic"');
    expect(group[1]).toContain('data-arm="betteroffice"');
    expect(group[1]!.match(/class="chart-value"/g)).toHaveLength(4);
    expect([...group[1]!.matchAll(/data-arm="([^"]+)" data-tier="([^"]+)"/g)].map(match => match[1] + ' ' + match[2])).toEqual(['generic complex', 'betteroffice complex', 'generic basic', 'betteroffice basic']);
  }
  expect(chart).toContain('style="height:44.8%"');
  expect(chart).toContain('style="height:79.2%"');
  expect(chart).toContain('<table class="sr-only">');
  expect(chart).toContain('shared scale from 0% to 100%');
  expect(chart).toContain('<td>BetterOffice MCP · Basic</td><td>262/288</td><td>91.0%</td>');
  expect(overview).not.toContain('rate-bar');
});

import { expect, test } from 'bun:test';
import { renderPages, taskAnchor } from './render';
import type { Suite } from './stats';
import data from '../../../eval/results/v2.json';

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
  const pages = renderPages({ suites: [future], frame });
  const overview = pages.get('index.html')!;
  const basic = overview.split('<h3 class="eyebrow my-6">Basic</h3>')[1]!.split('<h3')[0]!;
  const complex = overview.split('<h3 class="eyebrow my-6">Complex</h3>')[1]!.split('</section>')[0]!;
  for (const arm of suite.arms) {
    expect(basic).toContain(`Overall · ${arm.label}</span><strong>3/3</strong>`);
    expect(complex).toContain(`Overall · ${arm.label}</span><strong>0/3</strong>`);
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

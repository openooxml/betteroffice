import { expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';

const FORK =
  'github.event.pull_request && github.event.pull_request.head.repo.full_name != github.repository';
const directory = new URL('../.github/workflows/', import.meta.url);
const kyora = readdirSync(directory)
  .filter((file) => file.endsWith('.yml'))
  .flatMap((file) => {
    const workflow = Bun.YAML.parse(readFileSync(new URL(file, directory), 'utf8')) as any;
    return Object.entries<any>(workflow.jobs)
      .filter(([, job]) => String(job['runs-on'] ?? '').includes('kyora'))
      .map(([id, job]) => ({ name: `${file} ${id}`, job, triggers: Object.keys(workflow.on) }));
  });

test('kyora jobs route fork pull requests to GitHub-hosted runners and honor BO_CI_RUNNER', () => {
  expect(kyora.length).toBeGreaterThan(0);
  for (const { name, job } of kyora) {
    expect(job['runs-on'], name).toStartWith(`\${{ (${FORK}`);
    expect(job['runs-on'], name).toContain("&& 'ubuntu-latest' ||");
    expect(job['runs-on'], name).toContain('vars.BO_CI_RUNNER');
  }
});

test('kyora jobs hold no secrets and never run for pull_request_target', () => {
  for (const { name, job, triggers } of kyora) {
    expect(triggers, name).not.toContain('pull_request_target');
    expect(JSON.stringify(job), name).not.toContain('secrets.');
  }
});

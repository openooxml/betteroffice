import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  PYTHON_BINDINGS,
  PYTHON_BINDING_NAMES,
  PYTHON_PUBLISH_NAMES
} from './python-bindings.mjs';

const script = fileURLToPath(new URL('./python-bindings.mjs', import.meta.url));
const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));
const ciWorkflow = fileURLToPath(new URL('../.github/workflows/ci.yml', import.meta.url));

function cli(...args: string[]) {
  const result = spawnSync('node', [script, ...args], { encoding: 'utf8' });
  expect(result.status).toBe(0);
  return result.stdout.trim();
}

const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;

describe('registry', () => {
  test('every publishable binding is a registered binding', () => {
    expect(PYTHON_BINDING_NAMES).toEqual(PYTHON_BINDINGS.map((p) => p.split('/python-')[1]));
    for (const name of PYTHON_PUBLISH_NAMES) {
      expect(PYTHON_BINDING_NAMES).toContain(name);
    }
  });

  test('pptx builds and versions but does not publish', () => {
    expect(PYTHON_BINDINGS).toContain('bindings/python-pptx');
    expect(PYTHON_BINDING_NAMES).toContain('pptx');
    expect(PYTHON_PUBLISH_NAMES).not.toContain('pptx');
  });

  test('xlsx publishes', () => {
    expect(PYTHON_PUBLISH_NAMES).toContain('xlsx');
  });

  test('the CLI reports the registry three ways', () => {
    expect(cli('--paths').split('\n')).toEqual(PYTHON_BINDINGS);
    expect(JSON.parse(cli())).toEqual(PYTHON_BINDING_NAMES);
    expect(JSON.parse(cli('--publish'))).toEqual(PYTHON_PUBLISH_NAMES);
  });
});

describe('release wiring', () => {
  test('the release train exposes both lists', () => {
    expect(release.jobs['python-bindings'].outputs).toEqual({
      bindings: '${{ steps.registry.outputs.bindings }}',
      publish: '${{ steps.registry.outputs.publish }}'
    });
  });

  test('builds fan out over every binding', () => {
    expect(release.jobs['python-dist'].strategy.matrix.binding).toBe(
      '${{ fromJSON(needs.python-bindings.outputs.bindings) }}'
    );
  });

  test('publishes fan out only over the opted-in bindings', () => {
    expect(release.jobs['python-pypi'].strategy.matrix.binding).toBe(
      '${{ fromJSON(needs.python-bindings.outputs.publish) }}'
    );
    expect(release.jobs['python-pypi'].if).toContain(
      "needs.python-bindings.outputs.publish != '[]'"
    );
  });

  test('CI installs and tests every binding, publishable or not', () => {
    const python = Bun.YAML.parse(readFileSync(ciWorkflow, 'utf8')) as any;
    const steps = python.jobs.python.steps.map((step: any) => step.run ?? '').join('\n');
    expect(steps).toContain('scripts/python-bindings.mjs --paths');
    expect(steps).not.toContain('--publish');
  });
});

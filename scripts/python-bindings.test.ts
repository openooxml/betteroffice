import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  PYPI_DISTRIBUTIONS,
  PYTHON_BINDINGS,
  PYTHON_BINDING_NAMES,
  PYTHON_PUBLISH_NAMES
} from './python-bindings.mjs';

const script = fileURLToPath(new URL('./python-bindings.mjs', import.meta.url));
const repository = fileURLToPath(new URL('..', import.meta.url));
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

  test('docx publishes', () => {
    expect(PYTHON_BINDINGS).toContain('bindings/python-docx');
    expect(PYTHON_BINDING_NAMES).toContain('docx');
    expect(PYTHON_PUBLISH_NAMES).toContain('docx');
    expect(PYPI_DISTRIBUTIONS).toContain('betteroffice-docx');
  });

  test('pptx publishes', () => {
    expect(PYTHON_BINDINGS).toContain('bindings/python-pptx');
    expect(PYTHON_BINDING_NAMES).toContain('pptx');
    expect(PYTHON_PUBLISH_NAMES).toContain('pptx');
    expect(PYPI_DISTRIBUTIONS).toContain('betteroffice-pptx');
  });

  test('xlsx publishes', () => {
    expect(PYTHON_PUBLISH_NAMES).toContain('xlsx');
    expect(PYPI_DISTRIBUTIONS).toContain('betteroffice-xlsx');
  });

  test('a held-back binding stays out of the publish matrix', () => {
    const held = PYTHON_BINDING_NAMES.filter((name) => !PYTHON_PUBLISH_NAMES.includes(name));
    for (const name of held) {
      expect(PYPI_DISTRIBUTIONS).not.toContain(`betteroffice-${name}`);
    }
    expect(PYPI_DISTRIBUTIONS).toHaveLength(PYTHON_PUBLISH_NAMES.length);
  });

  test('the CLI reports the registry three ways', () => {
    expect(cli('--paths').split('\n')).toEqual(PYTHON_BINDINGS);
    expect(JSON.parse(cli())).toEqual(PYTHON_BINDING_NAMES);
    expect(JSON.parse(cli('--publish'))).toEqual(PYTHON_PUBLISH_NAMES);
  });

  test('an unknown argument fails instead of printing every binding', () => {
    const result = spawnSync('node', [script, '--publsh'], { encoding: 'utf8' });
    expect(result.status).toBe(1);
    expect(result.stderr).toContain('--publsh');
    expect(result.stdout).toBe('');
  });
});

describe('release wiring', () => {
  test('the registry step computes each list from the registry', () => {
    const step = release.jobs['python-bindings'].steps.find((s: any) => s.id === 'registry');
    expect(step.run).toContain('scripts/python-bindings.mjs --publish');

    const directory = mkdtempSync(join(tmpdir(), 'registry-'));
    const outputs = join(directory, 'outputs');
    const path = join(directory, 'registry.sh');
    writeFileSync(path, step.run);
    const result = spawnSync('bash', ['-e', path], {
      encoding: 'utf8',
      cwd: repository,
      env: { ...process.env, GITHUB_OUTPUT: outputs }
    });

    expect(result.status).toBe(0);
    expect(readFileSync(outputs, 'utf8').trim().split('\n')).toEqual([
      `bindings=${JSON.stringify(PYTHON_BINDING_NAMES)}`,
      `publish=${JSON.stringify(PYTHON_PUBLISH_NAMES)}`
    ]);
  });

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

describe('repository-scoped token guard', () => {
  const job = release.jobs['python-bindings'];
  const guard = job.steps.find((step: any) => step.name === 'Refuse a repository-scoped PyPI token');

  function runGuard(token: string) {
    const path = join(mkdtempSync(join(tmpdir(), 'pypi-guard-')), 'guard.sh');
    writeFileSync(path, guard.run);
    return spawnSync('bash', ['-e', path], {
      encoding: 'utf8',
      env: { ...process.env, TOKEN: token }
    });
  }

  test('runs where an environment secret is invisible', () => {
    expect(job.environment).toBeUndefined();
    expect(guard.env.TOKEN).toBe('${{ secrets.PYPI_API_TOKEN }}');
    expect(release.jobs['python-pypi'].environment).toBe('pypi-${{ matrix.binding }}');
  });

  test('passes when no token is visible', () => {
    expect(runGuard('').status).toBe(0);
  });

  test('fails and says what to do when a token is visible', () => {
    const result = runGuard('pypi-AgEIcHlwaS5vcmcSECAGENUINE-LOOKING-TOKEN');
    expect(result.status).toBe(1);
    expect(result.stdout).toContain('::error::');
    expect(result.stdout).toContain('pypi-<binding>');
    expect(result.stdout).toContain('Trusted Publisher');
    expect(result.stdout).toContain('project-scoped');
    expect(result.stdout).not.toContain('AgEIcHlwaS5vcmc');
  });
});

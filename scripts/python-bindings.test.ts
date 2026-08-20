import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  PYPI_DISTRIBUTIONS,
  PYTHON_BINDINGS,
  PYTHON_BINDING_NAMES,
  PYTHON_PUBLISH_NAMES,
  bindingVersion,
  pendingPublishNames
} from './python-bindings.mjs';

const script = fileURLToPath(new URL('./python-bindings.mjs', import.meta.url));
const repository = fileURLToPath(new URL('..', import.meta.url));
const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));
const publishWorkflow = fileURLToPath(
  new URL('../.github/workflows/publish-python-binding.yml', import.meta.url)
);
const distWorkflow = fileURLToPath(new URL('../.github/workflows/python-dist.yml', import.meta.url));
const ciWorkflow = fileURLToPath(new URL('../.github/workflows/ci.yml', import.meta.url));

function cli(...args: string[]) {
  const result = spawnSync('node', [script, ...args], { encoding: 'utf8' });
  expect(result.status).toBe(0);
  return result.stdout.trim();
}

const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
const publish = Bun.YAML.parse(readFileSync(publishWorkflow, 'utf8')) as any;
const dist = Bun.YAML.parse(readFileSync(distWorkflow, 'utf8')) as any;

const SHA = '9b3f0c1d2e4a5b6c7d8e9f0a1b2c3d4e5f607182';
const dispatchStep = release.jobs['python-pypi'].steps[0];
const RUN_IDS = PYTHON_PUBLISH_NAMES.map((_, index) => 4200 + index);

/** The dispatch step against a `gh` that dispatches (returning a run id) and watches. */
function dispatch({ dispatchStatus = 0, failing = 0 } = {}) {
  const directory = mkdtempSync(join(tmpdir(), 'dispatch-'));
  const log = join(directory, 'gh.log');
  const ids = PYTHON_PUBLISH_NAMES.map(
    (name, index) => `    *"inputs[binding]=${name}"*) echo ${RUN_IDS[index]} ;;`
  );
  writeFileSync(
    join(directory, 'gh'),
    [
      '#!/bin/sh',
      `echo "$@" >> "${log}"`,
      'case "$1 $2" in',
      `  "api --method") [ ${dispatchStatus} -ne 0 ] && exit ${dispatchStatus}; case "$*" in`,
      ...ids,
      '    esac ;;',
      `  "run watch") [ "$3" = "${failing}" ] && exit 1 ;;`,
      'esac',
      'exit 0'
    ].join('\n'),
    { mode: 0o755 }
  );
  const path = join(directory, 'dispatch.sh');
  writeFileSync(path, dispatchStep.run);
  const result = spawnSync('bash', ['-e', path], {
    encoding: 'utf8',
    env: {
      ...process.env,
      PATH: `${directory}:${process.env.PATH}`,
      GH_REPO: 'openooxml/betteroffice',
      PUBLISH: JSON.stringify(PYTHON_PUBLISH_NAMES),
      SHA
    }
  });
  return { ...result, log: readFileSync(log, 'utf8').trim().split('\n') };
}

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
  test('the registry step computes the publish list from PyPI state', () => {
    const step = release.jobs['python-bindings'].steps.find((s: any) => s.id === 'registry');
    expect(step.run).toContain('scripts/python-bindings.mjs --pending');

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
    const lines = readFileSync(outputs, 'utf8').trim().split('\n');
    expect(lines).toHaveLength(1);
    const pending = JSON.parse(lines[0]!.replace(/^publish=/, '')) as string[];
    expect(PYTHON_PUBLISH_NAMES.filter((name) => pending.includes(name))).toEqual(pending);
  });

  test('the release train exposes the publish list', () => {
    expect(release.jobs['python-bindings'].outputs).toEqual({
      publish: '${{ steps.registry.outputs.publish }}'
    });
  });

  test('the release publishes nothing to PyPI itself', () => {
    expect(release.jobs['python-dist']).toBeUndefined();
    const runs = JSON.stringify(release.jobs['python-pypi']);
    expect(runs).not.toContain('pypa/gh-action-pypi-publish');
    expect(runs).not.toContain('PYPI_API_TOKEN');
    expect(release.jobs['python-pypi'].permissions).toEqual({ actions: 'write' });
  });

  test('a dispatch fans out only over the opted-in bindings, at the commit that triggered it', () => {
    expect(release.jobs['python-pypi'].needs).toBe('python-bindings');
    expect(release.jobs['python-pypi'].if).toContain(
      "needs.python-bindings.outputs.publish != '[]'"
    );
    expect(dispatchStep.env.PUBLISH).toBe('${{ needs.python-bindings.outputs.publish }}');
    expect(dispatchStep.env.SHA).toBe('${{ github.sha }}');

    const result = dispatch();

    expect(result.status).toBe(0);
    expect(result.log.filter((line) => line.startsWith('api --method POST'))).toEqual(
      PYTHON_PUBLISH_NAMES.map(
        (name) =>
          `api --method POST -H X-GitHub-Api-Version: 2026-03-10 repos/openooxml/betteroffice/actions/workflows/publish-python-binding.yml/dispatches -f ref=main -f inputs[binding]=${name} -f inputs[sha]=${SHA} -F inputs[dry_run]=false --jq .workflow_run_id`
      )
    );
  });

  test('the wait watches exactly the runs the dispatch created', () => {
    const result = dispatch();

    expect(result.status).toBe(0);
    expect(result.log.filter((line) => line.startsWith('run watch'))).toEqual(
      RUN_IDS.map((id) => `run watch ${id} --exit-status`)
    );
    expect(publish['run-name']).toBe(
      'Publish betteroffice-${{ inputs.binding }} @ ${{ inputs.sha }}'
    );
  });

  test('the release ends only once every dispatched run has', () => {
    const result = dispatch();

    expect(result.status).toBe(0);
    expect(result.log.filter((line) => line.startsWith('run watch'))).toHaveLength(
      PYTHON_PUBLISH_NAMES.length
    );
  });

  test('a dispatched run that fails fails the release', () => {
    const result = dispatch({ failing: RUN_IDS[0]! });

    expect(result.status).toBe(1);
    expect(result.stdout).toContain(`::error::betteroffice-${PYTHON_PUBLISH_NAMES[0]} did not`);
    expect(result.log.filter((line) => line.startsWith('run watch'))).toHaveLength(
      PYTHON_PUBLISH_NAMES.length
    );
  });

  test('one failed dispatch fails the job without skipping the rest', () => {
    const result = dispatch({ dispatchStatus: 1 });

    expect(result.status).toBe(1);
    expect(result.log.filter((line) => line.startsWith('api --method POST'))).toHaveLength(
      PYTHON_PUBLISH_NAMES.length
    );
    expect(result.log.some((line) => line.startsWith('run watch'))).toBe(false);
  });

  test('the dispatched workflow is the one PyPI trusts', () => {
    expect(Object.keys(publish.on)).toEqual(['workflow_dispatch']);
    expect(Object.keys(publish.on.workflow_dispatch.inputs)).toEqual(['binding', 'sha', 'dry_run']);
    expect(publish.on.workflow_dispatch.inputs.sha.required).toBe(true);
    expect(publish.jobs.publish.environment).toBe('pypi-${{ inputs.binding }}');
    expect(publish.jobs.publish.permissions['id-token']).toBe('write');
  });

  test('the wheels are built from that commit, not from a branch that moves', () => {
    expect(dist.on.workflow_call.inputs.sha.required).toBe(true);
    expect(publish.jobs.dist.with.sha).toBe('${{ inputs.sha }}');

    const checkouts = Object.values(dist.jobs).flatMap((job: any) =>
      (job.steps ?? []).filter((step: any) => String(step.uses ?? '').startsWith('actions/checkout'))
    );
    expect(checkouts).toHaveLength(2);
    for (const step of checkouts) expect(step.with.ref).toBe('${{ inputs.sha }}');
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
    expect(publish.jobs.publish.environment).toBe('pypi-${{ inputs.binding }}');
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
    expect(result.stdout).not.toContain('AgEIcHlwaS5vcmc');
  });

  test('offers no token path, because the workflow that uploads refuses one', () => {
    const result = runGuard('pypi-AgEIcHlwaS5vcmcSECAGENUINE-LOOKING-TOKEN');
    expect(result.stdout).toContain('There is no token path');
    expect(result.stdout).not.toContain('project-scoped');

    const child = publish.jobs.publish.steps.find(
      (step: any) => step.name === 'Refuse a repository-scoped PyPI token'
    );
    expect(child.env.TOKEN).toBe('${{ secrets.PYPI_API_TOKEN }}');
    expect(child.run).toContain('exit 1');
  });
});

describe('pending publish selection', () => {
  const live = () =>
    Object.fromEntries(
      PYTHON_BINDINGS.map((path) => [
        `betteroffice-${path.replace('bindings/python-', '')}`,
        [bindingVersion(path)]
      ])
    );

  type Files = { yanked?: boolean }[];

  function pypi(releases: Record<string, string[] | Record<string, Files> | null>) {
    const requested: string[] = [];
    const fetchImpl = async (url: string) => {
      requested.push(url);
      const project = url.match(/\/pypi\/([^/]+)\/json$/)?.[1] ?? '';
      const entry = releases[project];
      if (entry == null) return { status: 404, ok: false } as Response;
      const map = Array.isArray(entry)
        ? Object.fromEntries(entry.map((version) => [version, [{ yanked: false }]]))
        : entry;
      return {
        status: 200,
        ok: true,
        json: async () => ({ releases: map })
      } as unknown as Response;
    };
    return { fetchImpl, requested };
  }

  test('every binding version is a semver', () => {
    for (const path of PYTHON_BINDINGS) expect(bindingVersion(path)).toMatch(/^\d+\.\d+\.\d+$/);
  });

  test('nothing is pending when PyPI serves every version', async () => {
    const { fetchImpl, requested } = pypi(live());
    expect(await pendingPublishNames({ fetchImpl })).toEqual([]);
    expect(requested).toEqual(PYPI_DISTRIBUTIONS.map((d) => `https://pypi.org/pypi/${d}/json`));
  });

  test('a version PyPI does not serve is pending', async () => {
    const releases = live();
    releases['betteroffice-docx'] = ['0.0.0'];
    const { fetchImpl } = pypi(releases);
    expect(await pendingPublishNames({ fetchImpl })).toEqual(['docx']);
  });

  test('a project PyPI does not know is pending', async () => {
    const releases = live();
    releases['betteroffice-xlsx'] = null;
    const { fetchImpl } = pypi(releases);
    expect(await pendingPublishNames({ fetchImpl })).toEqual(['xlsx']);
  });

  test('a version whose files are all deleted fails instead of counting as released', async () => {
    const releases = live();
    releases['betteroffice-docx'] = { [bindingVersion('bindings/python-docx')]: [] };
    const { fetchImpl } = pypi(releases);
    await expect(pendingPublishNames({ fetchImpl })).rejects.toThrow('no installable file');
  });

  test('a version whose files are all yanked fails instead of counting as released', async () => {
    const releases = live();
    releases['betteroffice-docx'] = {
      [bindingVersion('bindings/python-docx')]: [{ yanked: true }, { yanked: true }]
    };
    const { fetchImpl } = pypi(releases);
    await expect(pendingPublishNames({ fetchImpl })).rejects.toThrow('no installable file');
  });

  test('one live file among yanked ones counts as released', async () => {
    const releases = live();
    releases['betteroffice-docx'] = {
      [bindingVersion('bindings/python-docx')]: [{ yanked: true }, { yanked: false }]
    };
    const { fetchImpl } = pypi(releases);
    expect(await pendingPublishNames({ fetchImpl })).toEqual([]);
  });

  test('a version is read from [package], not another table', () => {
    const fixture = 'scripts/.manifest-fixture';
    const directory = join(repository, fixture);
    const read = (body: string) => {
      mkdirSync(directory, { recursive: true });
      writeFileSync(join(directory, 'Cargo.toml'), body);
      try {
        return bindingVersion(fixture);
      } catch (error) {
        return `THROWS: ${(error as Error).message}`;
      }
    };

    try {
      expect(read('[package]\nversion = "1.2.3" # release\n')).toBe('1.2.3');
      expect(read('[other]\nversion = "9.9.9"\n\n[package]\nversion = "1.2.3"\n')).toBe('1.2.3');
      expect(read('[package]\nversion.workspace = true\n\n[other]\nversion = "9.9.9"\n')).toContain(
        'THROWS'
      );
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test('an unreadable PyPI fails the run instead of skipping a publish', async () => {
    const fetchImpl = async () => ({ status: 503, ok: false }) as Response;
    await expect(
      pendingPublishNames({ fetchImpl, attempts: 2, retryDelayMs: 0 })
    ).rejects.toThrow('cannot read PyPI');
  });

  test('a transient error is retried', async () => {
    let calls = 0;
    const good = pypi(live()).fetchImpl;
    const fetchImpl = async (url: string) => {
      calls += 1;
      if (calls === 1) throw new Error('reset');
      return good(url);
    };
    expect(await pendingPublishNames({ fetchImpl, retryDelayMs: 0 })).toEqual([]);
  });
});

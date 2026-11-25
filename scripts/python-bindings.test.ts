import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import {
  FONT_PACK_NAMES,
  PYPI_DISTRIBUTIONS,
  PYTHON_BINDINGS,
  PYTHON_BINDING_NAMES,
  PYTHON_PUBLISH_NAMES,
  bindingVersion,
  fontPackVersion,
  pendingPublishNames
} from './python-bindings.mjs';
import { NATIVE_FORMATS } from './native-bindings.mjs';

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
const dispatchStep = release.jobs['python-pypi'].steps.find(
  (step: any) => step.name === 'Dispatch a publish per binding, then wait for each'
);
const RUN_IDS = PYTHON_PUBLISH_NAMES.map((_, index) => 4200 + index);
const ALL_NAMES = [...FONT_PACK_NAMES, ...PYTHON_PUBLISH_NAMES];
const ALL_RUN_IDS = ALL_NAMES.map((_, index) => 4300 + index);

/** The dispatch step against a `gh` that dispatches (returning a run id) and watches. */
function dispatch({ dispatchStatus = 0, failing = 0, timedOut = 0, checkStatus = 0, names = PYTHON_PUBLISH_NAMES, runIds = RUN_IDS } = {}) {
  const directory = mkdtempSync(join(tmpdir(), 'dispatch-'));
  const log = join(directory, 'gh.log');
  const ids = names.map(
    (name, index) => `    *"inputs[binding]=${name}"*) echo ${runIds[index]} ;;`
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
      `  "run watch") [ "$3" = "${timedOut}" ] && exit 124; [ "$3" = "${failing}" ] && exit 1 ;;`,
      'esac',
      'exit 0'
    ].join('\n'),
    { mode: 0o755 }
  );
  writeFileSync(join(directory, 'timeout'), '#!/bin/sh\nshift\nexec "$@"\n', { mode: 0o755 });
  const check = join(directory, 'check-publisher-files.sh');
  writeFileSync(check, `echo check-publisher-files >> "${log}"\nexit ${checkStatus}\n`);
  const path = join(directory, 'dispatch.sh');
  writeFileSync(path, dispatchStep.run);
  const result = spawnSync('bash', ['-e', path], {
    encoding: 'utf8',
    env: {
      ...process.env,
      PATH: `${directory}:${process.env.PATH}`,
      GH_REPO: 'openooxml/betteroffice',
      PUBLISH: JSON.stringify(names),
      SHA,
      CHECK_PUBLISHER_FILES: check
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
  const step = release.jobs.plan.steps.find((s: any) => s.id === 'bindings');

  test('the plan selects unpublished versions from both registries', () => {
    expect(step.run).toContain('scripts/native-bindings.mjs --pending');
    expect(step.run).toContain('scripts/python-bindings.mjs --pending');
    expect(step.run).not.toContain(',,}');

    const directory = mkdtempSync(join(tmpdir(), 'registry-'));
    const outputs = join(directory, 'outputs');
    const path = join(directory, 'registry.sh');
    const [served, ...unserved] = PYTHON_PUBLISH_NAMES;
    const releases = {
      [`betteroffice-${served}`]: { [bindingVersion(`bindings/python-${served}`)]: [{ yanked: false }] }
    };
    const pypi = join(directory, 'pypi.mjs');
    writeFileSync(
      pypi,
      `const served = ${JSON.stringify(releases)};
globalThis.fetch = async (url) => {
  if (new URL(url).host === 'registry.npmjs.org') return Response.json({ versions: {} });
  const releases = served[String(url).match(/\\/pypi\\/([^/]+)\\/json$/)?.[1]];
  return releases ? Response.json({ releases }) : new Response(null, { status: 404 });
};
`
    );
    writeFileSync(path, step.run);
    const result = spawnSync('bash', ['-eo', 'pipefail', path], {
      encoding: 'utf8',
      cwd: repository,
      env: {
        ...process.env,
        REHEARSE: 'false',
        PUBLISHING: 'true',
        RELEASE_SKIP_NATIVE: '',
        RELEASE_SKIP_PYTHON: 'false',
        GITHUB_OUTPUT: outputs,
        NODE_OPTIONS: `${process.env.NODE_OPTIONS ?? ''} --import=${pathToFileURL(pypi).href}`
      }
    });

    expect(result.status, result.stderr).toBe(0);
    expect(readFileSync(outputs, 'utf8')).toBe(
      `native=${JSON.stringify(NATIVE_FORMATS)}\npython=${JSON.stringify([...FONT_PACK_NAMES, ...unserved])}\nskip_native=false\nskip_python=false\n`
    );
  });

  test.each([
    ['TrUe', 'false', 'true', 'false'],
    ['false', 'TRUE', 'false', 'true'],
    ['TRUE', 'True', 'true', 'true']
  ] as const)('the release plan bypasses pending lookups with native=%s and python=%s', (native, python, skipNative, skipPython) => {
    const directory = mkdtempSync(join(tmpdir(), 'release-plan-'));
    try {
      const outputs = join(directory, 'outputs');
      writeFileSync(join(directory, 'node'), [
        '#!/bin/sh',
        'printf "%s\\n" "$*" >&2',
        'case "$*" in',
        '  "scripts/native-bindings.mjs --pending") echo \'["native"]\' ;;',
        '  "scripts/python-bindings.mjs --pending") echo \'["python"]\' ;;',
        '  *) exit 1 ;;',
        'esac',
        ''
      ].join('\n'), { mode: 0o755 });
      const result = spawnSync('bash', ['-eo', 'pipefail', '-c', step.run], {
        encoding: 'utf8',
        cwd: repository,
        env: {
          ...process.env,
          REHEARSE: 'false',
          PUBLISHING: 'true',
          RELEASE_SKIP_NATIVE: native,
          RELEASE_SKIP_PYTHON: python,
          PATH: `${directory}:${process.env.PATH}`,
          GITHUB_OUTPUT: outputs
        }
      });
      expect(result.status, result.stderr).toBe(0);
      expect(result.stderr).toBe([
        skipNative === 'false' ? 'scripts/native-bindings.mjs --pending\n' : '',
        skipPython === 'false' ? 'scripts/python-bindings.mjs --pending\n' : ''
      ].join(''));
      expect(readFileSync(outputs, 'utf8')).toBe([
        `native=${skipNative === 'true' ? '[]' : '["native"]'}`,
        `python=${skipPython === 'true' ? '[]' : '["python"]'}`,
        `skip_native=${skipNative}`,
        `skip_python=${skipPython}`,
        ''
      ].join('\n'));
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test('a registry lookup failure fails the step instead of publishing nothing', () => {
    const directory = mkdtempSync(join(tmpdir(), 'registry-fail-'));
    writeFileSync(join(directory, 'node'), '#!/bin/sh\nexit 1\n', { mode: 0o755 });
    const path = join(directory, 'registry.sh');
    writeFileSync(path, step.run);
    const result = spawnSync('bash', ['-e', path], {
      encoding: 'utf8',
      cwd: repository,
      env: {
        ...process.env,
        REHEARSE: 'false',
        PUBLISHING: 'true',
        RELEASE_SKIP_NATIVE: 'false',
        RELEASE_SKIP_PYTHON: 'false',
        PATH: `${directory}:${process.env.PATH}`,
        GITHUB_OUTPUT: join(directory, 'outputs')
      }
    });

    expect(result.status).not.toBe(0);
  });

  test.each([
    ['false', [], [], 'true'],
    ['true', NATIVE_FORMATS, [...FONT_PACK_NAMES, ...PYTHON_BINDING_NAMES], 'false']
  ] as const)('plans the release PR or rehearsal with REHEARSE=%s', (rehearse, native, python, skip) => {
    const directory = mkdtempSync(join(tmpdir(), 'release-plan-'));
    try {
      const outputs = join(directory, 'outputs');
      const result = spawnSync('bash', ['-eo', 'pipefail', '-c', step.run], {
        encoding: 'utf8',
        cwd: repository,
        env: {
          ...process.env,
          REHEARSE: rehearse,
          PUBLISHING: 'false',
          RELEASE_SKIP_NATIVE: 'TRUE',
          RELEASE_SKIP_PYTHON: 'True',
          GITHUB_OUTPUT: outputs
        }
      });
      expect(result.status, result.stderr).toBe(0);
      expect(readFileSync(outputs, 'utf8')).toBe(
        `native=${JSON.stringify(native)}\npython=${JSON.stringify(python)}\nskip_native=${skip}\nskip_python=${skip}\n`
      );
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test('release gates publishing on the Python build matrix for this commit', () => {
    const job = release.jobs['python-build'];
    expect(job.needs).toBe('plan');
    expect(job.if).toBe(
      "needs.plan.outputs.publishing == 'true' && needs.plan.outputs.python != '[]' && needs.plan.outputs.skip_python != 'true'"
    );
    expect(job.strategy).toEqual({
      'fail-fast': false,
      matrix: { binding: '${{ fromJSON(needs.plan.outputs.python) }}' }
    });
    expect(job.uses).toBe('./.github/workflows/python-dist.yml');
    expect(job.with).toEqual({ binding: '${{ matrix.binding }}', sha: '${{ github.sha }}' });
    expect(job.permissions).toEqual({ contents: 'read' });
    expect(job['continue-on-error']).toBeUndefined();
  });

  test('release retries the whole Python build matrix once after failure or cancellation', () => {
    const job = release.jobs['python-build'];
    const retry = release.jobs['python-build-retry'];
    expect(retry.needs).toEqual(['plan', 'python-build']);
    expect(retry.if).toBe("${{ !cancelled() && (needs.python-build.result == 'failure' || needs.python-build.result == 'cancelled') }}");
    for (const key of ['strategy', 'uses', 'with', 'permissions', 'secrets']) {
      expect(retry[key]).toEqual(job[key]);
    }
    expect(retry['continue-on-error']).toBeUndefined();
  });

  test('the release train exposes the publish list after the Python gate passes', () => {
    const job = release.jobs['python-bindings'];
    expect(job.needs).toEqual(['plan', 'release', 'python-build', 'python-build-retry']);
    expect(job.if).toBe([
      "${{ !cancelled() && needs.release.result == 'success' && needs.release.outputs.publishing == 'true' &&",
      "needs.plan.outputs.skip_python != 'true' &&",
      "(needs.python-build.result == 'success' ||",
      "((needs.python-build.result == 'failure' || needs.python-build.result == 'cancelled') && needs.python-build-retry.result == 'success') ||",
      "(needs.python-build.result == 'skipped' && needs.plan.outputs.python == '[]')) }}"
    ].join(' '));
    expect(job.outputs).toEqual({
      publish: '${{ needs.release.outputs.python }}'
    });
    expect(release.jobs.release.outputs.python).toBe('${{ needs.plan.outputs.python }}');
  });

  test('the release publishes nothing to PyPI itself', () => {
    expect(release.jobs['python-dist']).toBeUndefined();
    const runs = JSON.stringify(release.jobs['python-pypi']);
    expect(runs).not.toContain('pypa/gh-action-pypi-publish');
    expect(runs).not.toContain('PYPI_API_TOKEN');
    expect(release.jobs['python-pypi'].permissions).toEqual({ contents: 'read', actions: 'write' });
  });

  test('a dispatch fans out only over the opted-in bindings, at the commit that triggered it', () => {
    expect(release.jobs['python-pypi'].needs).toEqual(['release', 'python-bindings']);
    expect(release.jobs['python-pypi'].if).toBe(
      "needs.python-bindings.outputs.publish != '[]'"
    );
    expect(dispatchStep.env.PUBLISH).toBe('${{ needs.python-bindings.outputs.publish }}');
    expect(dispatchStep.env.SHA).toBe('${{ github.sha }}');
    expect(dispatchStep.env.CHECK_PUBLISHER_FILES).toBe('scripts/check-publisher-files.sh');

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
    expect(result.log.filter((line) => line.startsWith('run watch')).sort()).toEqual(
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
    expect(result.log.some((line) => line.startsWith('run cancel'))).toBe(false);
  });

  test('watch bounds fit the job deadline and a timeout cancels the stalled child', () => {
    const job = release.jobs['python-pypi'];
    expect(dispatchStep.run).toContain('timeout "${minutes}m" gh run watch "$run" --exit-status');
    const packs = Number(dispatchStep.run.match(/^publish (\d+) \$packs$/m)?.[1]);
    const formats = Number(dispatchStep.run.match(/^publish (\d+) \$formats$/m)?.[1]);
    const wheelAndLoad = dist.jobs.wheels['timeout-minutes'] + dist.jobs.load['timeout-minutes'];
    const sdist = dist.jobs.sdist['timeout-minutes'];
    const fontPack = dist.jobs['font-pack']['timeout-minutes'];
    const afterBuild = dist.jobs.built['timeout-minutes'] + publish.jobs.publish['timeout-minutes'] + 20;
    expect(packs).toBeGreaterThanOrEqual(fontPack + afterBuild);
    expect(formats).toBeGreaterThanOrEqual(Math.max(wheelAndLoad, sdist) + afterBuild);
    expect(job['timeout-minutes']).toBeGreaterThanOrEqual(packs + formats + 10);
    expect(dispatchStep.run).toContain('wait "$watcher" || status=1');

    const result = dispatch({ timedOut: RUN_IDS[0]! });
    expect(result.status).toBe(1);
    expect(result.log.filter((line) => line.startsWith('run cancel'))).toEqual([`run cancel ${RUN_IDS[0]}`]);
    expect(result.stdout).toContain(`::error::betteroffice-${PYTHON_PUBLISH_NAMES[0]} did not publish`);
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
    expect(checkouts).toHaveLength(4);
    for (const step of checkouts) expect(step.with.ref).toBe('${{ inputs.sha }}');
  });

  test('every Python dist upload can replace the first attempt artifact', () => {
    const uploads = Object.values(dist.jobs).flatMap((job: any) =>
      (job.steps ?? []).filter((step: any) => step.uses?.startsWith('actions/upload-artifact@'))
    );
    expect(uploads).toHaveLength(4);
    for (const upload of uploads) {
      expect(upload.uses).toBe('actions/upload-artifact@v7');
      expect(upload.with.overwrite).toBe(true);
    }
  });

  test('CI installs and tests every binding, publishable or not', () => {
    const python = Bun.YAML.parse(readFileSync(ciWorkflow, 'utf8')) as any;
    const steps = python.jobs.python.steps.map((step: any) => step.run ?? '').join('\n');
    expect(steps).toContain('scripts/python-bindings.mjs --paths');
    expect(steps).not.toContain('--publish');
  });

  test('CI installs the font pack wheels built from this commit before any binding', () => {
    const python = Bun.YAML.parse(readFileSync(ciWorkflow, 'utf8')) as any;
    const steps = python.jobs.python.steps.map((step: any) => step.run ?? '').join('\n');
    const packs = steps.indexOf('python -m pip install "$RUNNER_TEMP"/font-packs/*.whl');
    expect(steps.indexOf('scripts/python_font_packs.py --out "$RUNNER_TEMP/font-packs"')).toBeGreaterThan(-1);
    expect(packs).toBeGreaterThan(steps.indexOf('scripts/python_font_packs.py'));
    expect(packs).toBeLessThan(steps.indexOf('python -m pip install --find-links "$RUNNER_TEMP/font-packs" -e "$binding"'));
  });

  test('font packs publish first, and a failed pack dispatches no format binding', () => {
    const ok = dispatch({ names: ALL_NAMES, runIds: ALL_RUN_IDS });
    expect(ok.status, ok.stdout).toBe(0);
    const order = ok.log
      .filter((line) => line.startsWith('api --method POST') || line.startsWith('run watch'))
      .map((line) => line.match(/inputs\[binding\]=(\S+)/)?.[1] ?? line);
    const packs = FONT_PACK_NAMES.length;
    const formatWatches = 2 * packs + PYTHON_PUBLISH_NAMES.length;
    expect(order.slice(0, packs)).toEqual(FONT_PACK_NAMES);
    expect(order.slice(packs, 2 * packs).sort()).toEqual(
      ALL_RUN_IDS.slice(0, packs).map((id) => `run watch ${id} --exit-status`)
    );
    expect(order.slice(2 * packs, formatWatches)).toEqual(PYTHON_PUBLISH_NAMES);
    expect(order.slice(formatWatches).sort()).toEqual(
      ALL_RUN_IDS.slice(packs).map((id) => `run watch ${id} --exit-status`)
    );

    const failed = dispatch({ names: ALL_NAMES, runIds: ALL_RUN_IDS, failing: ALL_RUN_IDS[1]! });
    expect(failed.status).toBe(1);
    expect(failed.stdout).toContain(`::error::betteroffice-${FONT_PACK_NAMES[1]} did not publish`);
    expect(failed.stdout).toContain('no format binding was dispatched');
    const dispatched = failed.log
      .filter((line) => line.startsWith('api --method POST'))
      .map((line) => line.match(/inputs\[binding\]=(\S+)/)?.[1]);
    expect(dispatched).toEqual(FONT_PACK_NAMES);
  });

  test('the publisher files are compared again after the font packs and before any format binding', () => {
    const ok = dispatch({ names: ALL_NAMES, runIds: ALL_RUN_IDS });
    expect(ok.status, ok.stdout).toBe(0);
    const check = ok.log.indexOf('check-publisher-files');
    const lastPackWatch = Math.max(
      ...ALL_RUN_IDS.slice(0, FONT_PACK_NAMES.length).map((id) => ok.log.indexOf(`run watch ${id} --exit-status`))
    );
    const firstFormat = ok.log.findIndex((line) => line.includes(`inputs[binding]=${PYTHON_PUBLISH_NAMES[0]} `));
    expect(check).toBeGreaterThan(lastPackWatch);
    expect(check).toBeLessThan(firstFormat);

    const stale = dispatch({ names: ALL_NAMES, runIds: ALL_RUN_IDS, checkStatus: 1 });
    expect(stale.status).not.toBe(0);
    const dispatched = stale.log
      .filter((line) => line.startsWith('api --method POST'))
      .map((line) => line.match(/inputs\[binding\]=(\S+)/)?.[1]);
    expect(dispatched).toEqual(FONT_PACK_NAMES);

    expect(dispatch().log).not.toContain('check-publisher-files');
  });

  test('a font pack builds and checks one pure wheel from the committed assets, and no Rust', () => {
    for (const job of ['wheels', 'sdist', 'load']) {
      expect(dist.jobs[job].if).toBe("${{ !startsWith(inputs.binding, 'fonts-') }}");
    }
    const pack = dist.jobs['font-pack'];
    expect(pack.if).toBe("${{ startsWith(inputs.binding, 'fonts-') }}");
    expect(pack['continue-on-error']).toBeUndefined();
    const build = pack.steps.find((step: any) => step.name === 'Build the wheel');
    expect(build.run).toBe('python -I scripts/python_font_packs.py --out dist --pack "${BINDING#fonts-}"');
    const smoke = pack.steps[pack.steps.indexOf(build) + 1];
    expect(smoke.name).toBe('Test the installed font pack');
    expect(smoke.if).toBeUndefined();
    expect(smoke['continue-on-error']).toBeUndefined();
    expect(smoke.env).toEqual({ BINDING: '${{ inputs.binding }}' });
    expect(smoke.run).toBe([
      'python -m pip install --no-index --only-binary :all: dist/*.whl',
      'python -I - "$BINDING" <<\'PY\'',
      'import json',
      'import sys',
      'from importlib.resources import files',
      'from pathlib import Path',
      '',
      'source = Path("packages", sys.argv[1], "assets", "manifest.json")',
      'data = json.loads(source.read_text(encoding="utf-8"))',
      'pack = files("betteroffice_fonts_packs").joinpath(data["pack"], data["version"])',
      'for name in ["manifest.json", *(face["file"] for face in data["faces"])]:',
      '    assert pack.joinpath(name).is_file(), name',
      'PY',
      ''
    ].join('\n'));
    const upload = pack.steps.find((step: any) => String(step.uses).startsWith('actions/upload-artifact'));
    expect(pack.steps.indexOf(upload)).toBeGreaterThan(pack.steps.indexOf(smoke));
    expect(upload.with.name).toBe('python-${{ inputs.binding }}.wheel');
    expect(dist.jobs.built.needs).toEqual(['wheels', 'sdist', 'font-pack', 'load']);
    expect(dist.jobs.built.if).toContain("!contains(needs.*.result, 'failure')");
    expect(dist.jobs.built.if).toContain("needs.font-pack.result == 'success'");
  });

  test('loads each native wheel from its exact artifact and tests the installed binding', () => {
    const load = dist.jobs.load;
    expect(load.needs).toBe('wheels');
    expect(load['continue-on-error']).toBeUndefined();
    expect(load['runs-on']).toBe('${{ matrix.platform.os }}');
    expect(load.defaults.run.shell).toBe('bash');
    expect(load.strategy).toEqual({
      'fail-fast': false,
      matrix: { platform: [
        { os: 'ubuntu-24.04', artifact: 'ubuntu-latest.x86_64' },
        { os: 'ubuntu-24.04-arm', artifact: 'ubuntu-latest.aarch64' },
        { os: 'macos-latest', artifact: 'macos-latest.aarch64' },
        { os: 'macos-15-intel', artifact: 'macos-latest.x86_64' },
        { os: 'windows-latest', artifact: 'windows-latest.x64' }
      ] }
    });
    const wheels = dist.jobs.wheels;
    const upload = wheels.steps.find((step: any) => step.uses === 'actions/upload-artifact@v7');
    expect(upload.with.name).toBe('python-${{ inputs.binding }}.${{ matrix.platform.os }}.${{ matrix.platform.target }}');
    const wheelArtifacts = wheels.strategy.matrix.platform.map(
      (platform: any) => `${platform.os}.${platform.target}`
    );
    const loadArtifacts = load.strategy.matrix.platform.map((platform: any) => platform.artifact);
    expect(loadArtifacts.sort()).toEqual(wheelArtifacts.sort());
    const [lineEndings, checkout, python, download, install, pytest] = load.steps;
    expect(lineEndings.run).toBe('git config --global core.autocrlf false');
    expect(checkout.uses).toBe('actions/checkout@v7.0.1');
    expect(python.uses).toBe('actions/setup-python@v7.1.0');
    expect(python.with['python-version']).toBe('3.13');
    expect(download.uses).toBe('actions/download-artifact@v8');
    expect(download.with).toEqual({ name: 'python-${{ inputs.binding }}.${{ matrix.platform.artifact }}', path: 'wheel' });
    expect(install.if).toBeUndefined();
    expect(install['continue-on-error']).toBeUndefined();
    expect(install.run).toContain('python -I scripts/python_font_packs.py --out font-packs');
    expect(install.run).toContain('python -m pip install --only-binary :all: --find-links font-packs wheel/*.whl');
    expect(pytest.if).toBeUndefined();
    expect(pytest['continue-on-error']).toBeUndefined();
    expect(pytest['working-directory']).toBe('${{ runner.temp }}');
    expect(pytest.env).toEqual({ BINDING: '${{ inputs.binding }}', PYTHONUTF8: '1' });
    expect(pytest.run).toBe('python -m pytest -q "$GITHUB_WORKSPACE/bindings/python-$BINDING/tests"');
    expect(dist.jobs.built.if).toContain("needs.wheels.result == 'success' && needs.sdist.result == 'success' && needs.load.result == 'success'");
  });
});

describe('font pack wheels', () => {
  const versions = Object.fromEntries(FONT_PACK_NAMES.map((name) => [name, fontPackVersion(name)]));

  test('every pack version is a semver', () => {
    for (const version of Object.values(versions)) expect(version).toMatch(/^\d+\.\d+\.\d+$/);
  });

  test('every publishable format binding pins all three packs at the committed versions', () => {
    for (const name of PYTHON_PUBLISH_NAMES) {
      const pyproject = Bun.TOML.parse(
        readFileSync(join(repository, `bindings/python-${name}/pyproject.toml`), 'utf8')
      ) as any;
      expect(pyproject.project.dependencies).toEqual(
        FONT_PACK_NAMES.map((pack) => `betteroffice-${pack}==${versions[pack]}`)
      );
    }
  });

  test('the builder writes a py3-none-any wheel per pack with every face and the font licenses', () => {
    const out = mkdtempSync(join(tmpdir(), 'font-packs-'));
    try {
      const result = spawnSync('python3', ['-I', 'scripts/python_font_packs.py', '--out', out], {
        cwd: repository,
        encoding: 'utf8'
      });
      expect(result.status, result.stderr).toBe(0);
      for (const name of FONT_PACK_NAMES) {
        const pack = name.replace('fonts-', '');
        const version = versions[name];
        const wheel = `betteroffice_fonts_${pack}-${version}-py3-none-any.whl`;
        expect(result.stdout).toContain(wheel);
        const listing = spawnSync(
          'python3',
          ['-I', '-c', 'import sys, zipfile; print("\\n".join(zipfile.ZipFile(sys.argv[1]).namelist()))', join(out, wheel)],
          { encoding: 'utf8' }
        ).stdout.trim().split('\n');
        const manifest = JSON.parse(
          readFileSync(join(repository, `packages/${name}/assets/manifest.json`), 'utf8')
        );
        const data = listing.filter((entry) => entry.startsWith('betteroffice_fonts_packs/'));
        expect(data.sort()).toEqual(
          ['manifest.json', ...manifest.faces.map((face: { file: string }) => face.file)]
            .map((file) => `betteroffice_fonts_packs/${pack}/${version}/${file}`)
            .sort()
        );
        const info = `betteroffice_fonts_${pack}-${version}.dist-info`;
        for (const file of ['METADATA', 'WHEEL', 'RECORD', 'licenses/LICENSE']) {
          expect(listing).toContain(`${info}/${file}`);
        }
        expect(listing.some((entry) => /licenses\/LICENSES\/OFL-.+\.txt$/.test(entry))).toBe(true);
      }
    } finally {
      rmSync(out, { recursive: true, force: true });
    }
  }, 60_000);
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
    Object.fromEntries([
      ...FONT_PACK_NAMES.map((name) => [`betteroffice-${name}`, [fontPackVersion(name)]]),
      ...PYTHON_BINDINGS.map((path) => [
        `betteroffice-${path.replace('bindings/python-', '')}`,
        [bindingVersion(path)]
      ])
    ]);

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
    expect(requested).toEqual(
      [...FONT_PACK_NAMES.map((name) => `betteroffice-${name}`), ...PYPI_DISTRIBUTIONS].map(
        (d) => `https://pypi.org/pypi/${d}/json`
      )
    );
  });

  test('a font pack version PyPI does not serve is pending, ahead of the bindings', async () => {
    const releases = live();
    releases['betteroffice-fonts-cjk'] = null;
    releases['betteroffice-docx'] = ['0.0.0'];
    const { fetchImpl } = pypi(releases);
    expect(await pendingPublishNames({ fetchImpl })).toEqual(['fonts-cjk', 'docx']);
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
    let calls = 0;
    const fetchImpl = async () => {
      calls += 1;
      return new Response('unavailable', { status: 503 });
    };
    await expect(
      pendingPublishNames({ fetchImpl, attempts: 2, retryDelayMs: 0 })
    ).rejects.toThrow('cannot read PyPI');
    expect(calls).toBe(2);
  });

  test.each(['HTTP 503', 'connection reset'])('retries a transient PyPI %s', async (failure) => {
    const requested: string[] = [];
    const { fetchImpl: good } = pypi(live());
    const fetchImpl = async (url: string) => {
      requested.push(url);
      if (requested.length === 1) {
        if (failure === 'connection reset') throw new Error('reset');
        return new Response('unavailable', { status: 503 });
      }
      return good(url);
    };
    expect(await pendingPublishNames({ fetchImpl, retryDelayMs: 0 })).toEqual([]);
    expect(requested[1]).toBe(requested[0]);
    expect(requested).toHaveLength(ALL_NAMES.length + 1);
  });
});

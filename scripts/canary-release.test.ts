import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { canarySet, keepDependencyRanges, nativeDependents, pinToCanaries } from './canary-release.mjs';
import { NATIVE_PACKAGES } from './native-bindings.mjs';
import { workspaceEntries } from './published-packages.mjs';

const release = Bun.YAML.parse(
  readFileSync(fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url)), 'utf8')
) as any;
const canary = release.jobs.canary;
const step = (name: string) => canary.steps.find((candidate: any) => candidate.name === name);
const runs = () => canary.steps.map((candidate: any) => candidate.run ?? '').join('\n');

function entry(name: string, version: string, extra: Record<string, unknown> = {}) {
  return { directory: `/repo/${name}`, manifest: { name, version, ...extra } };
}

describe('canary workflow', () => {
  test('runs daily only when CANARY_RELEASE opts in, or on a dispatch that asks for it, from main', () => {
    expect(release.on.schedule).toEqual([{ cron: '17 3 * * *' }]);
    expect(release.on.workflow_dispatch.inputs.canary).toEqual({
      description: 'Build an npm canary from the newest main commit with green CI instead of releasing.',
      type: 'boolean',
      default: false
    });
    expect(release.on.workflow_dispatch.inputs.dry_run.default).toBe(true);
    expect(canary.if).toBe([
      "github.repository == 'openooxml/betteroffice' && github.ref == 'refs/heads/main' &&",
      "((github.event_name == 'schedule' && vars.CANARY_RELEASE == 'true') ||",
      "(github.event_name == 'workflow_dispatch' && inputs.canary == true))"
    ].join(' '));
    expect(canary.env.DRY_RUN).toBe("${{ github.event_name == 'workflow_dispatch' && inputs.dry_run != false }}");
  });

  test('never shares a concurrency slot with, or starts, the stable release', () => {
    expect(release.concurrency.group).toStartWith(
      "${{ (github.event_name == 'schedule' || inputs.canary == true) && 'release-canary' || "
    );
    for (const id of ['plan', 'release']) {
      expect(release.jobs[id].if, id).toContain("github.event_name != 'schedule' && inputs.canary != true");
    }
    expect(canary.needs).toBeUndefined();
    for (const [id, job] of Object.entries<any>(release.jobs)) {
      if (id !== 'canary') expect([job.needs ?? []].flat(), id).not.toContain('canary');
    }
    const checks = canary.steps.filter((candidate: any) => candidate.run?.includes('actions/workflows/release.yml/runs?status=in_progress'));
    expect(checks.map((candidate: any) => candidate.id)).toEqual(['idle', 'still-idle']);
  });

  test('has only read access, an OIDC token, and no secret at all', () => {
    expect(canary.permissions).toEqual({ contents: 'read', actions: 'read', 'id-token': 'write' });
    const source = JSON.stringify(canary);
    expect(source).not.toContain('secrets.');
    expect(source).not.toContain('NODE_AUTH_TOKEN');
    expect(source).not.toContain('registry-url');
    const checkout = canary.steps.find((candidate: any) => candidate.uses?.startsWith('actions/checkout@'));
    expect(checkout.with).toEqual({ ref: '${{ steps.commit.outputs.sha }}', 'persist-credentials': false });
  });

  test('writes nothing to the repository and never runs Changesets publishing', () => {
    for (const candidate of canary.steps) {
      expect(candidate.uses ?? '').not.toMatch(/^changesets\/action@/);
    }
    expect(runs()).not.toMatch(/\bgit\s+(?:push|commit|tag)\b|\bgh\s+(?:pr|release|workflow)\b|\bchangeset\s+publish\b|\bnpm\s+dist-tag\b/);
  });

  test('builds from the newest main commit with green CI and publishes only outside a dry run', () => {
    expect(step('Pick the newest main commit with green CI').run).toContain(
      'actions/workflows/ci.yml/runs?branch=main&event=push&status=success&per_page=1'
    );
    const publishers = canary.steps.filter((candidate: any) => / publish "/.test(candidate.run ?? ''));
    expect(publishers.map((candidate: any) => candidate.name)).toEqual(['Publish the canary']);
    expect(step('Publish the canary').if).toBe("steps.still-idle.outputs.idle == 'true' && env.DRY_RUN != 'true'");
    expect(step('Check again that no release is running').if).toBe(
      "steps.published.outputs.fresh == 'true' && env.DRY_RUN != 'true'"
    );
    expect(step('Install the canary tarballs in an empty project').if).toBe("steps.published.outputs.fresh == 'true'");
    expect(step('Upload the canary tarballs').if).toBe("steps.published.outputs.fresh == 'true'");
  });
});

describe('canary set', () => {
  const natives = new Set(['@x/a-native']);
  const workspaces = [
    entry('@x/a-native', '1.1.0-canary.1', { private: true }),
    entry('@x/agents', '1.1.0-canary.1', { dependencies: { '@x/a': 'workspace:^', '@x/a-native': 'workspace:^' } }),
    entry('@x/a', '1.1.0-canary.1', { dependencies: { '@x/fonts': 'workspace:^' } }),
    entry('@x/a-react', '1.1.0-canary.1', { dependencies: { '@x/a': 'workspace:*', '@x/a-i18n': 'workspace:*' } }),
    entry('@x/a-i18n', '1.1.0-canary.1'),
    entry('@x/fonts', '1.0.0'),
    entry('@x/hidden', '1.1.0-canary.1', { private: true }),
    entry('@x/tool', '1.1.0-canary.1', { optionalDependencies: { '@x/agents': 'workspace:^' } })
  ];

  test('leaves out native packages and everything that reaches one', () => {
    expect([...nativeDependents(workspaces, natives)].sort()).toEqual(['@x/agents', '@x/tool']);
    expect(canarySet(workspaces, natives).map(({ name }) => name)).toEqual(['@x/a', '@x/a-i18n', '@x/a-react']);
  });

  test('orders dependencies before their dependents', () => {
    const names = canarySet(workspaces, natives).map(({ name }) => name);
    expect(names.indexOf('@x/a')).toBeLessThan(names.indexOf('@x/a-react'));
    expect(names.indexOf('@x/a-i18n')).toBeLessThan(names.indexOf('@x/a-react'));
  });

  test('pins members exactly, including explicit ranges a prerelease would not satisfy', () => {
    const set = [{ name: '@x/fonts-cjk', version: '0.5.1-canary.20261127031700', directory: '' }];
    const manifest = {
      name: '@x/fonts',
      dependencies: { '@x/fonts-latin': '0.5.0' },
      peerDependencies: { '@x/fonts-cjk': '^0', react: '^19' }
    };
    expect(pinToCanaries(manifest, set)).toEqual({
      name: '@x/fonts',
      dependencies: { '@x/fonts-latin': '0.5.0' },
      peerDependencies: { '@x/fonts-cjk': '0.5.1-canary.20261127031700', react: '^19' }
    });
    expect(manifest.peerDependencies['@x/fonts-cjk']).toBe('^0');
  });

  test('keeps the dependency ranges Changesets rewrote, so only versions change', () => {
    const before = { name: '@x/a', version: '1.0.0', dependencies: { '@x/b': 'workspace:^0' } };
    const after = { name: '@x/a', version: '1.1.0-canary.1', dependencies: { '@x/b': 'workspace:^1.1.0-canary.1' } };
    expect(keepDependencyRanges(before, after)).toEqual({ ...after, dependencies: { '@x/b': 'workspace:^0' } });
  });

  test('in this repository: the format, React, i18n and font packages, without native roots or agents', () => {
    const roots = new Set(NATIVE_PACKAGES.map((path) =>
      JSON.parse(readFileSync(fileURLToPath(new URL(`../${path}/package.json`, import.meta.url)), 'utf8')).name));
    const everyVersionACanary = workspaceEntries().map(({ directory, manifest }) => ({
      directory,
      manifest: { ...manifest, version: `${manifest.version}-canary.1` }
    }));
    expect(canarySet(everyVersionACanary, roots).map(({ name }) => name).sort()).toEqual([
      '@betteroffice/docx', '@betteroffice/docx-i18n', '@betteroffice/docx-react',
      '@betteroffice/fonts', '@betteroffice/fonts-cjk', '@betteroffice/fonts-intl', '@betteroffice/fonts-latin',
      '@betteroffice/pptx', '@betteroffice/pptx-i18n', '@betteroffice/pptx-react',
      '@betteroffice/xlsx', '@betteroffice/xlsx-i18n', '@betteroffice/xlsx-react'
    ]);
  });
});

import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  attemptPublishWithFallback,
  cargoNoAuthEnv,
  cargoPublishEnv,
  formatBootstrapSummary,
  publishOneCrate,
  selectPublishToken
} from './publish-crates.mjs';
import { run } from './rust-crates.mjs';

const FAKE_OIDC = 'fake-oidc-token-aaa';
const FAKE_BOOTSTRAP = 'fake-bootstrap-token-bbb';
const FAKE_CONFLICT = 'fake-conflicting-token-ccc';
const FAKE_SELECTED = 'fake-selected-token-ddd';

function publishTracker(script: Array<{ token: string; status: number }>) {
  const calls: Array<{ token: string; source: string }> = [];
  let index = 0;
  return {
    calls,
    runPublish: (token: string, source: string) => {
      calls.push({ token, source });
      const step = script[Math.min(index, script.length - 1)]!;
      index += 1;
      return { status: step.status };
    }
  };
}

describe('selectPublishToken', () => {
  test('a crate on crates.io publishes with the OIDC token', () => {
    expect(
      selectPublishToken({
        name: 'betteroffice-xlsx',
        exists: true,
        oidcToken: 'oidc',
        bootstrapToken: 'bootstrap'
      })
    ).toEqual({ token: 'oidc', source: 'oidc' });
  });

  test('a crate missing from crates.io is created with the bootstrap token', () => {
    expect(
      selectPublishToken({
        name: 'betteroffice-pptx-raster',
        exists: false,
        oidcToken: 'oidc',
        bootstrapToken: 'bootstrap'
      })
    ).toEqual({ token: 'bootstrap', source: 'bootstrap' });
  });

  test('a new crate goes straight to the bootstrap token without using OIDC', () => {
    const selected = selectPublishToken({
      name: 'betteroffice-pptx-raster',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP
    });
    expect(selected).toEqual({ token: FAKE_BOOTSTRAP, source: 'bootstrap' });
    expect(selected.token).not.toBe(FAKE_OIDC);
  });

  test('an existing crate without an OIDC token falls back directly to bootstrap', () => {
    expect(
      selectPublishToken({
        name: 'betteroffice-xlsx',
        exists: true,
        oidcToken: '',
        bootstrapToken: FAKE_BOOTSTRAP
      })
    ).toEqual({ token: FAKE_BOOTSTRAP, source: 'bootstrap-fallback' });
  });

  test('a missing crate without a bootstrap token names the crate', () => {
    expect(() =>
      selectPublishToken({
        name: 'betteroffice-pptx-raster',
        exists: false,
        oidcToken: 'oidc',
        bootstrapToken: ''
      })
    ).toThrow('betteroffice-pptx-raster');
  });

  test('a missing crate without a bootstrap token says OIDC cannot create it', () => {
    expect(() =>
      selectPublishToken({
        name: 'betteroffice-pptx-raster',
        exists: false,
        oidcToken: 'oidc',
        bootstrapToken: undefined
      })
    ).toThrow('OIDC cannot create a crate');
  });

  test('a present crate without either token fails naming both', () => {
    expect(() =>
      selectPublishToken({
        name: 'betteroffice-xlsx',
        exists: true,
        oidcToken: '',
        bootstrapToken: ''
      })
    ).toThrow('CARGO_REGISTRY_TOKEN');
  });
});

describe('attemptPublishWithFallback', () => {
  test('existing OIDC success never attempts the bootstrap token', async () => {
    const tracker = publishTracker([{ token: FAKE_OIDC, status: 0 }]);
    const seen: unknown[] = [];
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-xlsx',
      version: '0.1.0',
      exists: true,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => {
        seen.push('check');
        return null;
      },
      runPublish: tracker.runPublish
    });
    expect(result).toEqual({ source: 'oidc' });
    expect(tracker.calls).toEqual([{ token: FAKE_OIDC, source: 'oidc' }]);
    expect(seen).toEqual([]);
  });

  test('a new crate attempts only the bootstrap token', async () => {
    const tracker = publishTracker([{ token: FAKE_BOOTSTRAP, status: 0 }]);
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-pptx-raster',
      version: '0.1.0',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: tracker.runPublish
    });
    expect(result).toEqual({ source: 'bootstrap' });
    expect(tracker.calls).toEqual([{ token: FAKE_BOOTSTRAP, source: 'bootstrap' }]);
  });

  test('OIDC failure then bootstrap success falls back once', async () => {
    const tracker = publishTracker([
      { token: FAKE_OIDC, status: 1 },
      { token: FAKE_BOOTSTRAP, status: 0 }
    ]);
    let checks = 0;
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-xlsx',
      version: '0.1.0',
      exists: true,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => {
        checks += 1;
        return null;
      },
      runPublish: tracker.runPublish
    });
    expect(result).toEqual({ source: 'bootstrap-fallback' });
    expect(tracker.calls).toEqual([
      { token: FAKE_OIDC, source: 'oidc' },
      { token: FAKE_BOOTSTRAP, source: 'bootstrap-fallback' }
    ]);
    expect(checks).toBe(1);
  });

  test('missing OIDC with bootstrap goes directly to the fallback', async () => {
    const tracker = publishTracker([{ token: FAKE_BOOTSTRAP, status: 0 }]);
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-xlsx',
      version: '0.1.0',
      exists: true,
      oidcToken: '',
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: tracker.runPublish
    });
    expect(result).toEqual({ source: 'bootstrap-fallback' });
    expect(tracker.calls).toEqual([{ token: FAKE_BOOTSTRAP, source: 'bootstrap-fallback' }]);
  });

  test('OIDC failure without a fallback token fails clearly', async () => {
    const tracker = publishTracker([{ token: FAKE_OIDC, status: 1 }]);
    await expect(
      attemptPublishWithFallback({
        name: 'betteroffice-xlsx',
        version: '0.1.0',
        exists: true,
        oidcToken: FAKE_OIDC,
        bootstrapToken: '',
        checkVersion: async () => null,
        runPublish: tracker.runPublish
      })
    ).rejects.toThrow('CRATES_IO_BOOTSTRAP_TOKEN');
    expect(tracker.calls).toHaveLength(1);
  });

  test('both tokens missing for an existing crate fails without publishing', async () => {
    const tracker = publishTracker([]);
    await expect(
      attemptPublishWithFallback({
        name: 'betteroffice-xlsx',
        version: '0.1.0',
        exists: true,
        oidcToken: '',
        bootstrapToken: '',
        checkVersion: async () => null,
        runPublish: tracker.runPublish
      })
    ).rejects.toThrow('both missing');
    expect(tracker.calls).toHaveLength(0);
  });

  test('fallback failure reports both credentials without hiding the OIDC error', async () => {
    const tracker = publishTracker([
      { token: FAKE_OIDC, status: 1 },
      { token: FAKE_BOOTSTRAP, status: 1 }
    ]);
    await expect(
      attemptPublishWithFallback({
        name: 'betteroffice-xlsx',
        version: '0.1.0',
        exists: true,
        oidcToken: FAKE_OIDC,
        bootstrapToken: FAKE_BOOTSTRAP,
        checkVersion: async () => null,
        runPublish: tracker.runPublish
      })
    ).rejects.toThrow('OIDC and with the bootstrap token');
    expect(tracker.calls).toHaveLength(2);
  });

  test('a version visible after a failed OIDC attempt skips fallback without attributing it', async () => {
    const tracker = publishTracker([{ token: FAKE_OIDC, status: 1 }]);
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-xlsx',
      version: '0.1.0',
      exists: true,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => ({ vers: '0.1.0', yanked: false } as never),
      runPublish: tracker.runPublish
    });
    expect(result).toEqual({ source: 'unknown' });
    expect(tracker.calls).toHaveLength(1);
  });

  test('per-crate choice resets: fallback on one crate keeps OIDC on the next', async () => {
    const first = publishTracker([
      { token: FAKE_OIDC, status: 1 },
      { token: FAKE_BOOTSTRAP, status: 0 }
    ]);
    const firstResult = await attemptPublishWithFallback({
      name: 'crate-a',
      version: '0.1.0',
      exists: true,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: first.runPublish
    });
    expect(firstResult).toEqual({ source: 'bootstrap-fallback' });

    const second = publishTracker([{ token: FAKE_OIDC, status: 0 }]);
    const secondResult = await attemptPublishWithFallback({
      name: 'crate-b',
      version: '0.1.0',
      exists: true,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: second.runPublish
    });
    expect(secondResult).toEqual({ source: 'oidc' });
    expect(second.calls).toEqual([{ token: FAKE_OIDC, source: 'oidc' }]);
  });

  test('a registry lookup error fails instead of inferring missing', async () => {
    const tracker = publishTracker([{ token: FAKE_OIDC, status: 1 }]);
    const lookupError = new Error('registry unavailable');
    await expect(
      attemptPublishWithFallback({
        name: 'betteroffice-xlsx',
        version: '0.1.0',
        exists: true,
        oidcToken: FAKE_OIDC,
        bootstrapToken: FAKE_BOOTSTRAP,
        checkVersion: async () => {
          throw lookupError;
        },
        runPublish: tracker.runPublish
      })
    ).rejects.toBe(lookupError);
    expect(tracker.calls).toHaveLength(1);
  });

  test('a yanked version after a failed upload is rejected, not retried', async () => {
    const tracker = publishTracker([{ token: FAKE_OIDC, status: 1 }]);
    await expect(
      attemptPublishWithFallback({
        name: 'betteroffice-xlsx',
        version: '0.1.0',
        exists: true,
        oidcToken: FAKE_OIDC,
        bootstrapToken: FAKE_BOOTSTRAP,
        checkVersion: async () => ({ vers: '0.1.0', yanked: true } as never),
        runPublish: tracker.runPublish
      })
    ).rejects.toThrow('yanked');
    expect(tracker.calls).toHaveLength(1);
  });
});

describe('formatBootstrapSummary', () => {
  test('lists the created crates and the Trusted Publisher step', () => {
    const summary = formatBootstrapSummary(['betteroffice-pptx-raster', 'betteroffice-pptx']);
    expect(summary).toContain('betteroffice-pptx-raster');
    expect(summary).toContain('betteroffice-pptx');
    expect(summary).toContain('openooxml');
    expect(summary).toContain('betteroffice');
    expect(summary).toContain('release.yml');
    expect(summary).toContain('CRATES_IO_BOOTSTRAP_TOKEN');
  });

  test('differentiates new crates from existing fallback, each with its reminder', () => {
    const summary = formatBootstrapSummary(['new-crate'], ['existing-crate']);
    expect(summary).toContain('new-crate');
    expect(summary).toContain('existing-crate');
    expect(summary).toContain('Created with CRATES_IO_BOOTSTRAP_TOKEN');
    expect(summary).toContain('fallback');
    const createdLine = summary.split('\n').find((line) => line.startsWith('Created with'))!;
    expect(createdLine).not.toContain('existing-crate');
    expect(createdLine).toContain('Trusted Publisher');
    const fallbackLine = summary.split('\n').find((line) => line.includes('fallback'))!;
    expect(fallbackLine).toContain('existing-crate');
    expect(fallbackLine).not.toContain('new-crate');
    expect(fallbackLine).toContain('Trusted Publisher');
  });
});

describe('publishOneCrate', () => {
  test.each([false, true])(
    'delayed visibility after a failed fallback stays unattributed (registry wait fails: %s)',
    async (waitFails) => {
      const dir = mkdtempSync(join(tmpdir(), 'crates-auth-'));
      const summaryFile = join(dir, 'summary.md');
      const previous = process.env.GITHUB_STEP_SUMMARY;
      process.env.GITHUB_STEP_SUMMARY = summaryFile;
      try {
        const tracker = publishTracker([
          { token: FAKE_OIDC, status: 1 },
          { token: FAKE_BOOTSTRAP, status: 1 }
        ]);
        let checks = 0;
        let waited = false;
        const publication = publishOneCrate({
          name: 'delayed-crate',
          version: '0.1.0',
          exists: true,
          oidcToken: FAKE_OIDC,
          bootstrapToken: FAKE_BOOTSTRAP,
          checkVersion: async () => (++checks === 1 ? null : { yanked: false }),
          runPublish: tracker.runPublish,
          waitRegistry: async () => {
            waited = true;
            expect(readFileSync(summaryFile, 'utf8')).toContain('credential could not be confirmed');
            if (waitFails) throw new Error('registry wait failed');
          }
        });
        if (waitFails) await expect(publication).rejects.toThrow('registry wait failed');
        else expect(await publication).toEqual({ source: 'unknown' });
        expect(checks).toBe(2);
        expect(waited).toBe(true);
        expect(tracker.calls).toEqual([
          { token: FAKE_OIDC, source: 'oidc' },
          { token: FAKE_BOOTSTRAP, source: 'bootstrap-fallback' }
        ]);
        const summary = readFileSync(summaryFile, 'utf8');
        expect(summary).toContain('delayed-crate@0.1.0');
        expect(summary).toContain('credential could not be confirmed');
        expect(summary).not.toContain('CRATES_IO_BOOTSTRAP_TOKEN');
        expect(summary).not.toContain('Trusted Publisher');
      } finally {
        if (previous === undefined) delete process.env.GITHUB_STEP_SUMMARY;
        else process.env.GITHUB_STEP_SUMMARY = previous;
        rmSync(dir, { recursive: true, force: true });
      }
    }
  );

  test('a failing registry wait retains the already recorded fallback use', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'crates-auth-'));
    const summaryFile = join(dir, 'summary.md');
    const previous = process.env.GITHUB_STEP_SUMMARY;
    process.env.GITHUB_STEP_SUMMARY = summaryFile;
    try {
      const tracker = publishTracker([
        { token: FAKE_OIDC, status: 1 },
        { token: FAKE_BOOTSTRAP, status: 0 }
      ]);
      await expect(
        publishOneCrate({
          name: 'first-crate',
          version: '0.1.0',
          exists: true,
          oidcToken: FAKE_OIDC,
          bootstrapToken: FAKE_BOOTSTRAP,
          checkVersion: async () => null,
          runPublish: tracker.runPublish,
          waitRegistry: async () => {
            throw new Error('registry wait failed');
          }
        })
      ).rejects.toThrow('registry wait failed');
      expect(tracker.calls).toEqual([
        { token: FAKE_OIDC, source: 'oidc' },
        { token: FAKE_BOOTSTRAP, source: 'bootstrap-fallback' }
      ]);
      const persisted = readFileSync(summaryFile, 'utf8');
      expect(persisted).toContain('first-crate');
      expect(persisted).toContain('CRATES_IO_BOOTSTRAP_TOKEN');
      expect(persisted).toContain('fallback');
      expect(persisted).toContain('Trusted Publisher');
    } finally {
      if (previous === undefined) delete process.env.GITHUB_STEP_SUMMARY;
      else process.env.GITHUB_STEP_SUMMARY = previous;
    }
  });

  test('no success is recorded when both attempts fail and the version is absent', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'crates-auth-'));
    const summaryFile = join(dir, 'summary.md');
    const previous = process.env.GITHUB_STEP_SUMMARY;
    process.env.GITHUB_STEP_SUMMARY = summaryFile;
    try {
      const tracker = publishTracker([
        { token: FAKE_OIDC, status: 1 },
        { token: FAKE_BOOTSTRAP, status: 1 }
      ]);
      let waited = false;
      await expect(
        publishOneCrate({
          name: 'betteroffice-xlsx',
          version: '0.1.0',
          exists: true,
          oidcToken: FAKE_OIDC,
          bootstrapToken: FAKE_BOOTSTRAP,
          checkVersion: async () => null,
          runPublish: tracker.runPublish,
          waitRegistry: async () => {
            waited = true;
          }
        })
      ).rejects.toThrow('OIDC and with the bootstrap token');
      expect(tracker.calls).toHaveLength(2);
      expect(waited).toBe(false);
      let persisted = '';
      try {
        persisted = readFileSync(summaryFile, 'utf8');
      } catch {
        persisted = '';
      }
      expect(persisted).not.toContain('betteroffice-xlsx');
    } finally {
      if (previous === undefined) delete process.env.GITHUB_STEP_SUMMARY;
      else process.env.GITHUB_STEP_SUMMARY = previous;
    }
  });
});

describe('cargo child environment isolation', () => {
  test('the publish child carries only the selected token under both aliases', () => {
    const saved = {
      bootstrap: process.env.CRATES_IO_BOOTSTRAP_TOKEN,
      reg: process.env.CARGO_REGISTRY_TOKEN,
      named: process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN
    };
    process.env.CRATES_IO_BOOTSTRAP_TOKEN = FAKE_BOOTSTRAP;
    process.env.CARGO_REGISTRY_TOKEN = FAKE_OIDC;
    process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = FAKE_CONFLICT;
    try {
      const result = run(
        'node',
        [
          '-e',
          'process.stdout.write(JSON.stringify({bootstrap: process.env.CRATES_IO_BOOTSTRAP_TOKEN ?? null, reg: process.env.CARGO_REGISTRY_TOKEN ?? null, named: process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN ?? null}))'
        ],
        { capture: true, env: cargoPublishEnv(FAKE_SELECTED) }
      );
      const child = JSON.parse(result.stdout as string);
      expect(child.bootstrap).toBeNull();
      expect(child.reg).toBe(FAKE_SELECTED);
      expect(child.named).toBe(FAKE_SELECTED);
    } finally {
      if (saved.bootstrap === undefined) delete process.env.CRATES_IO_BOOTSTRAP_TOKEN;
      else process.env.CRATES_IO_BOOTSTRAP_TOKEN = saved.bootstrap;
      if (saved.reg === undefined) delete process.env.CARGO_REGISTRY_TOKEN;
      else process.env.CARGO_REGISTRY_TOKEN = saved.reg;
      if (saved.named === undefined) delete process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN;
      else process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = saved.named;
    }
  });

  test('a conflicting named token cannot override the selected token', () => {
    const saved = process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN;
    process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = FAKE_CONFLICT;
    try {
      const result = run(
        'node',
        [
          '-e',
          'process.stdout.write(JSON.stringify({reg: process.env.CARGO_REGISTRY_TOKEN ?? null, named: process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN ?? null}))'
        ],
        { capture: true, env: cargoPublishEnv(FAKE_OIDC) }
      );
      const child = JSON.parse(result.stdout as string);
      expect(child.reg).toBe(FAKE_OIDC);
      expect(child.named).toBe(FAKE_OIDC);
      expect(child.named).not.toBe(FAKE_CONFLICT);
    } finally {
      if (saved === undefined) delete process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN;
      else process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = saved;
    }
  });

  test('metadata and package children carry no registry secrets', () => {
    const saved = {
      bootstrap: process.env.CRATES_IO_BOOTSTRAP_TOKEN,
      reg: process.env.CARGO_REGISTRY_TOKEN,
      named: process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN
    };
    process.env.CRATES_IO_BOOTSTRAP_TOKEN = FAKE_BOOTSTRAP;
    process.env.CARGO_REGISTRY_TOKEN = FAKE_OIDC;
    process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = FAKE_CONFLICT;
    try {
      const result = run(
        'node',
        [
          '-e',
          'process.stdout.write(JSON.stringify({bootstrap: process.env.CRATES_IO_BOOTSTRAP_TOKEN ?? null, reg: process.env.CARGO_REGISTRY_TOKEN ?? null, named: process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN ?? null}))'
        ],
        { capture: true, env: cargoNoAuthEnv() }
      );
      const child = JSON.parse(result.stdout as string);
      expect(child.bootstrap).toBeNull();
      expect(child.reg).toBeNull();
      expect(child.named).toBeNull();
    } finally {
      if (saved.bootstrap === undefined) delete process.env.CRATES_IO_BOOTSTRAP_TOKEN;
      else process.env.CRATES_IO_BOOTSTRAP_TOKEN = saved.bootstrap;
      if (saved.reg === undefined) delete process.env.CARGO_REGISTRY_TOKEN;
      else process.env.CARGO_REGISTRY_TOKEN = saved.reg;
      if (saved.named === undefined) delete process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN;
      else process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = saved.named;
    }
  });

});

describe('release workflow crates auth', () => {
  const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));
  const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
  const named = new Map(release.jobs.release.steps.map((step: any) => [step.name, step]));

  test('the OIDC action is pinned to the official v1 commit', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth.uses).toBe(
      'rust-lang/crates-io-auth-action@c6f97d42243bad5fab37ca0427f495c86d5b1a18'
    );
  });

  test('the OIDC exchange always runs on the publish path', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth.if).toBe("steps.pending.outputs.publishing == 'true'");
    expect(auth.if).not.toContain('crates-bootstrap');
  });

  test('an OIDC failure continues only when the bootstrap token is detected', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth['continue-on-error']).toBe(
      "${{ steps.crates-bootstrap.outputs.enabled == 'true' }}"
    );
  });

  test('the crates publish still gets the OIDC token and bootstrap separately', () => {
    const publish = named.get('Publish Rust crates');
    expect(publish.env.CARGO_REGISTRY_TOKEN).toBe('${{ steps.crates-auth.outputs.token }}');
    expect(publish.env.CRATES_IO_BOOTSTRAP_TOKEN).toBe('${{ secrets.CRATES_IO_BOOTSTRAP_TOKEN }}');
    expect(publish.if).toBe("steps.pending.outputs.publishing == 'true'");
  });
});

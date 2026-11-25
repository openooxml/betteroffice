import { describe, expect, spyOn, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  NEW_CRATE_WAIT_BUDGET_MS,
  attemptPublishWithFallback,
  cargoNoAuthEnv,
  cargoPublishEnv,
  formatBootstrapSummary,
  newCrateRateLimitRetryAt,
  publishOneCrate,
  runWithTee,
  selectPublishToken
} from './publish-crates.mjs';
import { run } from './rust-crates.mjs';

const FAKE_OIDC = 'fake-oidc-token-aaa';
const FAKE_BOOTSTRAP = 'fake-bootstrap-token-bbb';
const FAKE_CONFLICT = 'fake-conflicting-token-ccc';
const FAKE_SELECTED = 'fake-selected-token-ddd';
const START_MS = Date.parse('2026-11-25T09:12:00Z');
const RETRY_MS = Date.parse('2026-11-25T09:20:00Z');
const NEW_CRATE_REFUSAL =
  'error: failed to publish to registry at https://crates.io\n' +
  'Caused by:\n' +
  '  the remote server responded with an error (status 429 Too Many Requests): ' +
  'You have published too many new crates in a short period of time.';

const CARGO_NEW_CRATE_REFUSAL = readFileSync(fileURLToPath(new URL('./publish-crates-429.log.txt', import.meta.url)), 'utf8');

function newCrateRefusal(retryAt: number) {
  return `${NEW_CRATE_REFUSAL} Please try again after ${new Date(retryAt).toUTCString()} and see https://crates.io/docs/rate-limits for more details.`;
}

function fakeClock() {
  let nowMs = START_MS;
  const waits: number[] = [];
  return {
    now: () => nowMs,
    sleep: async (ms: number) => {
      waits.push(ms);
      nowMs += ms;
    },
    waits
  };
}

function publishTracker(script: Array<{ token: string; status: number; output?: string; stdout?: string }>) {
  const calls: Array<{ token: string; source: string }> = [];
  let index = 0;
  return {
    calls,
    runPublish: (token: string, source: string) => {
      calls.push({ token, source });
      const step = script[Math.min(index, script.length - 1)]!;
      index += 1;
      return { status: step.status, stdout: step.stdout ?? '', stderr: step.output ?? '' };
    }
  };
}

describe('newCrateRateLimitRetryAt', () => {
  const cases: Array<[string, string, number | null]> = [
    ['a retry date', newCrateRefusal(RETRY_MS), RETRY_MS],
    ['a past retry date', newCrateRefusal(START_MS - 60_000), START_MS - 60_000],
    ['a missing retry date', NEW_CRATE_REFUSAL, START_MS + 600_000],
    ['cargo\'s refusal naming the crate and version', CARGO_NEW_CRATE_REFUSAL, Date.parse('2026-11-25T16:06:03Z')],
    ['an invalid retry date', `${NEW_CRATE_REFUSAL} Please try again after an invalid date GMT`, START_MS + 600_000],
    ['the phrase without the 429 status line', 'HTTP 429: too many new crates', null],
    ['the status and the phrase on different lines', `status 429 Too Many Requests\nwarning: too many new crates`, null],
    ['a different status', NEW_CRATE_REFUSAL.replace('status 429', 'status 500'), null],
    ['the update limit', NEW_CRATE_REFUSAL.replace('new crates', 'updates'), null],
    ['a generic 429', 'HTTP 429: Too Many Requests', null],
    ['unrelated text on a 429 line', 'warning: status 429 seen; too many new crates in this workspace', null],
    ['the refusal quoted inside another line', `note: ${NEW_CRATE_REFUSAL.split('\n')[2].trim()}`, null],
    ['the refusal line without cargo\'s registry error above it', NEW_CRATE_REFUSAL.split('\n')[2], null],
    [
      'a coloured refusal',
      newCrateRefusal(RETRY_MS)
        .replace('error:', '\u001b[1m\u001b[31merror\u001b[0m\u001b[1m:\u001b[0m')
        .replace('Caused by:', '\u001b[1mCaused by:\u001b[0m'),
      RETRY_MS
    ],
    ['an unauthorized error', 'error: unauthorized', null]
  ];
  test.each(cases)('handles %s', (_label, output, expected) => {
    expect(newCrateRateLimitRetryAt(output, START_MS)).toBe(expected);
  });
});

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

describe('new-crate rate limit', () => {
  test('retries the sixth and later crates with the same token after each refill', async () => {
    const clock = fakeClock();
    const names = Array.from({ length: 11 }, (_, index) => `new-crate-${index + 1}`);
    const rateLimitBudget = { waitedMs: 0, newCratesLeft: names.length };
    const calls: Array<{ name: string; token: string; source: string }> = [];
    let published = 0;
    let refillAt = RETRY_MS;
    const log = spyOn(console, 'log').mockImplementation(() => {});
    try {
      for (const name of names) {
        await publishOneCrate({
          name,
          version: '0.1.0',
          exists: false,
          oidcToken: FAKE_OIDC,
          bootstrapToken: FAKE_BOOTSTRAP,
          checkVersion: async () => null,
          runPublish: (token: string, source: string) => {
            calls.push({ name, token, source });
            if (published >= 5) {
              if (clock.now() < refillAt) return { status: 101, stderr: newCrateRefusal(refillAt) };
              refillAt += 600_000;
            }
            published += 1;
            return { status: 0, stderr: '' };
          },
          recordUse: () => {},
          waitRegistry: async () => {},
          rateLimitBudget,
          now: clock.now,
          sleep: clock.sleep
        });
      }
      expect(clock.waits).toEqual([495_000, 600_000, 600_000, 600_000, 600_000, 600_000]);
      expect(log.mock.calls.map(([line]) => line).filter((line) => line.startsWith('waiting ')))
        .toEqual([
          'waiting 495 s for the new-crate rate limit, 6 new crates left',
          'waiting 600 s for the new-crate rate limit, 5 new crates left',
          'waiting 600 s for the new-crate rate limit, 4 new crates left',
          'waiting 600 s for the new-crate rate limit, 3 new crates left',
          'waiting 600 s for the new-crate rate limit, 2 new crates left',
          'waiting 600 s for the new-crate rate limit, 1 new crates left'
        ]);
      expect(calls).toEqual(names.flatMap((name, index) => {
        const call = { name, token: FAKE_BOOTSTRAP, source: 'bootstrap' };
        return index < 5 ? [call] : [call, call];
      }));
      expect(rateLimitBudget.newCratesLeft).toBe(0);
    } finally {
      log.mockRestore();
    }
  });

  const marginCases: Array<[string, string, number]> = [
    ['a missing date', NEW_CRATE_REFUSAL, 615_000],
    ['a past date', newCrateRefusal(START_MS - 60_000), 15_000]
  ];
  test.each(marginCases)('applies the retry margin for %s', async (_label, output, waitMs) => {
    const clock = fakeClock();
    const tracker = publishTracker([
      { token: FAKE_BOOTSTRAP, status: 101, output },
      { token: FAKE_BOOTSTRAP, status: 0 }
    ]);
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-fonts-latin',
      version: '0.1.0',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: tracker.runPublish,
      now: clock.now,
      sleep: clock.sleep
    });
    expect(result).toEqual({ source: 'bootstrap' });
    expect(clock.waits).toEqual([waitMs]);
    expect(tracker.calls).toEqual([
      { token: FAKE_BOOTSTRAP, source: 'bootstrap' },
      { token: FAKE_BOOTSTRAP, source: 'bootstrap' }
    ]);
  });

  test('a cargo killed by a signal never waits, even after printing the refusal', async () => {
    const clock = fakeClock();
    await expect(attemptPublishWithFallback({
      name: 'betteroffice-fonts-latin',
      version: '0.1.0',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: () => ({ status: null, signal: 'SIGKILL', stderr: newCrateRefusal(START_MS + 60_000) }),
      now: clock.now,
      sleep: clock.sleep
    })).rejects.toThrow('Failed to publish betteroffice-fonts-latin@0.1.0 with the bootstrap token');
    expect(clock.waits).toEqual([]);
  });

  test('a refusal printed on stdout never waits', async () => {
    const clock = fakeClock();
    const tracker = publishTracker([
      { token: FAKE_BOOTSTRAP, status: 101, stdout: newCrateRefusal(START_MS + 60_000), output: 'error: something else' }
    ]);
    await expect(attemptPublishWithFallback({
      name: 'betteroffice-fonts-latin',
      version: '0.1.0',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: tracker.runPublish,
      now: clock.now,
      sleep: clock.sleep
    })).rejects.toThrow('Failed to publish betteroffice-fonts-latin@0.1.0 with the bootstrap token');
    expect(clock.waits).toEqual([]);
  });

  test('an existing crate never waits on the new-crate refusal text', async () => {
    const clock = fakeClock();
    const tracker = publishTracker([
      { token: FAKE_OIDC, status: 101, output: newCrateRefusal(START_MS + 60_000) },
      { token: FAKE_BOOTSTRAP, status: 0 }
    ]);
    const result = await attemptPublishWithFallback({
      name: 'betteroffice-xlsx',
      version: '0.1.0',
      exists: true,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      runPublish: tracker.runPublish,
      now: clock.now,
      sleep: clock.sleep
    });
    expect(result).toEqual({ source: 'bootstrap-fallback' });
    expect(clock.waits).toEqual([]);
  });

  const otherFailures: Array<[string, string]> = [
    ['the update limit', NEW_CRATE_REFUSAL.replace('new crates', 'updates')],
    ['a generic 429', 'HTTP 429: Too Many Requests'],
    ['an unauthorized error', 'error: unauthorized']
  ];
  test.each(otherFailures)('checks the version and falls back or fails without waiting for %s', async (_label, output) => {
    const clock = fakeClock();
    const tracker = publishTracker([
      { token: FAKE_OIDC, status: 101, output },
      { token: FAKE_BOOTSTRAP, status: 0 }
    ]);
    let checks = 0;
    const options = {
      name: 'crate-a',
      version: '0.1.0',
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => {
        checks += 1;
        return null;
      },
      now: clock.now,
      sleep: clock.sleep
    };
    expect(await attemptPublishWithFallback({
      ...options,
      exists: true,
      runPublish: tracker.runPublish
    })).toEqual({ source: 'bootstrap-fallback' });
    expect(tracker.calls).toEqual([
      { token: FAKE_OIDC, source: 'oidc' },
      { token: FAKE_BOOTSTRAP, source: 'bootstrap-fallback' }
    ]);
    await expect(attemptPublishWithFallback({
      ...options,
      exists: false,
      runPublish: () => ({ status: 101, stderr: output })
    })).rejects.toThrow('Failed to publish crate-a@0.1.0 with the bootstrap token');
    expect(checks).toBe(2);
    expect(clock.waits).toEqual([]);
  });

  test('shares the 100-minute cap across crates and stops repeated refusals', async () => {
    const clock = fakeClock();
    const rateLimitBudget = { waitedMs: 0, newCratesLeft: 2 };
    const first = publishTracker([
      { token: FAKE_BOOTSTRAP, status: 101, output: newCrateRefusal(START_MS + 5_970_000) },
      { token: FAKE_BOOTSTRAP, status: 0 }
    ]);
    const options = {
      version: '0.1.0',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      checkVersion: async () => null,
      recordUse: () => {},
      waitRegistry: async () => {},
      rateLimitBudget,
      now: clock.now,
      sleep: clock.sleep
    };
    await publishOneCrate({ ...options, name: 'first-crate', runPublish: first.runPublish });

    const second = publishTracker([
      { token: FAKE_BOOTSTRAP, status: 101, output: newCrateRefusal(START_MS + 5_985_000) },
      { token: FAKE_BOOTSTRAP, status: 101, output: newCrateRefusal(START_MS + 6_600_000) }
    ]);
    await expect(publishOneCrate({ ...options, name: 'second-crate', runPublish: second.runPublish }))
      .rejects.toThrow(
        /second-crate.*100 minutes.*1 new crates left.*2026-11-25T11:02:00\.000Z.*release job.*re-run.*published versions are skipped/i
      );
    expect(NEW_CRATE_WAIT_BUDGET_MS).toBe(6_000_000);
    expect(clock.waits).toEqual([5_985_000, 15_000]);
    expect(rateLimitBudget.waitedMs).toBe(NEW_CRATE_WAIT_BUDGET_MS);
    expect(second.calls).toHaveLength(2);
  });

  test('an OIDC token that expires during earlier waits falls back to bootstrap', async () => {
    const clock = fakeClock();
    const rateLimitBudget = { waitedMs: 0, newCratesLeft: 1 };
    const options = {
      version: '0.1.0',
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      recordUse: () => {},
      waitRegistry: async () => {},
      rateLimitBudget,
      now: clock.now,
      sleep: clock.sleep
    };
    await publishOneCrate({
      ...options,
      name: 'new-crate',
      exists: false,
      checkVersion: async () => null,
      runPublish: () => clock.now() < RETRY_MS
        ? { status: 101, stderr: newCrateRefusal(RETRY_MS) }
        : { status: 0, stderr: '' }
    });

    const tokens: string[] = [];
    let checks = 0;
    const result = await publishOneCrate({
      ...options,
      name: 'existing-crate',
      exists: true,
      checkVersion: async () => {
        checks += 1;
        return null;
      },
      runPublish: (token: string) => {
        tokens.push(token);
        return token === FAKE_OIDC && clock.now() >= RETRY_MS
          ? { status: 101, stderr: 'error: unauthorized, token expired' }
          : { status: 0, stderr: '' };
      }
    });
    expect(result).toEqual({ source: 'bootstrap-fallback' });
    expect(tokens).toEqual([FAKE_OIDC, FAKE_BOOTSTRAP]);
    expect(checks).toBe(1);
    expect(clock.waits).toEqual([495_000]);
  });
  test('cargo\'s real new-crate refusal waits until its retry time, then publishes', async () => {
    let nowMs = Date.parse('2026-11-25T16:05:24Z');
    const retryMs = Date.parse('2026-11-25T16:06:03Z');
    const waits: number[] = [];
    const attempts: string[] = [];
    const result = await publishOneCrate({
      name: 'betteroffice-fonts-cjk-kr',
      version: '0.6.0',
      exists: false,
      oidcToken: FAKE_OIDC,
      bootstrapToken: FAKE_BOOTSTRAP,
      recordUse: () => {},
      waitRegistry: async () => {},
      rateLimitBudget: { waitedMs: 0, newCratesLeft: 6 },
      now: () => nowMs,
      sleep: async (ms: number) => {
        waits.push(ms);
        nowMs += ms;
      },
      checkVersion: async () => null,
      runPublish: (token: string) => {
        attempts.push(token);
        return nowMs < retryMs ? { status: 101, stderr: CARGO_NEW_CRATE_REFUSAL } : { status: 0, stderr: '' };
      }
    });
    expect(result).toEqual({ source: 'bootstrap' });
    expect(attempts).toEqual([FAKE_BOOTSTRAP, FAKE_BOOTSTRAP]);
    expect(waits).toEqual([retryMs - Date.parse('2026-11-25T16:05:24Z') + 15_000]);
  });
});

describe('runWithTee', () => {
  test('returns within the drain time when a grandchild keeps the pipes open', async () => {
    const stdout = spyOn(process.stdout, 'write').mockImplementation(() => true);
    try {
      const started = Date.now();
      const result = await runWithTee('node', [
        '-e',
        "require('node:child_process').spawn(process.execPath, ['-e', 'setTimeout(() => {}, 20000)'], { stdio: 'inherit', detached: true }).unref(); process.stdout.write('parent done\\n'); process.exitCode = 7;"
      ], { drainMs: 500 });
      expect(Date.now() - started).toBeLessThan(5_000);
      expect(result.status).toBe(7);
      expect(result.stdout).toContain('parent done');
    } finally {
      stdout.mockRestore();
    }
  });

  test('tees stdout and stderr and returns the output and nonzero exit status', async () => {
    const stdout = spyOn(process.stdout, 'write').mockImplementation(() => true);
    const stderr = spyOn(process.stderr, 'write').mockImplementation(() => true);
    try {
      const result = await runWithTee('node', [
        '-e',
        "process.stdout.write('publish stdout\\n'); process.stderr.write('publish stderr\\n'); process.exitCode = 101;"
      ]);
      expect(result.status).toBe(101);
      expect(result.stdout).toContain('publish stdout\n');
      expect(result.stderr).toContain('publish stderr\n');
      expect(stdout.mock.calls.map(([chunk]) => String(chunk)).join('')).toContain('publish stdout\n');
      expect(stderr.mock.calls.map(([chunk]) => String(chunk)).join('')).toContain('publish stderr\n');
    } finally {
      stdout.mockRestore();
      stderr.mockRestore();
    }
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
  test('the publish child carries only the selected token under both aliases', async () => {
    const saved = {
      bootstrap: process.env.CRATES_IO_BOOTSTRAP_TOKEN,
      reg: process.env.CARGO_REGISTRY_TOKEN,
      named: process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN
    };
    process.env.CRATES_IO_BOOTSTRAP_TOKEN = FAKE_BOOTSTRAP;
    process.env.CARGO_REGISTRY_TOKEN = FAKE_OIDC;
    process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN = FAKE_CONFLICT;
    try {
      const result = await runWithTee(
        'node',
        [
          '-e',
          'const token = process.argv[1];' +
            'process.exitCode = process.env.CRATES_IO_BOOTSTRAP_TOKEN === undefined &&' +
            ' process.env.CARGO_REGISTRY_TOKEN === token &&' +
            ' process.env.CARGO_REGISTRIES_CRATES_IO_TOKEN === token ? 0 : 1;',
          FAKE_SELECTED
        ],
        { env: cargoPublishEnv(FAKE_SELECTED) }
      );
      expect(result.status).toBe(0);
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
  const releaseSource = readFileSync(releaseWorkflow, 'utf8');
  const release = Bun.YAML.parse(releaseSource) as any;
  const named = new Map(release.jobs.release.steps.map((step: any) => [step.name, step]));

  test('the release timeout has room for rate-limit waits', () => {
    expect(release.jobs.release['timeout-minutes']).toBeGreaterThanOrEqual(160);
  });

  test('the OIDC action is pinned to a full commit of the official v1 release line', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth.uses).toMatch(/^rust-lang\/crates-io-auth-action@[0-9a-f]{40}$/);
    const pins = releaseSource.split('\n').filter((line) => line.includes(`uses: ${auth.uses}`));
    expect(pins).toHaveLength(1);
    expect(pins[0]).toMatch(/ # v1$/);
  });

  test('the OIDC exchange runs on the publish path unless crates are skipped', () => {
    const auth = named.get('Authenticate to crates.io');
    expect(auth.if).toBe(
      "needs.plan.outputs.publishing == 'true' && vars.RELEASE_SKIP_CRATES != 'true'"
    );
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
    expect(publish.if).toBe(
      "needs.plan.outputs.publishing == 'true' && vars.RELEASE_SKIP_CRATES != 'true'"
    );
  });

  test('the type check and tests run unless RELEASE_SKIP_TESTS is true', () => {
    const job = release.jobs.test;
    expect(job.needs).toBe('plan');
    expect(job.if).toBe(
      "needs.plan.outputs.publishing == 'true' && vars.RELEASE_SKIP_TESTS != 'true' && github.ref == 'refs/heads/main' && inputs.rehearse != true"
    );
    expect(job['runs-on']).toBe('ubuntu-latest');
    expect(job['timeout-minutes']).toBe(100);
    expect(job.permissions).toEqual({ contents: 'read' });
    expect(job.steps.map((step: any) => step.uses ?? step.run)).toEqual([
      'actions/checkout@v7.0.1',
      'oven-sh/setup-bun@v2',
      'actions/setup-node@v7.1.0',
      'bun install --frozen-lockfile',
      'dtolnay/rust-toolchain@stable',
      'Swatinem/rust-cache@v2',
      './.github/actions/wasm-toolchain',
      'taiki-e/install-action@v2',
      'bun run typecheck:packages',
      'bun run test'
    ]);
    expect(job.steps[1].with['bun-version']).toBe('1.4.3');
    expect(job.steps[2].with['node-version']).toBe('24');
    expect(job.steps[4].with.targets).toBe('wasm32-unknown-unknown');
    expect(job.steps[7].with.tool).toBe('wasm-pack@0.15.0');
    for (const name of ['Type check', 'Test']) {
      expect(named.has(name)).toBe(false);
    }
  });

  test('RELEASE_SKIP_TESTS appears only in the test condition and the plan warning env', () => {
    const note = release.jobs.plan.steps.find((step: any) => step.name === 'Note skipped tests');
    expect(note.if).toBeUndefined();
    expect(note.env).toEqual({
      SKIP_TESTS: '${{ vars.RELEASE_SKIP_TESTS }}',
      PUBLISHING: '${{ steps.pending.outputs.publishing }}',
      REHEARSE: '${{ inputs.rehearse == true }}'
    });
    expect(release.jobs.test.if).toContain('RELEASE_SKIP_TESTS');
    for (const [id, job] of Object.entries<any>(release.jobs)) {
      if (id !== 'test') expect(job.if ?? '', id).not.toContain('RELEASE_SKIP_TESTS');
      for (const step of job.steps ?? []) {
        expect(step.if ?? '', id).not.toContain('RELEASE_SKIP_TESTS');
      }
    }
    expect(releaseSource.match(/RELEASE_SKIP_TESTS/g)).toHaveLength(2);
  });

  test.each([
    ['true', 'true', 'false', true],
    ['TRUE', 'true', 'false', true],
    ['True', 'true', 'false', true],
    ['false', 'true', 'false', false],
    ['true', 'false', 'false', false],
    ['true', 'true', 'true', false]
  ] as const)('plan warning with skip=%s, publishing=%s and rehearse=%s', (skip, publishing, rehearse, warns) => {
    const note = release.jobs.plan.steps.find((step: any) => step.name === 'Note skipped tests');
    const result = spawnSync('bash', ['-eo', 'pipefail', '-c', note.run], {
      encoding: 'utf8',
      env: { ...process.env, SKIP_TESTS: skip, PUBLISHING: publishing, REHEARSE: rehearse }
    });
    expect(result.status, result.stderr).toBe(0);
    expect(result.stdout).toBe(warns
      ? '::warning::RELEASE_SKIP_TESTS is true, so this release skips its type check and tests.\n'
      : '');
  });
});

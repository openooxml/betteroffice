import { describe, expect, test } from 'bun:test';
import { formatBootstrapSummary, selectPublishToken } from './publish-crates.mjs';

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

  test('a present crate without an OIDC token fails', () => {
    expect(() =>
      selectPublishToken({
        name: 'betteroffice-xlsx',
        exists: true,
        oidcToken: '',
        bootstrapToken: 'bootstrap'
      })
    ).toThrow('CARGO_REGISTRY_TOKEN');
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
});

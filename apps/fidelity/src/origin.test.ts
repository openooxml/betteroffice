import { expect, test } from 'bun:test';
import { sameOrigin } from './origin';

const page = 'https://benchmarks.betteroffice.dev/?report=x';

test('same-origin paths resolve to absolute URLs on this origin', () => {
  expect(sameOrigin('/renders/abc/report.json', page)).toBe(
    'https://benchmarks.betteroffice.dev/renders/abc/report.json'
  );
  expect(sameOrigin('local/e2e', page)).toBe('https://benchmarks.betteroffice.dev/local/e2e');
});

test('overrides that could reach another origin are refused or pinned here', () => {
  for (const value of ['https://evil.example/r.json', '//evil.example/r.json', 'javascript:alert(1)', 'http://['])
    expect(sameOrigin(value, page)).toBeNull();
  for (const value of ['/.//evil.example/r.json', '/..//evil.example/r.json', 'https://benchmarks.betteroffice.dev//evil.example/r.json'])
    expect(new URL(sameOrigin(value, page)!).host).toBe('benchmarks.betteroffice.dev');
  expect(sameOrigin(null, page)).toBeNull();
  expect(sameOrigin('', page)).toBeNull();
});

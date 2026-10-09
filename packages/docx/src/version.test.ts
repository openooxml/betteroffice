import { describe, expect, test } from 'bun:test';

import { version as packageVersion } from '../package.json';
import { VERSION } from './core';

describe('VERSION', () => {
  test('matches package.json', () => {
    expect(VERSION).toBe(packageVersion);
  });
});

import { getDefaultEnvironment } from '@modelcontextprotocol/sdk/client/stdio.js';

export function cliEnv(): Record<string, string> {
  const { BUN_OPTIONS } = process.env;
  return { ...getDefaultEnvironment(), ...(BUN_OPTIONS ? { BUN_OPTIONS } : {}) };
}

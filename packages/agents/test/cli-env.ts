import { getDefaultEnvironment } from '@modelcontextprotocol/sdk/client/stdio.js';

export function cliEnv(): Record<string, string> {
  const { BUN_OPTIONS } = process.env;
  return { ...getDefaultEnvironment(), ...(BUN_OPTIONS ? { BUN_OPTIONS } : {}) };
}

/** Child stderr without the one warning a run without native prebuilts prints. */
export function unexpectedStderr(stderr: string): string {
  return stderr.replace(/\(node:\d+\) \[BETTEROFFICE_AGENTS_NATIVE_UNAVAILABLE\][\s\S]*?(\(Use `\S+ --trace-warnings[^\n]*\n|$)/, '');
}

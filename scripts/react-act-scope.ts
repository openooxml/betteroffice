const INTERNALS = '__CLIENT_INTERNALS_DO_NOT_USE_OR_WARN_USERS_THEY_CANNOT_UPGRADE';
const REACT_ENTRY = /[\\/]node_modules[\\/]react[\\/](?:index|cjs[\\/]react\.(?:development|production))\.js$/;

export type ActScopeLog = { started: number; open: number; first?: number };

const field = (value: unknown, key: string): unknown =>
  (typeof value === 'object' || typeof value === 'function') && value !== null ? Reflect.get(value, key) : undefined;

export function openActQueueLength(cache: unknown): number | undefined {
  try {
    if (typeof cache !== 'object' || cache === null) return undefined;
    for (const [path, entry] of Object.entries(cache)) {
      if (!REACT_ENTRY.test(path)) continue;
      const queue = field(field(field(entry, 'exports'), INTERNALS), 'actQueue');
      if (Array.isArray(queue)) return queue.length;
    }
    return undefined;
  } catch {
    return undefined;
  }
}

export function openAtStartLine(testNumber: number, length: number): string {
  return `ACT-SCOPE-OPEN at the start of test #${testNumber} in this file: queue length ${Number.isFinite(length) ? length : '?'} (a test before this one left a React act scope open, or caught a rejected act)`;
}

export function openSummaryLine(log: ActScopeLog): string | undefined {
  if (log.open === 0 || log.first === undefined) return undefined;
  return `ACT-SCOPE-OPEN summary: ${log.open} tests started with the queue open; first at test #${log.first} in this file`;
}

export function logTestStart(log: ActScopeLog, cache: unknown, write: (line: string) => void): void {
  log.started++;
  const length = openActQueueLength(cache);
  if (length === undefined) return;
  log.open++;
  log.first ??= log.started;
  write(openAtStartLine(log.started, length));
}

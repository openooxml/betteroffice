import type { ProjectionAllowance } from './projectionParity.allowlist';

export const tables: readonly ProjectionAllowance[] = [
  { kind: 'table', path: 'blocks[]', reason: 'Deferred tables can be absent or retain source cell stories.' },
];

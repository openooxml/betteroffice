import type { ProjectionAllowance } from './projectionParity.allowlist';

export const tables: readonly ProjectionAllowance[] = [
  { kind: 'table', path: 'deferred', reason: 'Table and cell story lowering is deferred.' },
  { kind: 'table', path: 'blocks[]', reason: 'Deferred tables can be absent or retain source cell stories.' },
  { kind: 'table', path: 'blocks[].formatting.styleCascade', reason: 'Retained tables can lack resolved table style provenance.' },
  { kind: 'table', path: 'blocks[].rows[].cells[].formatting', reason: 'A retained table cell can lack generated cell formatting.' },
  { kind: 'table', path: 'blocks[].rows[].cells[].formatting.width', reason: 'A retained table cell can lack its generated width.' },
  { kind: 'table', path: 'blocks[].rows[].cells[].formatting.verticalAlign', reason: 'A retained vertical merge cell can lack copied alignment.' },
  { kind: 'table', path: 'blocks[].rows[].cells[].content[]', reason: 'A retained vertical merge continuation can contain a source empty paragraph.' },
];

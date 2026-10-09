import type { Layout } from '../layout/pagination';

export const LAYOUT_META_VERSION = 1;

export interface LayoutMetaV1 {
  v: 1;
  layoutRevision: number;
  pageCount: number;
  partial: boolean;
  provisional: boolean;
  notesConverged: boolean;
  pageSizes: Float64Array;
  layoutShell: string;
  headersFootersEpoch: number;
  headersFooters?: string;
}

export type RetainedLayoutMeta = Pick<
  LayoutMetaV1,
  'pageCount' | 'partial' | 'provisional' | 'notesConverged' | 'pageSizes' | 'layoutShell'
>;

export function isLayoutMetaV1(value: unknown): value is LayoutMetaV1 {
  if (value === null || typeof value !== 'object') return false;
  const meta = value as Partial<LayoutMetaV1>;
  return (
    meta.v === LAYOUT_META_VERSION &&
    Number.isSafeInteger(meta.layoutRevision) && (meta.layoutRevision ?? -1) >= 0 &&
    Number.isSafeInteger(meta.pageCount) && (meta.pageCount ?? -1) >= 0 &&
    typeof meta.partial === 'boolean' && typeof meta.provisional === 'boolean' &&
    typeof meta.notesConverged === 'boolean' && meta.pageSizes instanceof Float64Array &&
    meta.pageSizes.length === meta.pageCount! * 2 &&
    typeof meta.layoutShell === 'string' &&
    Number.isSafeInteger(meta.headersFootersEpoch) && (meta.headersFootersEpoch ?? -1) >= 0 &&
    (meta.headersFooters === undefined || typeof meta.headersFooters === 'string')
  );
}

export function readRetainedLayoutMeta(meta: {
  readonly page_count: number;
  readonly partial: boolean;
  readonly provisional: boolean;
  readonly notes_converged: boolean;
  page_sizes(): Float64Array;
  layout_shell_json(): string;
  free(): void;
}): RetainedLayoutMeta {
  try {
    return {
      pageCount: meta.page_count,
      partial: meta.partial,
      provisional: meta.provisional,
      notesConverged: meta.notes_converged,
      pageSizes: meta.page_sizes(),
      layoutShell: meta.layout_shell_json(),
    };
  } finally {
    meta.free();
  }
}

export function layoutMetaSummary(meta: LayoutMetaV1): Layout & { summaryOnly: true } {
  const layout = JSON.parse(meta.layoutShell) as Layout;
  const refuse = (): never => {
    throw new Error('the fragments of a worker layout summary live in the worker');
  };
  for (const page of layout.pages) {
    Object.defineProperty(page, 'fragments', { get: refuse, enumerable: true, configurable: true });
  }
  return Object.assign(layout, { summaryOnly: true as const });
}

import type { DocxDocumentReadValues } from '@betteroffice/docx/yrs';

/** Estimated accessibility-tree nodes the unbuilt sections may add (sections, paragraphs, text, text lines). */
export const UNBUILT_MIRROR_AX_BUDGET = 30_000;
const READ_PAGES = 16;
const EDGE_NODES = 5;

export type PageTexts = NonNullable<DocxDocumentReadValues['pageTexts']>;
export type MirrorPageText = PageTexts['pages'][number];
export type MirrorStamp = { frameEpoch: number; layoutRevision: number };
export type UnbuiltMirrorPage = { id: string; pageIndex: number; source: unknown; revision: number; unbuilt: boolean };
export type MirrorRange = { from: number; to: number };
export type UnbuiltMirrorSnapshot<T> = {
  sections: { pageIndex: number; value: T }[];
  ranges: MirrorRange[];
  nodeCount: number;
};
type RecordEntry<T> = { page: UnbuiltMirrorPage; stale: boolean; generation: number; nodes?: number; value?: T };

export function scheduleUnbuiltMirrorIdle(run: (deadline: Pick<IdleDeadline, 'timeRemaining'>) => void, now = () => performance.now()): () => void {
  if (typeof requestIdleCallback === 'function') {
    const id = requestIdleCallback(run);
    return () => cancelIdleCallback(id);
  }
  const id = setTimeout(() => {
    const start = now();
    run({ timeRemaining: () => Math.max(0, 8 - (now() - start)) });
  }, 32);
  return () => clearTimeout(id);
}

export function createUnbuiltPageMirror<T>(options: {
  idle(run: (deadline: Pick<IdleDeadline, 'timeRemaining'>) => void): () => void;
  query(pages: number[], stamp: MirrorStamp): Promise<PageTexts | null>;
  blocked(): boolean;
  build(page: MirrorPageText): { value: T; nodes: number };
  focusedPage(): number | null;
  publish(snapshot: UnbuiltMirrorSnapshot<T>): void;
  budget?: number;
  onError?(error: unknown): void;
}) {
  const budget = options.budget ?? UNBUILT_MIRROR_AX_BUDGET;
  const records = new Map<string, RecordEntry<T>>();
  let pages: readonly UnbuiltMirrorPage[] = [];
  let stamp: MirrorStamp = { frameEpoch: 0, layoutRevision: 0 };
  let identity = '';
  let center = 0;
  let enabled = false;
  let disposed = false;
  let inFlight = false;
  let cancel: (() => void) | null = null;
  let generation = 0;
  let queued: { text: MirrorPageText; record: RecordEntry<T>; generation: number; stamp: MirrorStamp }[] = [];
  let selected = new Set<number>();
  let last: UnbuiltMirrorSnapshot<T> = { sections: [], ranges: [], nodeCount: 0 };
  const publish = (next: UnbuiltMirrorSnapshot<T>) => {
    if (next.nodeCount === last.nodeCount && next.sections.length === last.sections.length && next.ranges.length === last.ranges.length &&
      next.sections.every((section, at) => section.pageIndex === last.sections[at]!.pageIndex && section.value === last.sections[at]!.value) &&
      next.ranges.every((range, at) => range.from === last.ranges[at]!.from && range.to === last.ranges[at]!.to)) return;
    last = next;
    options.publish(next);
  };

  const current = (read: MirrorStamp) => read.frameEpoch === stamp.frameEpoch && read.layoutRevision === stamp.layoutRevision;
  const rangesFor = (indices: Set<number>): MirrorRange[] => {
    const ranges: MirrorRange[] = [];
    for (const page of pages) {
      if (!page.unbuilt || indices.has(page.pageIndex)) continue;
      const last = ranges.at(-1);
      if (last && last.to === page.pageIndex) last.to += 1;
      else ranges.push({ from: page.pageIndex + 1, to: page.pageIndex + 1 });
    }
    return ranges;
  };
  const windowFor = (limit: number) => {
    const indices = new Set<number>();
    if (!pages.length) return indices;
    const focused = options.focusedPage();
    const pinned = [...records.values()].find((record) => record.page.pageIndex === focused && record.value !== undefined);
    let nodes = pinned?.nodes ?? 0;
    if (pinned) indices.add(pinned.page.pageIndex);
    let unknown = 0;
    let left = Math.max(0, Math.min(pages.length - 1, center));
    let right = left + 1;
    let before = true;
    let after = true;
    while ((before && left >= 0) || (after && right < pages.length)) {
      const takeLeft = before && left >= 0 && (!after || right >= pages.length || center - left <= right - center);
      const at = takeLeft ? left-- : right++;
      const page = pages[at]!;
      const record = records.get(page.id);
      const cost = !page.unbuilt || indices.has(page.pageIndex) ? 0 : record?.nodes ?? 1;
      if (nodes + cost > limit || (page.unbuilt && record?.nodes === undefined && unknown >= READ_PAGES)) {
        if (takeLeft) before = false; else after = false;
        continue;
      }
      if (page.unbuilt) {
        if (record?.nodes === undefined) unknown += 1;
        indices.add(page.pageIndex);
      }
      nodes += cost;
    }
    return indices;
  };
  const refresh = () => {
    selected = enabled ? windowFor(budget) : new Set();
    let ranges = enabled ? rangesFor(selected) : [];
    let reserved = 0;
    while (ranges.length * EDGE_NODES > reserved) {
      reserved = ranges.length * EDGE_NODES;
      selected = windowFor(Math.max(0, budget - reserved));
      ranges = rangesFor(selected);
    }
    const sections: UnbuiltMirrorSnapshot<T>['sections'] = [];
    let nodeCount = ranges.length * EDGE_NODES;
    for (const page of pages) {
      const record = records.get(page.id);
      if (!record) continue;
      if (!selected.has(page.pageIndex)) record.value = undefined;
      if (record.value !== undefined) {
        sections.push({ pageIndex: page.pageIndex, value: record.value });
        nodeCount += record.nodes ?? 0;
      }
    }
    publish({ sections, ranges, nodeCount });
  };
  const pending = () => [...records.values()].filter((record) => selected.has(record.page.pageIndex) &&
    (record.stale || record.value === undefined)).sort((a, b) => a.page.pageIndex - b.page.pageIndex);
  const schedule = () => {
    if (disposed || !enabled || cancel || inFlight || (!queued.length && !pending().length)) return;
    cancel = options.idle(slice);
  };
  const slice = (deadline: Pick<IdleDeadline, 'timeRemaining'>) => {
    cancel = null;
    if (disposed || !enabled) return;
    if (options.blocked()) { schedule(); return; }
    let built = false;
    while (queued.length && deadline.timeRemaining() > 3) {
      const item = queued.shift()!;
      if (!current(item.stamp) || item.record.generation !== item.generation ||
        records.get(item.record.page.id) !== item.record || !selected.has(item.text.pageIndex)) continue;
      const made = options.build(item.text);
      item.record.value = made.value;
      item.record.nodes = made.nodes;
      item.record.stale = false;
      built = true;
    }
    if (built) refresh();
    if (!queued.length && deadline.timeRemaining() > 3 && !options.blocked()) {
      const batch: RecordEntry<T>[] = [];
      for (const record of pending()) {
        if (batch.length === READ_PAGES || deadline.timeRemaining() <= 3) break;
        batch.push(record);
      }
      if (batch.length) {
        const readStamp = { ...stamp };
        const readIdentity = identity;
        const generations = batch.map((record) => record.generation);
        inFlight = true;
        void options.query(batch.map((record) => record.page.pageIndex), readStamp).then((answer) => {
          if (disposed || !enabled || identity !== readIdentity || !answer || !current(answer) || !current(readStamp)) return;
          queued = answer.pages.flatMap((text) => {
            const at = batch.findIndex((record) => record.page.pageIndex === text.pageIndex);
            return at < 0 ? [] : [{ text, record: batch[at]!, generation: generations[at]!, stamp: readStamp }];
          }).sort((a, b) => a.text.pageIndex - b.text.pageIndex);
        }).catch((error) => { if (!disposed) options.onError?.(error); }).finally(() => {
          inFlight = false;
          schedule();
        });
      }
    }
    schedule();
  };
  return {
    resume() { disposed = false; },
    update(next: MirrorStamp & { identity: string; pages: readonly UnbuiltMirrorPage[]; center: number; enabled: boolean }) {
      if (disposed) return;
      if (identity !== next.identity || !next.enabled) { records.clear(); queued = []; }
      identity = next.identity;
      stamp = { frameEpoch: next.frameEpoch, layoutRevision: next.layoutRevision };
      pages = next.pages;
      center = next.center;
      enabled = next.enabled;
      if (!enabled) {
        cancel?.(); cancel = null;
        selected.clear();
        publish({ sections: [], ranges: [], nodeCount: 0 });
        return;
      }
      const retained = new Set(pages.filter((page) => page.unbuilt).map((page) => page.id));
      for (const id of records.keys()) if (!retained.has(id)) records.delete(id);
      for (const page of pages) {
        if (!page.unbuilt) continue;
        const record = records.get(page.id);
        if (!record) records.set(page.id, { page, stale: true, generation: ++generation });
        else if (record.page.pageIndex !== page.pageIndex || record.page.source !== page.source || record.page.revision !== page.revision) {
          record.page = page;
          record.stale = true;
          record.generation = ++generation;
        }
      }
      refresh();
      schedule();
    },
    dispose() { disposed = true; cancel?.(); records.clear(); queued = []; },
  };
}

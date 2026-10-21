import type {
  CommandDraft, CommandLane, CommandTicket, Outbox, OutboxEntry, OutboxOptions,
} from './types';

type Node<P, X> = {
  entry: OutboxEntry<P, X>;
  tickets: CommandTicket[];
  segment: Segment<P, X>;
  order: number;
  afterId?: number;
  previous?: Node<P, X>;
  next?: Node<P, X>;
};

type Segment<P, X> = {
  size: number;
  lanes: Record<CommandLane, ReadyLane<Node<P, X>>>;
  next?: Segment<P, X>;
};

type ReadyLane<T> = {
  set(order: number, node: T): void;
  delete(order: number): void;
  values(): IterableIterator<T>;
};

type Controls = {
  fence(): void;
  cancel(tickets: readonly CommandTicket[]): readonly OutboxEntry<unknown, unknown>[];
};

const controls = new WeakMap<object, Controls>();

/** @internal Closes the current submission segment without adding a command. */
export function fenceOutbox<P, X>(outbox: Outbox<P, X>): void {
  controls.get(outbox)?.fence();
}

/** @internal Removes a bounded descendant batch; empty tickets continue draining cancellations. */
export function cancelOutboxAfter<P, X>(
  outbox: Outbox<P, X>, tickets: readonly CommandTicket[]
): readonly OutboxEntry<P, X>[] {
  return (controls.get(outbox)?.cancel(tickets) ?? []) as readonly OutboxEntry<P, X>[];
}

function readyLane<T>(): ReadyLane<T> {
  const nodes = new Map<number, T>();
  const levels = Array.from({ length: 11 }, () => new Map<number, number>());

  function set(order: number, node: T): void {
    nodes.set(order, node);
    let index = order;
    for (const masks of levels) {
      const key = Math.floor(index / 32);
      const previous = masks.get(key) ?? 0;
      masks.set(key, previous | (1 << (index % 32)));
      if (previous !== 0) break;
      index = key;
    }
  }

  function remove(order: number): void {
    if (!nodes.delete(order)) return;
    let index = order;
    for (const masks of levels) {
      const key = Math.floor(index / 32);
      const mask = (masks.get(key) ?? 0) & ~(1 << (index % 32));
      if (mask !== 0) { masks.set(key, mask); break; }
      masks.delete(key);
      index = key;
    }
  }

  function* walk(level: number, key: number): IterableIterator<T> {
    let mask = levels[level]!.get(key) ?? 0;
    while (mask !== 0) {
      const bit = 31 - Math.clz32(mask & -mask);
      const index = key * 32 + bit;
      if (level === 0) yield nodes.get(index)!;
      else yield* walk(level - 1, index);
      mask &= mask - 1;
    }
  }

  return { set, delete: remove, values: () => walk(levels.length - 1, 0) };
}

function segment<P, X>(): Segment<P, X> {
  return {
    size: 0,
    lanes: { input: readyLane<Node<P, X>>(), collab: readyLane<Node<P, X>>(), background: readyLane<Node<P, X>>() },
  };
}

/** @internal Creates an unsent draft queue with causal lanes, fences, and ticket aliases. */
export function createOutbox<P, X>(options: OutboxOptions<P, X>): Outbox<P, X> {
  const entries = new Map<number, Node<P, X>>();
  let owners = new WeakMap<CommandTicket, { id: number }>();
  const dependents = new Map<number, Set<Node<P, X>>>();
  const cancelledChildren = new Set<IterableIterator<Node<P, X>>>();
  let first = segment<P, X>();
  let lastSegment = first;
  let last: Node<P, X> | undefined;
  let order = 0;
  const burstCap = Number.isSafeInteger(options.inputBurstCap) && options.inputBurstCap! > 0
    ? options.inputBurstCap! : 8;
  let inputBurst = 0;

  function fence(): void {
    if (lastSegment.size === 0) return;
    const next = segment<P, X>();
    lastSegment.next = next;
    lastSegment = next;
  }

  function ready(node: Node<P, X>): void {
    node.segment.lanes[laneOf(node.entry.draft)].set(node.order, node);
  }

  function merge(previous: Node<P, X>, draft: CommandDraft<P, X>): CommandDraft<P, X> | undefined {
    const base = previous.entry.draft;
    const after = draft.options.after;
    const next = after && owners.get(after) === owners.get(previous.tickets[0]!)
      ? { ...draft, options: { ...draft.options, after: base.options.after } } : draft;
    const merged = compatible(base, next) ? options.coalesce?.(base, next) : undefined;
    return merged && laneOf(merged) === laneOf(base) && merged.options.after === base.options.after &&
      !merged.options.fence ? merged : undefined;
  }

  function enqueue(draft: CommandDraft<P, X>, ticket: CommandTicket): OutboxEntry<P, X> {
    const previous = last?.segment === lastSegment ? last : undefined;
    const merged = previous && merge(previous, draft);
    if (merged && previous) {
      previous.tickets.push(ticket);
      previous.entry = { id: previous.entry.id, draft: merged, tickets: previous.tickets };
      owners.set(ticket, owners.get(previous.tickets[0]!)!);
      return previous.entry;
    }
    if (draft.options.fence) fence();
    if (order === Number.MAX_SAFE_INTEGER) throw new RangeError('Outbox order space is exhausted');
    const tickets = [ticket];
    const entry = { id: ticket.id, draft, tickets };
    const node: Node<P, X> = { entry, tickets, segment: lastSegment, order: ++order, previous: last };
    if (last) last.next = node;
    last = node;
    entries.set(entry.id, node);
    owners.set(ticket, { id: entry.id });
    lastSegment.size += 1;
    const after = draft.options.after;
    if (after) {
      node.afterId = owners.get(after)?.id ?? after.id;
      let children = dependents.get(node.afterId);
      if (!children) dependents.set(node.afterId, children = new Set());
      children.add(node);
    }
    if (!after || !entries.has(node.afterId!)) ready(node);
    if (draft.options.fence) fence();
    return entry;
  }

  function unlink(node: Node<P, X>): void {
    entries.delete(node.entry.id);
    node.segment.lanes[laneOf(node.entry.draft)].delete(node.order);
    node.segment.size -= 1;
    if (node.previous) node.previous.next = node.next;
    if (node.next) node.next.previous = node.previous;
    if (last === node) last = node.previous;
    if (node.afterId !== undefined) {
      const children = dependents.get(node.afterId);
      children?.delete(node);
      if (children?.size === 0) dependents.delete(node.afterId);
    }
    const children = dependents.get(node.entry.id);
    if (children) {
      for (const child of children) ready(child);
    }
    while (first.size === 0 && first.next) first = first.next;
    if (entries.size === 0) inputBurst = 0;
  }

  function candidate(
    lane: CommandLane, eligible: (entry: OutboxEntry<P, X>) => boolean, hold?: (entry: OutboxEntry<P, X>) => boolean
  ): Node<P, X> | undefined {
    for (const node of first.lanes[lane].values()) {
      if (hold?.(node.entry)) return undefined;
      const after = node.entry.draft.options.after;
      if (after && (entries.has(node.afterId!) || after.seq === undefined)) continue;
      if (eligible(node.entry)) return node;
    }
    return undefined;
  }

  function takeNext(
    eligible: (entry: OutboxEntry<P, X>) => boolean, hold?: (entry: OutboxEntry<P, X>) => boolean
  ): OutboxEntry<P, X> | undefined {
    while (first.size === 0 && first.next) first = first.next;
    const collab = inputBurst >= burstCap ? candidate('collab', eligible, hold) : undefined;
    const node = collab ?? candidate('input', eligible, hold) ??
      candidate('collab', eligible, hold) ?? candidate('background', eligible, hold);
    if (!node) return undefined;
    const nextBurst = laneOf(node.entry.draft) === 'input' ? Math.min(burstCap, inputBurst + 1) : 0;
    unlink(node);
    inputBurst = nextBurst;
    return node.entry;
  }

  function descendants(tickets: readonly CommandTicket[], removed: Set<Node<P, X>>): void {
    const pending = tickets.map((ticket) => owners.get(ticket)?.id ?? ticket.id);
    for (let index = 0; index < pending.length; index += 1) {
      const children = dependents.get(pending[index]!);
      if (!children) continue;
      for (const child of children) {
        if (removed.has(child)) continue;
        removed.add(child);
        pending.push(child.entry.id);
      }
    }
  }

  function cancel(tickets: readonly CommandTicket[]): readonly OutboxEntry<P, X>[] {
    for (const ticket of tickets) {
      const id = owners.get(ticket)?.id ?? ticket.id;
      const children = dependents.get(id);
      if (children) {
        cancelledChildren.add(children.values());
        dependents.delete(id);
      }
    }
    const result: OutboxEntry<P, X>[] = [];
    while (cancelledChildren.size && result.length < 64) {
      const iterator = cancelledChildren.values().next().value!;
      const next = iterator.next();
      if (next.done) { cancelledChildren.delete(iterator); continue; }
      const node = next.value;
      if (!entries.has(node.entry.id)) continue;
      const children = dependents.get(node.entry.id);
      if (children) {
        cancelledChildren.add(children.values());
        dependents.delete(node.entry.id);
      }
      result.push(node.entry);
      unlink(node);
    }
    return result;
  }

  function remove(predicate: (entry: OutboxEntry<P, X>) => boolean): readonly OutboxEntry<P, X>[] {
    const removed = new Set<Node<P, X>>();
    const tickets: CommandTicket[] = [];
    for (const node of entries.values()) {
      if (!predicate(node.entry)) continue;
      removed.add(node);
      tickets.push(node.tickets[0]!);
    }
    descendants(tickets, removed);
    const result: OutboxEntry<P, X>[] = [];
    for (const node of entries.values()) {
      if (!removed.has(node)) continue;
      result.push(node.entry);
      unlink(node);
    }
    return result;
  }

  const outbox: Outbox<P, X> = {
    get size() { return entries.size; },
    get entries() { return [...entries.values()].map((node) => node.entry); },
    enqueue, takeNext, remove,
    clear() {
      const removed = [...entries.values()].map((node) => node.entry);
      entries.clear();
      owners = new WeakMap();
      dependents.clear();
      cancelledChildren.clear();
      first = segment<P, X>();
      lastSegment = first;
      last = undefined;
      inputBurst = 0;
      return removed;
    },
  };
  controls.set(outbox, { fence, cancel });
  return outbox;
}

function laneOf<P, X>(draft: CommandDraft<P, X>): CommandLane {
  return draft.options.lane ?? 'input';
}

function compatible<P, X>(prev: CommandDraft<P, X>, next: CommandDraft<P, X>): boolean {
  const a = prev.options;
  const b = next.options;
  return !a.fence && !b.fence && a.coalesceKey !== undefined &&
    a.coalesceKey === b.coalesceKey && prev.kind === next.kind && laneOf(prev) === laneOf(next) &&
    a.group === b.group && a.after?.id === b.after?.id &&
    (a.expectVersion ?? false) === (b.expectVersion ?? false);
}

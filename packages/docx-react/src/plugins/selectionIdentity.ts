export interface DocxSelectionIdentity {
  readonly settleKey: string;
  readonly stickyKey: string;
  readonly resolvedKey: string;
}

export function docxSelectionIdentity(localRevision: number, sticky: unknown, resolved: unknown, collapsed: boolean): DocxSelectionIdentity {
  const exists = resolved !== null;
  return Object.freeze({
    settleKey: JSON.stringify([localRevision, exists, collapsed]),
    stickyKey: JSON.stringify(sticky),
    resolvedKey: JSON.stringify([resolved, exists, collapsed]),
  });
}

const activities = new WeakMap<object, { revision: number; listeners: Set<() => void> }>();

function activity(owner: object) {
  let current = activities.get(owner);
  if (!current) {
    current = { revision: 0, listeners: new Set() };
    activities.set(owner, current);
  }
  return current;
}

export function localSelectionRevision(owner: object | null | undefined): number {
  return owner ? activity(owner).revision : 0;
}

export function notifySelectionActivity(owner: object): void {
  for (const listener of activity(owner).listeners) listener();
}

export function bumpLocalSelection(owner: object): void {
  activity(owner).revision += 1;
  notifySelectionActivity(owner);
}

export function subscribeSelectionActivity(owner: object, listener: () => void): () => void {
  const { listeners } = activity(owner);
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}

export interface DocxSelectionIdentity {
  readonly settleKey: string;
  readonly stickyKey: string;
  readonly resolvedKey: string;
}

export function createDocxSelectionObserver() {
  let previous: { owner: object; identity: DocxSelectionIdentity } | undefined;
  let localRevision = 0;
  return (owner: object, identity: DocxSelectionIdentity, cause?: 'remote' | 'seed'): DocxSelectionIdentity => {
    if (previous?.owner !== owner) localRevision = 0;
    else if (previous.identity.resolvedKey !== identity.resolvedKey && cause !== 'remote' && cause !== 'seed') localRevision += 1;
    const unchanged = previous?.owner === owner && previous.identity.resolvedKey === identity.resolvedKey;
    const settleKey = previous?.owner === owner && (unchanged || cause === 'seed')
      ? previous.identity.settleKey : JSON.stringify([localRevision, ...JSON.parse(identity.settleKey).slice(1)]);
    const observed = Object.freeze({ ...identity, settleKey });
    previous = { owner, identity: observed };
    return observed;
  };
}

export function docxSelectionIdentity(localRevision: number, sticky: unknown, resolved: unknown, collapsed: boolean): DocxSelectionIdentity {
  const exists = resolved !== null;
  return Object.freeze({
    settleKey: JSON.stringify([localRevision, exists, collapsed]),
    stickyKey: JSON.stringify(sticky),
    resolvedKey: JSON.stringify([resolved, exists, collapsed]),
  });
}

const activities = new WeakMap<object, { listeners: Set<() => void> }>();

function activity(owner: object) {
  let current = activities.get(owner);
  if (!current) {
    current = { listeners: new Set() };
    activities.set(owner, current);
  }
  return current;
}

export function notifySelectionActivity(owner: object): void {
  for (const listener of activity(owner).listeners) listener();
}

export function subscribeSelectionActivity(owner: object, listener: () => void): () => void {
  const { listeners } = activity(owner);
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}

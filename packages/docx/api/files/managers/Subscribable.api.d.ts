// @betteroffice/docx declaration managers/Subscribable.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Subscribable Base Class
 *
 * Framework-agnostic base for manager classes that need to notify
 * UI frameworks of state changes.
 *
 * Compatible with:
 * - React: useSyncExternalStore(manager.subscribe, manager.getSnapshot)
 * - Vue: watchEffect(() => { manager.subscribe(triggerRef) })
 */
export declare abstract class Subscribable<TSnapshot> {
    private listeners;
    private snapshot;
    constructor(initialSnapshot: TSnapshot);
    /**
     * Subscribe to state changes. Returns an unsubscribe function.
     * Bound method — safe to pass as `useSyncExternalStore(manager.subscribe, ...)`.
     */
    subscribe: (listener: () => void) => (() => void);
    /**
     * Get the current snapshot. Returns a stable reference unless state has changed.
     * Bound method — safe to pass as `useSyncExternalStore(..., manager.getSnapshot)`.
     */
    getSnapshot: () => TSnapshot;
    /**
     * Update the snapshot and notify all subscribers.
     * Subclasses should call this whenever their state changes.
     */
    protected setSnapshot(snapshot: TSnapshot): void;
    private notify;
}

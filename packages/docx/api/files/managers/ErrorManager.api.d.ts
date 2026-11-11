// @betteroffice/docx declaration managers/ErrorManager.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * ErrorManager
 *
 * Framework-agnostic pub/sub error notification system.
 * Replaces React's `componentDidCatch` + context pattern for error notifications.
 *
 * Usage with React:
 * ```ts
 * const { notifications } = useSyncExternalStore(manager.subscribe, manager.getSnapshot);
 * ```
 */
import { Subscribable } from './Subscribable';
import type { ErrorManagerSnapshot } from './types';
export declare class ErrorManager extends Subscribable<ErrorManagerSnapshot> {
    private notifications;
    private idCounter;
    private timers;
    constructor();
    /** Show an error notification (persistent, not auto-dismissed). */
    showError(message: string, details?: string): string;
    /** Show a warning notification (auto-dismissed after 5s). */
    showWarning(message: string, details?: string): string;
    /** Show an info notification (auto-dismissed after 5s). */
    showInfo(message: string, details?: string): string;
    /** Dismiss a notification by ID. */
    dismiss(id: string): void;
    /** Clear all notifications and cancel pending timers. */
    clearAll(): void;
    /** Destroy the manager and clean up all timers. */
    destroy(): void;
    private addNotification;
    private emitSnapshot;
}

/**
 * Manager Classes — Framework-Agnostic Business Logic
 *
 * These classes contain the state machines and coordination logic
 * extracted from React components and hooks. They can be consumed
 * by any UI framework via the subscribe/getSnapshot pattern.
 */

// Base class
export { Subscribable } from './Subscribable';

// Types
export type {
  CellCoordinates,
  TableSelectionSnapshot,
  ErrorSeverity,
  ErrorNotification,
  ErrorManagerSnapshot,
} from './types';

export { TableSelectionManager } from './TableSelectionManager';
export {
  TABLE_DATA_ATTRIBUTES,
  findTableFromClick,
  getTableFromDocument,
  updateTableInDocument,
  deleteTableFromDocument,
} from './TableSelectionManager';

export { ErrorManager } from './ErrorManager';

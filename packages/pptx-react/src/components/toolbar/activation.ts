import { pptxCommandController } from '../../commands/createPptxCommandStore';
import type {
  PptxCommandArgs,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandStore,
} from '../../commands/types';

let modality: 'pointer' | 'keyboard' = 'pointer';
let tracking = false;

function track(): void {
  if (tracking || typeof document === 'undefined') return;
  tracking = true;
  const pointer = () => {
    modality = 'pointer';
  };
  document.addEventListener('pointerdown', pointer, true);
  document.addEventListener('mousedown', pointer, true);
  document.addEventListener(
    'keydown',
    () => {
      modality = 'keyboard';
    },
    true
  );
}

/**
 * Returns focus to the slide after a pointer choice in a control that took
 * focus, such as a combobox. Keyboard choices and opened pickers keep focus.
 */
export function restoreFocusAfterPointer(store: PptxCommandStore, result: PptxCommandResult): void {
  track();
  if (modality !== 'pointer' || (result.ok && result.status === 'opened')) return;
  requestAnimationFrame(() => pptxCommandController(store)?.focusEditor());
}

/** Runs a command from a control, returning focus to the slide after pointer use. */
export function runFromControl<K extends PptxCommandId>(
  store: PptxCommandStore,
  id: K,
  args: PptxCommandArgs[K]
): void {
  track();
  void store.execute(id, args).then((result) => restoreFocusAfterPointer(store, result));
}

track();

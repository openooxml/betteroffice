import { docxCommandController } from '../../commands/createDocxCommandStore';
import type { DocxCommandResult, DocxCommandStore } from '../../commands/types';

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

track();

/**
 * Returns focus to the document after a pointer choice in a control that took
 * focus, such as a select. Keyboard choices and opened dialogs keep focus.
 */
export function restoreFocusAfterPointer(store: DocxCommandStore, result: DocxCommandResult): void {
  track();
  if (modality !== 'pointer' || (result.ok && result.status === 'opened')) return;
  requestAnimationFrame(() => docxCommandController(store)?.focusEditor());
}

/** Returns focus to the document, as after committing a font size. */
export function restoreEditorFocus(store: DocxCommandStore): void {
  requestAnimationFrame(() => docxCommandController(store)?.focusEditor());
}

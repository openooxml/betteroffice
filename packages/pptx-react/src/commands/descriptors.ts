import type {
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandId,
  PptxCommandShortcut,
  PptxPluginCommandId,
} from './types';
import {
  normalizeChord as normalizeHostChord,
  parseChord,
} from '../../../../shared/keyboard/format';

export { formatChord, isMacPlatform } from '../../../../shared/keyboard/format';

/** `chord` in canonical form, or null; contributed chords name `Mod`, never a physical `Ctrl`. */
export function normalizeChord(chord: string): string | null {
  const canonical = normalizeHostChord(chord);
  return canonical && !parseChord(canonical).ctrl ? canonical : null;
}

type DescriptorTable = { readonly [K in PptxCommandId]: PptxCommandDescriptor<K> };

function descriptor<K extends PptxCommandId>(
  id: K,
  labelKey: PptxCommandDescriptor<K>['labelKey'],
  mutatesDocument: boolean,
  shortcuts: readonly PptxCommandShortcut<K>[] = []
): PptxCommandDescriptor<K> {
  return Object.freeze({ id, labelKey, mutatesDocument, shortcuts: Object.freeze(shortcuts) });
}

const plain = <K extends PptxCommandId>(chord: string) =>
  ({ chord, args: null } as PptxCommandShortcut<K>);

/** Every PPTX command, with its label and keyboard bindings. */
export const PPTX_COMMAND_DESCRIPTORS: DescriptorTable = Object.freeze({
  cut: descriptor('cut', 'commands.cut', true, [plain('Mod+X')]),
  copy: descriptor('copy', 'commands.copy', false, [plain('Mod+C')]),
  paste: descriptor('paste', 'commands.paste', true, [plain('Mod+V')]),
  pasteAsPlainText: descriptor('pasteAsPlainText', 'commands.pastePlainText', true, [
    plain('Mod+Shift+V'),
  ]),
  deleteObject: descriptor('deleteObject', 'commands.delete', true),
  deleteSlide: descriptor('deleteSlide', 'toolbar.deleteSlide', true),
  duplicateSlide: descriptor('duplicateSlide', 'commands.duplicateSlide', true),
  tableInsert: descriptor('tableInsert', 'commands.insert', true),
  tableDelete: descriptor('tableDelete', 'commands.delete', true),
  bold: descriptor('bold', 'toolbar.bold', true, [plain('Mod+B')]),
  italic: descriptor('italic', 'toolbar.italic', true, [plain('Mod+I')]),
  underline: descriptor('underline', 'toolbar.underline', true, [plain('Mod+U')]),
  strikethrough: descriptor('strikethrough', 'toolbar.strikethrough', true),
  fontFamily: descriptor('fontFamily', 'toolbar.fontFamily', true),
  fontSize: descriptor('fontSize', 'toolbar.fontSize', true),
  fontSizeStep: descriptor('fontSizeStep', 'commands.fontSizeStep', true),
  textColor: descriptor('textColor', 'toolbar.textColor', true),
  alignment: descriptor('alignment', 'toolbar.groups.alignment', true),
  insertSlide: descriptor('insertSlide', 'toolbar.newSlide', true),
  insertImage: descriptor('insertImage', 'toolbar.insertImage', true),
  tool: descriptor('tool', 'toolbar.groups.tools', false),
  shapeFill: descriptor('shapeFill', 'toolbar.fillColor', true),
  shapeStrokeColor: descriptor('shapeStrokeColor', 'toolbar.borderColor', true),
  shapeStrokeWidth: descriptor('shapeStrokeWidth', 'toolbar.borderWidth', true),
  shapeAdjustment: descriptor('shapeAdjustment', 'toolbar.shapeAdjustment', true),
  zOrder: descriptor('zOrder', 'toolbar.arrange', true),
  save: descriptor('save', 'toolbar.save', false, [plain('Mod+S')]),
  exportPng: descriptor('exportPng', 'toolbar.exportPng', false),
  slideshow: descriptor('slideshow', 'toolbar.present', false),
  undo: descriptor('undo', 'toolbar.undo', true, [plain('Mod+Z')]),
  redo: descriptor('redo', 'toolbar.redo', true, [plain('Mod+Shift+Z'), plain('Mod+Y')]),
  zoom: descriptor('zoom', 'toolbar.groups.zoom', false),
  proposalsPanel: descriptor('proposalsPanel', 'proposals.title', false),
  proposalSelect: descriptor('proposalSelect', 'proposals.canvasTitle', false),
  proposalDiff: descriptor('proposalDiff', 'proposals.showDiff', false),
  proposalAccept: descriptor('proposalAccept', 'proposals.accept', true),
  proposalReject: descriptor('proposalReject', 'proposals.reject', true),
});

export const PPTX_COMMAND_IDS = Object.freeze(
  Object.keys(PPTX_COMMAND_DESCRIPTORS) as PptxCommandId[]
);

export function isPptxCommandId(value: unknown): value is PptxCommandId {
  return (
    typeof value === 'string' &&
    Object.prototype.hasOwnProperty.call(PPTX_COMMAND_DESCRIPTORS, value)
  );
}

const PLUGIN_COMMAND_ID =
  /^plugin:[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?\/[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?$/;

export function isPluginCommandId(value: unknown): value is PptxPluginCommandId {
  return typeof value === 'string' && PLUGIN_COMMAND_ID.test(value);
}

/** Arguments a control uses when none are bound. */
export function defaultArgs<K extends PptxCommandId>(id: K): PptxCommandArgs[K] {
  return (id === 'insertSlide' ? {} : null) as PptxCommandArgs[K];
}

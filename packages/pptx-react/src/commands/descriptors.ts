import type {
  PptxCommandArgs,
  PptxCommandDescriptor,
  PptxCommandId,
  PptxCommandShortcut,
} from './types';

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
  bold: descriptor('bold', 'toolbar.bold', true, [plain('Mod+B')]),
  italic: descriptor('italic', 'toolbar.italic', true, [plain('Mod+I')]),
  underline: descriptor('underline', 'toolbar.underline', true, [plain('Mod+U')]),
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
  redo: descriptor('redo', 'toolbar.redo', true, [plain('Mod+Shift+Z')]),
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

/** Arguments a control uses when none are bound. */
export function defaultArgs<K extends PptxCommandId>(id: K): PptxCommandArgs[K] {
  return (id === 'insertSlide' ? {} : null) as PptxCommandArgs[K];
}

interface ParsedChord {
  mod: boolean;
  shift: boolean;
  alt: boolean;
  key: string;
}

function parseChord(chord: string): ParsedChord {
  const parts = chord.split('+');
  const key = parts.length > 1 && parts[parts.length - 1] === '' ? '+' : parts[parts.length - 1];
  const modifiers = new Set(parts.slice(0, -1));
  return {
    mod: modifiers.has('Mod'),
    shift: modifiers.has('Shift'),
    alt: modifiers.has('Alt'),
    key: key.toLowerCase(),
  };
}

export function isMacPlatform(): boolean {
  return typeof navigator !== 'undefined' && /Mac|iPod|iPhone|iPad/.test(navigator.platform);
}

/** Whether a keydown event presses `chord`. */
export function matchesChord(chord: string, event: KeyboardEvent, mac = isMacPlatform()): boolean {
  const parsed = parseChord(chord);
  const mod = mac ? event.metaKey : event.ctrlKey;
  const other = mac ? event.ctrlKey : event.metaKey;
  return (
    parsed.mod === mod &&
    !other &&
    parsed.shift === event.shiftKey &&
    parsed.alt === event.altKey &&
    event.key.toLowerCase() === parsed.key
  );
}

/** A chord as people read it on the current platform (`Ctrl+B`, `⌘B`). */
export function formatChord(chord: string, mac = isMacPlatform()): string {
  const parsed = parseChord(chord);
  const key = parsed.key.length === 1 ? parsed.key.toUpperCase() : parsed.key;
  if (mac) {
    return `${parsed.alt ? '⌥' : ''}${parsed.shift ? '⇧' : ''}${parsed.mod ? '⌘' : ''}${key}`;
  }
  return [parsed.mod && 'Ctrl', parsed.alt && 'Alt', parsed.shift && 'Shift', key]
    .filter(Boolean)
    .join('+');
}

/** The display chord for a command invoked with `args`, if it has one. */
export function commandShortcut<K extends PptxCommandId>(
  id: K,
  args?: PptxCommandArgs[K]
): string | null {
  const shortcuts = PPTX_COMMAND_DESCRIPTORS[id].shortcuts as readonly PptxCommandShortcut<K>[];
  const key = JSON.stringify(args ?? defaultArgs(id));
  const match = shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key);
  return match ? formatChord(match.chord) : null;
}

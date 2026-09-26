import type {
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandId,
  DocxCommandShortcut,
} from './types';

type DescriptorTable = { readonly [K in DocxCommandId]: DocxCommandDescriptor<K> };

function descriptor<K extends DocxCommandId>(
  id: K,
  labelKey: DocxCommandDescriptor<K>['labelKey'],
  mutatesDocument: boolean,
  shortcuts: readonly DocxCommandShortcut<K>[] = []
): DocxCommandDescriptor<K> {
  return Object.freeze({ id, labelKey, mutatesDocument, shortcuts: Object.freeze(shortcuts) });
}

const plain = <K extends DocxCommandId>(chord: string) =>
  ({ chord, args: null }) as DocxCommandShortcut<K>;

/** Every DOCX command, with its label and keyboard bindings. */
export const DOCX_COMMAND_DESCRIPTORS: DescriptorTable = Object.freeze({
  undo: descriptor('undo', 'formattingBar.undo', true, [plain('Mod+Z')]),
  redo: descriptor('redo', 'formattingBar.redo', true, [plain('Mod+Y'), plain('Mod+Shift+Z')]),
  bold: descriptor('bold', 'formattingBar.bold', true, [plain('Mod+B')]),
  italic: descriptor('italic', 'formattingBar.italic', true, [plain('Mod+I')]),
  underline: descriptor('underline', 'formattingBar.underline', true, [plain('Mod+U')]),
  strikethrough: descriptor('strikethrough', 'formattingBar.strikethrough', true),
  superscript: descriptor('superscript', 'formattingBar.superscript', true, [
    plain('Mod+Shift+='),
  ]),
  subscript: descriptor('subscript', 'formattingBar.subscript', true, [plain('Mod+=')]),
  clearFormatting: descriptor('clearFormatting', 'formattingBar.clearFormatting', true),
  paragraphStyle: descriptor('paragraphStyle', 'commands.paragraphStyle', true),
  fontFamily: descriptor('fontFamily', 'commands.fontFamily', true),
  fontSize: descriptor('fontSize', 'fontSize.label', true),
  textColor: descriptor('textColor', 'formattingBar.fontColor', true),
  highlightColor: descriptor('highlightColor', 'formattingBar.highlightColor', true),
  alignment: descriptor('alignment', 'formattingBar.groups.alignment', true, [
    { chord: 'Mod+L', args: { value: 'left' } },
    { chord: 'Mod+E', args: { value: 'center' } },
    { chord: 'Mod+R', args: { value: 'right' } },
    { chord: 'Mod+J', args: { value: 'both' } },
  ]),
  lineSpacing: descriptor('lineSpacing', 'lineSpacing.label', true),
  bulletList: descriptor('bulletList', 'lists.bulletList', true),
  numberedList: descriptor('numberedList', 'lists.numberedList', true),
  indent: descriptor('indent', 'lists.increaseIndent', true),
  outdent: descriptor('outdent', 'lists.decreaseIndent', true),
  setLtr: descriptor('setLtr', 'toolbar.leftToRight', true),
  setRtl: descriptor('setRtl', 'toolbar.rightToLeft', true),
  insertLink: descriptor('insertLink', 'formattingBar.insertLink', true, [plain('Mod+K')]),
  insertImage: descriptor('insertImage', 'toolbar.image', true),
  insertTable: descriptor('insertTable', 'toolbar.table', true),
  insertPageBreak: descriptor('insertPageBreak', 'toolbar.pageBreak', true),
  insertSectionBreakNextPage: descriptor(
    'insertSectionBreakNextPage',
    'toolbar.sectionBreakNextPage',
    true
  ),
  insertSectionBreakContinuous: descriptor(
    'insertSectionBreakContinuous',
    'toolbar.sectionBreakContinuous',
    true
  ),
  insertTOC: descriptor('insertTOC', 'toolbar.tableOfContents', true),
  imageWrap: descriptor('imageWrap', 'commands.imageWrap', true),
  imageTransform: descriptor('imageTransform', 'imageTransform.tooltip', true),
  imageProperties: descriptor('imageProperties', 'formattingBar.imageProperties', true),
  tableAction: descriptor('tableAction', 'formattingBar.groups.table', true),
  pageSetup: descriptor('pageSetup', 'toolbar.pageSetup', true),
  watermark: descriptor('watermark', 'toolbar.watermark', true),
  editingMode: descriptor('editingMode', 'commands.editingMode', false),
  reviewAccept: descriptor('reviewAccept', 'commands.reviewAccept', true),
  reviewReject: descriptor('reviewReject', 'commands.reviewReject', true),
  reviewPrevious: descriptor('reviewPrevious', 'commands.reviewPrevious', false),
  reviewNext: descriptor('reviewNext', 'commands.reviewNext', false),
  commentsSidebar: descriptor('commentsSidebar', 'editor.toggleCommentsSidebar', false),
  open: descriptor('open', 'toolbar.open', false, [plain('Mod+O')]),
  save: descriptor('save', 'toolbar.save', false, [plain('Mod+S')]),
  print: descriptor('print', 'toolbar.print', false),
  find: descriptor('find', 'dialogs.findReplace.titleFind', false, [plain('Mod+F')]),
  replace: descriptor('replace', 'dialogs.findReplace.titleFindReplace', true, [plain('Mod+H')]),
  reportIssue: descriptor('reportIssue', 'toolbar.reportIssue', false),
  zoom: descriptor('zoom', 'zoom.zoomLevel', false),
});

export const DOCX_COMMAND_IDS = Object.freeze(
  Object.keys(DOCX_COMMAND_DESCRIPTORS) as DocxCommandId[]
);

export function isDocxCommandId(value: unknown): value is DocxCommandId {
  return typeof value === 'string' && Object.hasOwn(DOCX_COMMAND_DESCRIPTORS, value);
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

function keyMatches(key: string, event: KeyboardEvent): boolean {
  if (key === '=') return event.code === 'Equal' || event.key === '=' || event.key === '+';
  return event.key.toLowerCase() === key;
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
    keyMatches(parsed.key, event)
  );
}

/** A chord as people read it on the current platform (`Ctrl+B`, `⌘B`). */
export function formatChord(chord: string, mac = isMacPlatform()): string {
  const parsed = parseChord(chord);
  const key = parsed.key.length === 1 ? parsed.key.toUpperCase() : parsed.key;
  if (mac) {
    return `${parsed.mod ? '⌘' : ''}${parsed.alt ? '⌥' : ''}${parsed.shift ? '⇧' : ''}${key}`;
  }
  return [parsed.mod && 'Ctrl', parsed.alt && 'Alt', parsed.shift && 'Shift', key]
    .filter(Boolean)
    .join('+');
}

/** The display chord for a command invoked with `args`, if it has one. */
export function commandShortcut<K extends DocxCommandId>(
  id: K,
  args?: DocxCommandArgs[K]
): string | null {
  const shortcuts = DOCX_COMMAND_DESCRIPTORS[id].shortcuts as readonly DocxCommandShortcut<K>[];
  const key = JSON.stringify(args ?? null);
  const match = shortcuts.find((shortcut) => JSON.stringify(shortcut.args) === key);
  return match ? formatChord(match.chord) : null;
}

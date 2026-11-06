import type {
  DocxCommandArgs,
  DocxCommandDescriptor,
  DocxCommandId,
  DocxCommandShortcut,
  DocxPluginCommandId,
} from './types';
import {
  formatChord,
  isMacPlatform,
  normalizeChord,
} from '../../../../shared/keyboard/chords';

export { formatChord, isMacPlatform, matchesChord, normalizeChord } from '../../../../shared/keyboard/chords';

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
  cut: descriptor('cut', 'contextMenu.cut', true, [plain('Mod+X')]),
  copy: descriptor('copy', 'contextMenu.copy', false, [plain('Mod+C')]),
  paste: descriptor('paste', 'contextMenu.paste', true, [plain('Mod+V')]),
  pasteAsPlainText: descriptor('pasteAsPlainText', 'contextMenu.pastePlainText', true, [
    plain('Mod+Shift+V'),
  ]),
  selectAll: descriptor('selectAll', 'contextMenu.selectAll', false, [plain('Mod+A')]),
  undo: descriptor('undo', 'formattingBar.undo', true, [plain('Mod+Z')]),
  redo: descriptor('redo', 'formattingBar.redo', true, [plain('Mod+Shift+Z'), plain('Mod+Y')]),
  bold: descriptor('bold', 'formattingBar.bold', true, [plain('Mod+B')]),
  italic: descriptor('italic', 'formattingBar.italic', true, [plain('Mod+I')]),
  underline: descriptor('underline', 'formattingBar.underline', true, [plain('Mod+U')]),
  strikethrough: descriptor('strikethrough', 'formattingBar.strikethrough', true),
  superscript: descriptor('superscript', 'formattingBar.superscript', true, [
    plain('Mod+Shift+='),
  ]),
  subscript: descriptor('subscript', 'formattingBar.subscript', true, [plain('Mod+=')]),
  clearFormatting: descriptor('clearFormatting', 'formattingBar.clearFormatting', true, [
    { chord: 'Mod+Space', args: null, platform: 'other' },
    { chord: 'Mod+\\', args: null, platform: 'mac' },
  ]),
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
  paragraphSpacing: descriptor('paragraphSpacing', 'lineSpacing.paragraphSpacing', true),
  paragraphSpacingDialog: descriptor('paragraphSpacingDialog', 'lineSpacing.options', true),
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
  tableGridlines: descriptor('tableGridlines', 'table.viewGridlines', false),
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

const PLUGIN_COMMAND_ID =
  /^plugin:[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?\/[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?$/;

export function isPluginCommandId(value: unknown): value is DocxPluginCommandId {
  return typeof value === 'string' && PLUGIN_COMMAND_ID.test(value);
}

/** Whether a binding exists on the current platform. */
export function shortcutMatchesPlatform(
  shortcut: Pick<DocxCommandShortcut, 'chord' | 'platform'>,
  mac = isMacPlatform()
): boolean {
  return shortcut.platform === undefined || shortcut.platform === (mac ? 'mac' : 'other');
}

/** Canonical chords the text input handles itself (clipboard and select all). */
export const EDITING_CHORDS: ReadonlySet<string> = new Set(
  ['Mod+A', 'Mod+C', 'Mod+V', 'Mod+X', 'Mod+Shift+V'].map((chord) => normalizeChord(chord)!)
);

/** Canonical chords of the built-in commands, which always win over contributed ones. */
export const BUILT_IN_CHORDS: ReadonlySet<string> = new Set(
  Object.values(DOCX_COMMAND_DESCRIPTORS).flatMap((descriptor) =>
    (descriptor.shortcuts as readonly DocxCommandShortcut[]).map(
      (shortcut) => normalizeChord(shortcut.chord)!
    )
  )
);

/** The display chord for a command invoked with `args`, if it has one. */
export function commandShortcut<K extends DocxCommandId>(
  id: K,
  args?: DocxCommandArgs[K],
  mac = isMacPlatform()
): string | null {
  const shortcuts = DOCX_COMMAND_DESCRIPTORS[id].shortcuts as readonly DocxCommandShortcut<K>[];
  const key = JSON.stringify(args ?? null);
  const match = shortcuts.find(
    (shortcut) => shortcutMatchesPlatform(shortcut, mac) && JSON.stringify(shortcut.args) === key
  );
  return match ? formatChord(match.chord, mac) : null;
}

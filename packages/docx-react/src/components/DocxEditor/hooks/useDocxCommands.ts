import { useEffect, useMemo, useRef, useState } from 'react';
import { createT, deepMerge, en } from '@betteroffice/docx-i18n';
import type { LocaleStrings, Translations } from '@betteroffice/docx-i18n';
import type { ImageLayoutTarget } from '@betteroffice/docx/docx';
import { captureInlinePositionEmuFromDisplayList } from '@betteroffice/docx/layout/render';
import type { DisplayList, DisplayListQueries } from '@betteroffice/docx/layout/render';
import { createYrsSidebarProjection } from '@betteroffice/docx/layout/render';
import type { createStyleResolver } from '@betteroffice/docx/styles';
import type { ColorValue, Document, Theme } from '@betteroffice/docx/types/document';
import { resolveParagraphStyleOptions } from '@betteroffice/docx/utils/stylePreview';
import { sameYrsSelection, type YrsSession } from '@betteroffice/docx/yrs';
import {
  createDocxCommandController,
  DocxCommandAdmissionError,
  type DocxCommandBinding,
  type DocxCommandController,
  type DocxCommandOrigin,
  type DocxDeferredCommand,
  type DocxDeferredTarget,
} from '../../../commands/createDocxCommandStore';
import type {
  DocxCommandEnvironment,
  DocxCommandControl,
  DocxRevisionEnvironment,
  DocxSelectionEnvironment,
} from '../../../commands/evaluate';
import type {
  DocxCommandArgs,
  DocxCommandId,
  DocxCommandOption,
  DocxCommandResult,
  DocxTableAction,
} from '../../../commands/types';
import type { useFindReplace } from '../../dialogs/FindReplaceDialog';
import type { useHyperlinkDialog } from '../../dialogs/HyperlinkDialog';
import { openReportIssue } from '../../reportIssue';
import { DEFAULT_FONTS, type FontOption } from '../../ui/FontPicker';
import { getPrimaryFontFamily } from '../../ui/fontPickerValue';
import { normalizeFontFamilies } from '../../ui/normalizeFontFamilies';
import { DEFAULT_STYLES } from '../../ui/StylePicker';
import type { EditorMode } from '../internals/editing-modes';
import type { PagedEditorRef } from '../PagedEditor';
import {
  yrsHyperlinkAtSelection,
  yrsSelectedText,
  yrsTableSelectionStories,
} from '../yrsCommands';
import type { YrsToolbarSelection } from '../yrsToolbar';
import type { DocxPrintJob } from './useFileIO';
import type {
  PagedEditorCommandBridge,
  PagedEditorSelectedImage,
} from './usePagedEditorRefApi';

/** Outcome of the built-in save workflow. */
export type DocxSaveOutcome = 'saved' | 'requested' | 'failed';

/** Table routing result: changed, unchanged, or a dialog opened. */
export type DocxTableActionOutcome = boolean | 'opened';

type StyleResolver = ReturnType<typeof createStyleResolver>;

/** Everything the command binding reads from the editor, refreshed every render. */
export interface DocxCommandInputs {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  bridgeRef: React.RefObject<PagedEditorCommandBridge | null>;
  isLoading: boolean;
  parseError: string | null;
  document: Document | null;
  session: YrsSession | null;
  readOnly: boolean;
  mode: EditorMode;
  modeControlled: boolean;
  onModeChange: ((mode: EditorMode) => void) | undefined;
  setEditingMode: (mode: EditorMode) => void;
  sidebarOpen: boolean;
  sidebarControlled: boolean;
  sidebarHasSetter: boolean;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  setExpandedSidebarItem: React.Dispatch<React.SetStateAction<string | null>>;
  zoom: number;
  setZoom: (zoom: number) => void;
  showFileOpen: boolean;
  showHelpMenu: boolean;
  partEditing: boolean;
  fontFamilies: ReadonlyArray<string | FontOption> | undefined;
  documentFonts: readonly FontOption[];
  theme: Theme | null;
  i18n: Translations | undefined;
  isDark: boolean;
  displayListQueries: DisplayListQueries | null;
  getCachedStyleResolver: (styles: Parameters<typeof createStyleResolver>[0]) => StyleResolver;
  hyperlinkDialog: ReturnType<typeof useHyperlinkDialog>;
  findReplace: ReturnType<typeof useFindReplace>;
  save: () => Promise<DocxSaveOutcome>;
  /** Opens the print window while the user gesture is still active. */
  reservePrint: () => DocxPrintJob;
  /** The display list once it shows the document as it is now. */
  renderedDisplayList: () => Promise<DisplayList>;
  openDocument: () => void;
  /** Opens the image picker; `insert` places the chosen picture. */
  pickImage: (insert: DocxImageInsert) => void;
  tableAction: (action: DocxTableAction) => DocxTableActionOutcome;
  openImageProperties: (image: PagedEditorSelectedImage) => void;
  openPageSetup: () => void;
  openWatermark: () => void;
  refreshTrackedChanges: (session: YrsSession) => void;
}

const IMMEDIATE_COMMANDS: ReadonlySet<DocxCommandId> = new Set<DocxCommandId>([
  'editingMode',
  'commentsSidebar',
  'open',
  'save',
  'print',
  'find',
  'replace',
  'reportIssue',
  'zoom',
  'insertImage',
  'imageProperties',
  'pageSetup',
  'watermark',
]);

const CELL_STORY = /:t\d+:r\d+c\d+/;

/** Inserts a decoded picture at the selection its picker was opened for. */
export type DocxImageInsert = (
  image: Readonly<Record<string, unknown>>
) => Promise<DocxCommandResult>;

/** Dialogs a command opens and a later submission completes. */
export type DocxCommandDialog =
  | 'link'
  | 'linkPopup'
  | 'imageProperties'
  | 'splitCell'
  | 'tableProperties'
  | 'pageSetup'
  | 'watermark'
  | 'replace';

const DIALOG_COMMANDS: {
  readonly [D in DocxCommandDialog]: readonly [DocxCommandId, unknown, DocxDeferredTarget];
} = {
  link: ['insertLink', null, 'selection'],
  linkPopup: ['insertLink', null, 'selection'],
  imageProperties: ['imageProperties', null, 'document'],
  splitCell: ['tableAction', 'splitCell', 'selection'],
  tableProperties: ['tableAction', { type: 'openTableProperties' }, 'selection'],
  pageSetup: ['pageSetup', null, 'document'],
  watermark: ['watermark', null, 'document'],
  replace: ['replace', null, 'document'],
};

interface EditorOrigin extends DocxCommandOrigin {
  readonly document: YrsSession;
  readonly selection: ReturnType<YrsSession['encodeSelection']> | null;
  readonly cells: readonly [string, string] | null;
}

function sameCells(a: readonly [string, string] | null, b: readonly [string, string] | null) {
  return a === b || (a !== null && b !== null && a[0] === b[0] && a[1] === b[1]);
}

function control(controlled: boolean, hasSetter: boolean): DocxCommandControl {
  return controlled ? (hasSetter ? 'host' : 'fixed') : 'internal';
}

/** `executed` when something changed, else `noop`. */
export function commandOutcome(changed: boolean): DocxCommandResult {
  return { ok: true, status: changed ? 'executed' : 'noop' };
}

const executed = commandOutcome;

const OPENED: DocxCommandResult = { ok: true, status: 'opened' };

function finiteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

/** Word's wrap target for projected image attributes. */
export function imageWrapTarget(attrs: Readonly<Record<string, unknown>>): ImageLayoutTarget | null {
  const wrapType = typeof attrs.wrapType === 'string' ? attrs.wrapType : 'inline';
  if (attrs.displayMode === 'float' && attrs.cssFloat === 'left') return 'squareLeft';
  if (attrs.displayMode === 'float' && attrs.cssFloat === 'right') return 'squareRight';
  switch (wrapType) {
    case 'inline':
    case 'square':
    case 'tight':
    case 'through':
    case 'topAndBottom':
    case 'behind':
    case 'inFront':
      return wrapType;
    default:
      return null;
  }
}

function styleOptions(
  document: Document | null,
  t: (key: Parameters<ReturnType<typeof createT>>[0]) => string
): DocxCommandOption<'paragraphStyle'>[] {
  const resolved = resolveParagraphStyleOptions(document?.package.styles?.styles);
  const source = resolved.length > 0 ? resolved : DEFAULT_STYLES;
  return source.map((style) => {
    const nameKey = DEFAULT_STYLES.find((option) => option.styleId === style.styleId)?.nameKey;
    const preview: NonNullable<DocxCommandOption['preview']> = {};
    if (style.fontSize != null) preview.fontSize = style.fontSize / 2;
    if (style.bold != null) preview.bold = style.bold;
    if (style.italic != null) preview.italic = style.italic;
    if (style.color != null) preview.color = style.color;
    return {
      args: { styleId: style.styleId },
      label: nameKey ? t(nameKey) : style.name,
      ...(Object.keys(preview).length > 0 ? { preview } : {}),
    };
  });
}

function fontOptions(
  fontFamilies: ReadonlyArray<string | FontOption> | undefined,
  documentFonts: readonly FontOption[]
): DocxCommandOption<'fontFamily'>[] {
  const fonts = normalizeFontFamilies(fontFamilies) ?? DEFAULT_FONTS;
  const listed = new Set(fonts.map((font) => font.name.toLowerCase()));
  const option = (font: FontOption, group: string): DocxCommandOption<'fontFamily'> => ({
    args: { family: getPrimaryFontFamily(font.fontFamily) || font.name },
    label: font.name,
    preview: { fontFamily: font.fontFamily, group },
  });
  return [
    ...documentFonts
      .filter((font) => !listed.has(font.name.toLowerCase()))
      .map((font) => option(font, 'document')),
    ...fonts.map((font) => option(font, font.category ?? 'other')),
  ];
}

interface RevisionIndex {
  session: YrsSession;
  version: number;
  ids: ReadonlySet<string>;
  navigation: DocxRevisionEnvironment[];
  spans: { revisionId: string; story: string; start: number; end: number }[];
}

function revisionIndex(session: YrsSession, version: number): RevisionIndex {
  const revisions = session.listRevisions();
  const projection = createYrsSidebarProjection(session);
  const ids = new Set(revisions.map((revision) => revision.revisionId));
  const navigation: DocxRevisionEnvironment[] = [];
  const spans: RevisionIndex['spans'] = [];
  const seen = new Set<string>();
  for (const revision of revisions) {
    const story = revision.range.story;
    try {
      spans.push({
        revisionId: revision.revisionId,
        story,
        start:
          session.locateParagraph(story, revision.range.start.paraId).start +
          revision.range.start.offset,
        end:
          session.locateParagraph(story, revision.range.end.paraId).start +
          revision.range.end.offset,
      });
    } catch {
      continue;
    }
    if (seen.has(revision.revisionId)) continue;
    const point = projection.locToDisplayPoint({ story, ...revision.range.start });
    if (!point || point.hfRid) continue;
    seen.add(revision.revisionId);
    navigation.push({ revisionId: revision.revisionId, position: point.position });
  }
  navigation.sort((a, b) => a.position - b.position);
  return { session, version, ids, navigation, spans };
}

function selectionEnvironment(
  toolbar: YrsToolbarSelection,
  inputs: DocxCommandInputs
): DocxSelectionEnvironment {
  const { context } = toolbar;
  let fontFamily = context.fontFamily;
  let halfPoints = context.fontSize;
  const defaults = context.paragraphProperties.defaultTextFormatting;
  if (halfPoints == null && defaults && typeof defaults === 'object') {
    const size = (defaults as { fontSize?: unknown }).fontSize;
    halfPoints =
      finiteNumber(size) ??
      (size && typeof size === 'object' ? finiteNumber((size as { size?: unknown }).size) : null);
  }
  const styles = inputs.document?.package.styles;
  if ((!fontFamily || halfPoints == null) && styles && context.styleId) {
    const resolved = inputs.getCachedStyleResolver(styles).resolveParagraphStyle(context.styleId);
    fontFamily ||=
      resolved.runFormatting?.fontFamily?.ascii ?? resolved.runFormatting?.fontFamily?.hAnsi ?? null;
    halfPoints ??= resolved.runFormatting?.fontSize ?? null;
  }
  return {
    context,
    fontFamily: fontFamily || null,
    fontSize: halfPoints == null ? null : halfPoints / 2,
  };
}

export interface DocxCommandsHandle {
  controller: DocxCommandController;
  /** Opens `dialog` outside a command, capturing its document and target now. */
  open(dialog: DocxCommandDialog): void;
  /** Finishes `dialog` against the document and target it opened for. */
  complete(
    dialog: DocxCommandDialog,
    write: () => DocxCommandResult | Promise<DocxCommandResult>
  ): Promise<DocxCommandResult>;
}

/** Creates the editor's command store and binds it to live editor state. */
export function useDocxCommandBinding(inputs: DocxCommandInputs): DocxCommandsHandle {
  const [controller] = useState(createDocxCommandController);
  const latest = useRef(inputs);
  latest.current = inputs;
  const documentVersion = useRef(0);
  const revisions = useRef<RevisionIndex | null>(null);
  const continuations = useRef(new Map<DocxCommandDialog, DocxDeferredCommand>());
  const openDialog = useRef((dialog: DocxCommandDialog): void => {
    const [id, args, target] = DIALOG_COMMANDS[dialog];
    continuations.current.set(dialog, controller.defer(id, args as never, target));
  }).current;

  const translation = useMemo(() => {
    const merged = deepMerge(
      en as Record<string, unknown>,
      inputs.i18n as Record<string, unknown> | undefined
    ) as LocaleStrings;
    return createT(merged, typeof inputs.i18n?._lang === 'string' ? inputs.i18n._lang : 'en');
  }, [inputs.i18n]);
  const translationRef = useRef(translation);
  translationRef.current = translation;

  const styleCache = useRef<{
    key: unknown;
    t: unknown;
    options: DocxCommandOption<'paragraphStyle'>[];
  } | null>(null);
  const fontCache = useRef<{
    families: unknown;
    documentFonts: unknown;
    options: DocxCommandOption<'fontFamily'>[];
  } | null>(null);

  const binding = useMemo<DocxCommandBinding>(() => {
    const bridge = () => latest.current.bridgeRef.current;

    const environment = (executing: boolean): DocxCommandEnvironment => {
      const current = latest.current;
      const editor = bridge();
      const session = current.session;
      const t = translationRef.current;
      const status: DocxCommandEnvironment['status'] = current.isLoading
        ? 'loading'
        : current.parseError || !current.document
          ? 'empty'
          : !session || !editor || editor.session() !== session
            ? 'loading'
            : 'ready';
      let toolbar: YrsToolbarSelection | null | undefined;
      const readToolbar = () => {
        if (toolbar === undefined) toolbar = editor?.toolbarSelection(executing) ?? null;
        return toolbar;
      };
      let selectionMemo: DocxCommandEnvironment['selection'] | undefined;
      let imageMemo: DocxCommandEnvironment['image'] | undefined;
      let currentRevision: string | null | undefined;
      const index = () => {
        if (!session) return null;
        const cached = revisions.current;
        if (cached && cached.session === session && cached.version === documentVersion.current) {
          return cached;
        }
        revisions.current = revisionIndex(session, documentVersion.current);
        return revisions.current;
      };
      return {
        status,
        readOnly: current.readOnly,
        mode: current.mode,
        modeControl: control(current.modeControlled, current.onModeChange !== undefined),
        sidebarOpen: current.sidebarOpen,
        sidebarControl: control(current.sidebarControlled, current.sidebarHasSetter),
        zoom: current.zoom,
        canOpen: current.showFileOpen,
        canReportIssue: current.showHelpMenu,
        bodyStory: !current.partEditing && (editor?.rootStory() ?? 'body') === 'body',
        pendingInput: !executing && (editor?.hasPendingInput() ?? false),
        canUndo: status === 'ready' && (session?.canUndo() ?? false),
        canRedo: status === 'ready' && (session?.canRedo() ?? false),
        get selection() {
          if (selectionMemo !== undefined) return selectionMemo;
          const read = status === 'ready' ? readToolbar() : null;
          selectionMemo = read
            ? selectionEnvironment(read, current)
            : status === 'ready' && editor?.hasSelection()
              ? 'unsupported'
              : null;
          return selectionMemo;
        },
        get table() {
          const read = status === 'ready' ? readToolbar() : null;
          if (read?.tableContext?.isInTable) return read.tableContext;
          const story = session?.selection()?.head.story;
          return status === 'ready' && story && CELL_STORY.test(story)
            ? { isInTable: true, canSplitCell: true }
            : null;
        },
        get image() {
          if (imageMemo !== undefined) return imageMemo;
          const selected = status === 'ready' ? (editor?.selectedImage() ?? null) : null;
          imageMemo = selected ? { wrap: imageWrapTarget(selected.attrs) } : null;
          return imageMemo;
        },
        get revisions() {
          return status === 'ready' ? (index()?.navigation ?? []) : [];
        },
        get revisionIds() {
          return status === 'ready' ? (index()?.ids ?? new Set<string>()) : new Set<string>();
        },
        get currentRevisionId() {
          if (currentRevision !== undefined) return currentRevision;
          currentRevision = null;
          const head = status === 'ready' ? session?.selection()?.head : null;
          const spans = head ? index()?.spans : null;
          if (session && head && spans) {
            try {
              const offset = session.locateParagraph(head.story, head.paraId).start + head.offset;
              currentRevision =
                spans.find(
                  (span) => span.story === head.story && span.start <= offset && offset <= span.end
                )?.revisionId ?? null;
            } catch {
              currentRevision = null;
            }
          }
          return currentRevision;
        },
        get styles() {
          const key = current.document?.package.styles;
          const cachedStyles = styleCache.current;
          if (cachedStyles && cachedStyles.key === key && cachedStyles.t === t) {
            return cachedStyles.options;
          }
          const options = styleOptions(current.document, t);
          styleCache.current = { key, t, options };
          return options;
        },
        get fonts() {
          const cached = fontCache.current;
          if (
            cached &&
            cached.families === current.fontFamilies &&
            cached.documentFonts === current.documentFonts
          ) {
            return cached.options;
          }
          const options = fontOptions(current.fontFamilies, current.documentFonts);
          fontCache.current = {
            families: current.fontFamilies,
            documentFonts: current.documentFonts,
            options,
          };
          return options;
        },
        translate: (key) => t(key),
      };
    };

    const format = (action: Parameters<PagedEditorCommandBridge['format']>[0]) =>
      executed(bridge()?.format(action) ?? false);

    const perform = <K extends DocxCommandId>(
      id: K,
      rawArgs: DocxCommandArgs[K],
      env: DocxCommandEnvironment
    ): DocxCommandResult | Promise<DocxCommandResult> => {
      const current = latest.current;
      const editor = bridge();
      const args = rawArgs as never;
      switch (id) {
        case 'undo':
        case 'redo':
          return executed(editor?.history(id === 'redo') ?? false);
        case 'bold':
        case 'italic':
        case 'underline':
        case 'strikethrough':
        case 'superscript':
        case 'subscript':
        case 'clearFormatting':
        case 'bulletList':
        case 'numberedList':
        case 'indent':
        case 'outdent':
        case 'setLtr':
        case 'setRtl':
          return format(id as 'bold');
        case 'paragraphStyle':
          return format({
            type: 'applyStyle',
            value: (args as DocxCommandArgs['paragraphStyle']).styleId,
          });
        case 'fontFamily':
          return format({ type: 'fontFamily', value: (args as DocxCommandArgs['fontFamily']).family });
        case 'fontSize':
          return format({ type: 'fontSize', value: (args as DocxCommandArgs['fontSize']).points });
        case 'textColor': {
          const { color } = args as DocxCommandArgs['textColor'];
          return format({
            type: 'textColor',
            value: color === 'auto' ? { auto: true } : (color as ColorValue),
          });
        }
        case 'highlightColor':
          return format({
            type: 'highlightColor',
            value: (args as DocxCommandArgs['highlightColor']).color,
          });
        case 'alignment':
          return format({ type: 'alignment', value: (args as DocxCommandArgs['alignment']).value });
        case 'lineSpacing':
          return format({
            type: 'lineSpacing',
            value: (args as DocxCommandArgs['lineSpacing']).value,
          });
        case 'insertLink': {
          const session = editor?.session();
          if (!session) return executed(false);
          const selectedText = yrsSelectedText(session);
          const existing = yrsHyperlinkAtSelection(session);
          openDialog('link');
          if (existing) {
            current.hyperlinkDialog.openEdit({
              url: existing.href,
              displayText: selectedText || existing.text,
              tooltip: existing.tooltip,
            });
          } else {
            current.hyperlinkDialog.openInsert(selectedText);
          }
          return OPENED;
        }
        case 'insertImage': {
          const picker = controller.defer('insertImage', null, 'selection');
          current.pickImage((image) =>
            picker.complete(() => {
              const target = bridge();
              if (!target) throw new DocxCommandAdmissionError('editor-unavailable');
              return executed(target.command({ type: 'insertImage', image }));
            })
          );
          return OPENED;
        }
        case 'insertTable': {
          const { rows, columns } = args as DocxCommandArgs['insertTable'];
          return executed(editor?.command({ type: 'insertTable', rows, columns }) ?? false);
        }
        case 'insertPageBreak':
          return executed(editor?.command({ type: 'insertPageBreak' }) ?? false);
        case 'insertSectionBreakNextPage':
        case 'insertSectionBreakContinuous':
          return executed(
            editor?.command({
              type: 'insertSectionBreak',
              breakType: id === 'insertSectionBreakNextPage' ? 'nextPage' : 'continuous',
            }) ?? false
          );
        case 'imageWrap':
        case 'imageTransform':
        case 'imageProperties': {
          const image = editor?.selectedImage();
          if (!image) return executed(false);
          if (id === 'imageProperties') {
            openDialog('imageProperties');
            current.openImageProperties(image);
            return OPENED;
          }
          if (id === 'imageTransform') {
            return executed(
              editor!.command({
                type: 'imageTransform',
                pmPos: image.pos,
                action: (args as DocxCommandArgs['imageTransform']).action,
              })
            );
          }
          const target = (args as DocxCommandArgs['imageWrap']).wrap;
          const initialPositionEmu =
            imageWrapTarget(image.attrs) === 'inline' &&
            target !== 'inline' &&
            current.displayListQueries
              ? captureInlinePositionEmuFromDisplayList(current.displayListQueries, image.pos)
              : undefined;
          return executed(
            editor!.command({
              type: 'imageWrap',
              pmPos: image.pos,
              target,
              options: initialPositionEmu ? { initialPositionEmu } : undefined,
            })
          );
        }
        case 'tableAction': {
          const action = args as DocxTableAction;
          const outcome = current.tableAction(action);
          if (outcome !== 'opened') return executed(outcome);
          openDialog(action === 'splitCell' ? 'splitCell' : 'tableProperties');
          return OPENED;
        }
        case 'pageSetup':
          openDialog('pageSetup');
          current.openPageSetup();
          return OPENED;
        case 'watermark':
          openDialog('watermark');
          current.openWatermark();
          return OPENED;
        case 'editingMode': {
          const { mode } = args as DocxCommandArgs['editingMode'];
          if (mode === current.mode) return executed(false);
          current.setEditingMode(mode);
          if (mode === 'suggesting') current.setShowCommentsSidebar(true);
          return { ok: true, status: current.modeControlled ? 'requested' : 'executed' };
        }
        case 'reviewAccept':
        case 'reviewReject': {
          const session = editor?.session();
          const target = args as DocxCommandArgs['reviewAccept'];
          const revisionId = target?.revisionId ?? env.currentRevisionId;
          if (!session || !revisionId) return executed(false);
          if (id === 'reviewAccept') session.acceptChange({ revisionId });
          else session.rejectChange({ revisionId });
          current.pagedEditorRef.current?.syncYrsInputState(true);
          current.refreshTrackedChanges(session);
          return executed(true);
        }
        case 'reviewPrevious':
        case 'reviewNext': {
          const session = editor?.session();
          const selection = current.pagedEditorRef.current?.getSelectionRange();
          const navigation = env.revisions;
          if (!session || navigation.length === 0) return executed(false);
          const caret = selection ? (id === 'reviewNext' ? selection.to : selection.from) : -1;
          const next =
            id === 'reviewNext'
              ? (navigation.find((revision) => revision.position > caret) ?? navigation[0])
              : ([...navigation].reverse().find((revision) => revision.position < caret) ??
                navigation[navigation.length - 1]);
          const revision = session
            .listRevisions()
            .find((candidate) => candidate.revisionId === next.revisionId);
          if (!revision) return executed(false);
          const story = revision.range.story;
          return executed(
            editor!.select(
              { story, ...revision.range.start },
              { story, ...revision.range.end }
            )
          );
        }
        case 'commentsSidebar':
          current.setShowCommentsSidebar((open) => !open);
          current.setExpandedSidebarItem(null);
          return { ok: true, status: current.sidebarControlled ? 'requested' : 'executed' };
        case 'open':
          current.openDocument();
          return OPENED;
        case 'save':
          return current.save().then(
            (outcome): DocxCommandResult =>
              outcome === 'saved'
                ? { ok: true, status: 'executed' }
                : outcome === 'requested'
                  ? { ok: true, status: 'requested' }
                  : {
                      ok: false,
                      failure: {
                        code: 'command-failed',
                        message: env.translate('commands.reasons.commandFailed'),
                      },
                    }
          );
        case 'print': {
          const job = current.reservePrint();
          const session = editor?.session() ?? null;
          const assertCurrent = () => {
            if (bridge()?.session() !== session) {
              throw new DocxCommandAdmissionError('document-replaced');
            }
          };
          return (async (): Promise<DocxCommandResult> => {
            try {
              if (!editor || !session) throw new DocxCommandAdmissionError('editor-unavailable');
              await editor.runAfterPendingInput(() => undefined);
              const displayList = await latest.current.renderedDisplayList();
              assertCurrent();
              await job.prepare(displayList);
              assertCurrent();
              return executed(job.print());
            } catch (error) {
              job.cancel();
              throw error;
            }
          })();
        }
        case 'find':
        case 'replace': {
          const session = editor?.session();
          const selectedText = session ? yrsSelectedText(session) : '';
          openDialog('replace');
          if (id === 'find') current.findReplace.openFind(selectedText);
          else current.findReplace.openReplace(selectedText);
          return OPENED;
        }
        case 'reportIssue':
          openReportIssue();
          return OPENED;
        case 'zoom': {
          const { scale } = args as DocxCommandArgs['zoom'];
          if (scale === current.zoom) return executed(false);
          current.setZoom(scale);
          return executed(true);
        }
        default:
          return {
            ok: false,
            failure: {
              code: 'unsupported-command',
              message: env.translate('commands.reasons.unsupportedCommand'),
            },
          };
      }
    };

    return {
      environment,
      ordered: (id, args) =>
        !IMMEDIATE_COMMANDS.has(id) &&
        !(
          id === 'tableAction' &&
          (args === 'splitCell' ||
            (typeof args === 'object' &&
              args !== null &&
              (args as { type?: string }).type === 'openTableProperties'))
        ),
      admit(operation) {
        const editor = bridge();
        if (!editor) return Promise.reject(new DocxCommandAdmissionError('editor-unavailable'));
        return editor.runAfterPendingInput(operation);
      },
      perform,
      capture(target): EditorOrigin | null {
        const session = bridge()?.session() ?? null;
        if (!session) return null;
        return {
          document: session,
          selection: target === 'selection' ? session.encodeSelection() : null,
          cells: target === 'selection' ? yrsTableSelectionStories(session) : null,
        };
      },
      resume(origin) {
        const { document, selection, cells } = origin as EditorOrigin;
        const session = bridge()?.session() ?? null;
        if (!session) return 'editor-unavailable';
        if (session !== document) return 'document-replaced';
        if (selection) {
          const opened = session.resolveSelection(selection);
          if (!opened || !sameYrsSelection(opened, session.selection())) return 'target-changed';
        }
        if (cells && !sameCells(cells, yrsTableSelectionStories(session))) return 'target-changed';
        return null;
      },
      chrome: () => ({
        i18n: latest.current.i18n,
        isDark: latest.current.isDark,
        theme: latest.current.theme,
      }),
      focusEditor: () => latest.current.pagedEditorRef.current?.focus(),
    };
  }, []);

  useEffect(() => {
    controller.attach(binding);
    return () => controller.detach(binding);
  }, [binding, controller]);

  const { session } = inputs;
  useEffect(() => {
    if (!session) return;
    documentVersion.current += 1;
    return session.onUpdate(() => {
      documentVersion.current += 1;
    });
  }, [session]);

  const subscription = useRef<{ bridge: PagedEditorCommandBridge; release: () => void } | null>(
    null
  );
  useEffect(() => {
    const bridge = latest.current.bridgeRef.current;
    if (subscription.current?.bridge !== bridge) {
      subscription.current?.release();
      subscription.current = bridge
        ? { bridge, release: bridge.subscribe(controller.refresh) }
        : null;
    }
    controller.refresh();
  });
  useEffect(
    () => () => {
      subscription.current?.release();
      subscription.current = null;
    },
    []
  );

  return useMemo(() => {
    const complete: DocxCommandsHandle['complete'] = (dialog, write) => {
      const continuation = continuations.current.get(dialog);
      if (continuation) return continuation.complete(write);
      return Promise.resolve({
        ok: false,
        failure: {
          code: 'editor-unavailable',
          message: translationRef.current('commands.reasons.editorUnavailable'),
        },
      });
    };
    return { controller, open: openDialog, complete };
  }, [controller, openDialog]);
}

import type {
  CollaborationReplica, DeckSnapshot, HitTestResult, PptxEditRequest, PptxEditResult, PptxFindRequest, PptxFindResult,
  PptxFontFace, PptxPresence, PptxReadRequest, PptxReadResult, PptxValidationResult, PresentationHandle,
} from '@betteroffice/pptx';
import type { Translations } from '@betteroffice/pptx-i18n';
import { Component, lazy, Suspense, useEffect, useMemo, useState, type CSSProperties, type JSX, type ReactNode } from 'react';
import { LocaleProvider, useTranslation } from './i18n';
import type { PptxCommandStore } from './commands/types';
import type { PptxEditorPluginProps, PptxPlugin } from './plugins/types';
import type { PptxWorkerEditorProps } from './worker/PptxWorkerEditor';
import type { PptxWorkerEditorApi } from './worker/createWorkerEditorApi';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import { LEGACY_SYNC_API } from './runtimeOptions';
import { warnPptxEngine } from './engineWarnings';
import { PeerlessPptxEditorLoader } from './PeerlessPptxEditorLoader';
import { PresentationFailure } from './components/PresentationFailure';

export const pptxPeerlessEditorLoader = { load: () => import('./worker/PeerlessPptxSurface') };
export const pptxLegacyEditorLoader = { load: () => import('./PptxLegacyEditorSession') };
// Browsers cache a failed module fetch for the page's lifetime, so only a reload fetches the chunk again.
const reloadPage = () => globalThis.location?.reload();

export class PptxWorkerEditorChunkLoadError extends Error {
  readonly code = 'chunk-load-failed';
  constructor(readonly cause: unknown) {
    super('Failed to load the editor code. Reload the page to retry.');
    this.name = 'PptxWorkerEditorChunkLoadError';
  }
}

type PeerlessEditorBoundaryProps = Pick<PptxAsyncEditorProps, 'onError' | 'className' | 'style'> & {
  children: ReactNode;
};

function PeerlessEditorError({ error, className, style }: { error: Error; className?: string; style?: CSSProperties }) {
  const chunk = error instanceof PptxWorkerEditorChunkLoadError;
  return <div dir="ltr" className={className} style={{ position: 'relative', width: '100%', height: '100%', ...style }}>
    <PresentationFailure kind={chunk ? 'load' : 'open'} details={error.message} onReload={chunk ? reloadPage : undefined} />
  </div>;
}

class PeerlessEditorBoundary extends Component<PeerlessEditorBoundaryProps, { error: Error | null }> {
  state = { error: null as Error | null };
  private reported: Error | null = null;

  static getDerivedStateFromError(error: unknown) {
    return { error: error instanceof Error ? error : new Error(String(error)) };
  }

  componentDidCatch() {
    const error = this.state.error;
    if (!error || this.reported === error) return;
    this.reported = error;
    try { this.props.onError?.(error); } catch {}
  }

  render() {
    return this.state.error ? <PeerlessEditorError error={this.state.error} className={this.props.className}
      style={this.props.style} /> : this.props.children;
  }
}

export interface PptxTextSelection {
  shapeId: string;
  storyId: string;
  anchor: number;
  focus: number;
  focusLine?: number;
}

export interface PptxTextSelectionTarget {
  /** 1-based slide number. */
  slide: number;
  shapeId: string;
  storyId: string;
  start: number;
  end: number;
}

export type PptxPointPosition = HitTestResult & { slide: number; slideId: string };

export interface PptxEditorApi {
  /** Waits for accepted input; rejects during unfinished pointer gestures. */
  flushPendingInput: () => Promise<void>;
  /** Client coordinates; returns null outside slide content. */
  getPositionAtPoint: (clientX: number, clientY: number) => PptxPointPosition | null;
  clearSelection: () => void;
  focus: () => void;
  /** Accepts a 1-based slide number. */
  goToSlide: (slide: number) => boolean;
  handle: PresentationHandle;
  /**
   * The editor's commands, shared by built-in and host chrome; `api.handle` edits bypass them.
   * @experimental
   */
  readonly commands: PptxCommandStore;
  refresh: () => void;
  refreshProposals: () => void;
  /** Serialize the presentation back to .pptx bytes, edits included. */
  save: () => Uint8Array;
  /** Also navigates to the slide. */
  selectText: (target: PptxTextSelectionTarget) => boolean;
  /** Flushes pending input, then returns the session version. */
  version: () => Promise<string>;
  /** Flushes pending input, then reads slides and story text with the version they were read at. */
  readContent: (request?: PptxReadRequest) => Promise<PptxReadResult>;
  /** Flushes pending input, then searches exactly and case-sensitively. */
  findText: (request: PptxFindRequest) => Promise<PptxFindResult>;
  /** Flushes pending input, then checks a batch without changing anything. */
  validateEdits: (request: PptxEditRequest) => Promise<PptxValidationResult>;
  /**
   * Flushes pending input, then applies every step or none against `expectVersion` and refreshes
   * the editor once. A refusal is data and never rolls back the flushed input. Read-only editors
   * refuse with `read-only`. Rejects when the presentation is replaced while input flushes.
   */
  applyEdits: (request: PptxEditRequest) => Promise<PptxEditResult>;
}

/** @experimental */
export interface PptxWorkerViewerApi extends Omit<
  PptxEditorApi, 'handle' | 'save' | 'getPositionAtPoint' | 'selectText' |
  'version' | 'readContent' | 'findText' | 'validateEdits' | 'applyEdits'
> {
  handle: null;
  save: () => null;
  getPositionAtPoint: (clientX: number, clientY: number) => null;
  selectText: (target: PptxTextSelectionTarget) => false;
  saveAsync: () => Promise<Uint8Array | null>;
  version: () => Promise<string | null>;
  readContent: (request?: PptxReadRequest) => Promise<PptxReadResult | null>;
  findText: (request: PptxFindRequest) => Promise<PptxFindResult | null>;
  validateEdits: (request: PptxEditRequest) => Promise<PptxValidationResult | null>;
  applyEdits: (request: PptxEditRequest) => Promise<PptxEditResult | null>;
  /** Accepts a 1-based slide number; resolves after painting. */
  goToSlideAsync: (slide: number) => Promise<boolean>;
}

export interface PptxAsyncCollaborationOptions extends Omit<PptxEditorCollaborationOptions, 'clientId' | 'onReplica'> {
  /** Generated for each opened session when omitted. */
  clientId?: number;
  /** Change to open a different collaboration room. */
  roomId?: string;
  onReplica?: (replica: import('@betteroffice/pptx').AsyncCollaborationReplica | null) => void;
}

/** @internal */
export type PptxContentCollaborationOptions = Omit<PptxEditorCollaborationOptions, 'clientId'> & { clientId?: number };

export interface PptxEditorCollaborationOptions {
  clientId: number;
  initialUpdate?: Uint8Array;
  onReplica?: (replica: CollaborationReplica | null) => void;
  presence?: PptxPresence;
}

export interface PptxInThreadEditorProps extends PptxEditorPluginProps {
  file?: Uint8Array;
  /** Font faces; equivalent inline arrays do not reopen the presentation. */
  fonts: ReadonlyArray<PptxFontFace>;
  clientId?: number;
  /** Uses the deprecated in-thread engine when set. */
  collaboration?: PptxEditorCollaborationOptions;
  /** @deprecated Omit this prop to use the default worker-owned editor. */
  experimentalWorkerOpen?: false;
  i18n?: Translations;
  className?: string;
  /** Download name for the save button; falls back to `presentation.pptx`. */
  fileName?: string;
  /** 1-based; clamped to the deck. */
  initialSlide?: number;
  /** Called once an opened deck's first slide is painted. */
  onReady?: (api: PptxEditorApi) => void;
  onChange?: (snapshot: DeckSnapshot) => void;
  onError?: (error: Error) => void;
  /** Receives the saved bytes; without it, saving downloads the file. */
  onSave?: (bytes: Uint8Array) => void;
  /** Return true for built-in saving; false or void handles/cancels the request. */
  onSaveRequest?: () => boolean | void | Promise<boolean | void>;
  /** Blocks user edits; navigation and selection remain available. */
  readOnly?: boolean;
  /**
   * Replaces the toolbar: omit it for the default controls, pass `null` for
   * none, or pass chrome composed from the toolbar parts, which also renders
   * while `readOnly`. `showToolbar={false}` hides the whole region.
   * @experimental
   */
  toolbar?: ReactNode;
  /** Shows the toolbar region; defaults to true. */
  showToolbar?: boolean;
}

/** @deprecated Removed in 0.6.0. Use PptxAsyncEditorProps. */
export type PptxWorkerViewerProps = Omit<
  PptxInThreadEditorProps, 'onReady' | 'readOnly' | 'experimentalWorkerOpen' | 'collaboration' | 'plugins' | 'onChange'
> & {
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]: true;
  collaboration?: undefined;
  plugins?: readonly [];
  onChange?: undefined;
  readOnly: true;
  /** Opens in the worker; without it the legacy runtime runs in-thread. */
  experimentalWorkerOpen: true;
  /** Without `Worker`, falls back to the in-thread engine and passes `PptxEditorApi`. */
  onReady?: (api: PptxWorkerViewerApi) => void;
};

export interface PptxDocumentChange {
  version: string;
}

export interface PptxAsyncPresentationProps {
  /** Receives the version after a committed edit changes the presentation. */
  onDocumentChange?: (change: PptxDocumentChange) => void;
  /** Whether Save also downloads a copy. Defaults to true without onSave, false with it. */
  downloadOnSave?: boolean;
  /** Initial zoom scale, clamped to 0.1 to 5; omit to fit the slide. */
  initialZoom?: number;
  /** Rendered while the presentation opens. */
  loadingIndicator?: ReactNode;
  /** Inline styles for the editor root. */
  style?: CSSProperties;
  /**
   * Loads bundled faces before the first paint for document families that
   * `fonts` does not supply; supplied families keep the app's faces. Pass
   * false to fetch no default fonts. Defaults to true.
   */
  defaultFonts?: boolean;
}

export type PptxAsyncEditorProps = Omit<PptxInThreadEditorProps,
  'file' | 'onReady' | 'collaboration' | 'experimentalWorkerOpen'> & PptxAsyncPresentationProps & {
  file?: Uint8Array | ArrayBuffer;
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]?: false;
  /** @deprecated False selects the in-thread engine. */
  experimentalWorkerOpen?: boolean;
  collaboration?: PptxAsyncCollaborationOptions;
  onReady?: (api: PptxAsyncEditorApi) => void | (() => void);
};

export type PptxAsyncEditorInternalProps = Omit<PptxAsyncEditorProps, 'file'> & { file?: Uint8Array };

/** @deprecated Removed in 0.6.0. Use PptxAsyncEditorProps. */
export type PptxLegacyEditorProps = PptxInThreadEditorProps & {
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]: true;
};

type WorkerSessionProps = Omit<PptxWorkerEditorProps, 'readOnly' | 'onReady' | 'experimentalWorkerOpen'> & {
  experimentalWorkerOpen: true;
  readOnly?: boolean;
  onReady?: (api: PptxWorkerEditorApi | PptxWorkerViewerApi) => void;
};

type WorkerEditorSessionProps = Omit<PptxWorkerEditorProps, 'readOnly'> & { readOnly?: boolean };
type CollaborationSessionProps = Omit<PptxLegacyEditorProps, 'experimentalWorkerOpen'> & {
  collaboration: PptxEditorCollaborationOptions;
  experimentalWorkerOpen?: boolean;
};
type GeneralEditorSessionProps = Omit<PptxLegacyEditorProps, 'experimentalWorkerOpen' | 'onReady'> & {
  experimentalWorkerOpen?: boolean;
  onReady?: (api: PptxEditorApi | PptxWorkerEditorApi | PptxWorkerViewerApi) => void;
};

export type PptxEditorProps = PptxAsyncEditorProps | GeneralEditorSessionProps;

export function PptxEditor(props: PptxAsyncEditorProps): JSX.Element;
export function PptxEditor(props: PptxLegacyEditorProps): JSX.Element;
export function PptxEditor(props: CollaborationSessionProps): JSX.Element;
export function PptxEditor(props: PptxWorkerEditorProps): JSX.Element;
export function PptxEditor(props: WorkerEditorSessionProps & { onChange: NonNullable<PptxInThreadEditorProps['onChange']> }): JSX.Element;
export function PptxEditor(props: WorkerEditorSessionProps & { plugins: readonly [PptxPlugin, ...PptxPlugin[]] }): JSX.Element;
export function PptxEditor(props: PptxWorkerViewerProps): JSX.Element;
export function PptxEditor(props: WorkerSessionProps): JSX.Element;
export function PptxEditor(props: PptxEditorProps): JSX.Element;
export function PptxEditor({ i18n, ...props }:
  PptxAsyncEditorProps | WorkerEditorSessionProps | PptxWorkerViewerProps | WorkerSessionProps |
  CollaborationSessionProps | PptxLegacyEditorProps | GeneralEditorSessionProps): JSX.Element {
  const legacy = props[LEGACY_SYNC_API] === true;
  return (
    <LocaleProvider i18n={i18n}>
      {legacy ? (
        <PptxLegacySessionLoader editorProps={{ ...props as GeneralEditorSessionProps, i18n }} />
      ) : (
        <PptxAsyncEditorRoot {...props as PptxAsyncEditorProps} i18n={i18n} />
      )}
    </LocaleProvider>
  );
}

function PptxAsyncEditorRoot(props: PptxAsyncEditorProps) {
  const file = useMemo(() => props.file instanceof ArrayBuffer ? new Uint8Array(props.file) : props.file, [props.file]);
  return <PptxAsyncEditorSession {...props} file={file} />;
}

function PptxInThreadSessionLoader({ fallback, editorProps }: { fallback: ReactNode; editorProps: PptxAsyncEditorInternalProps }) {
  const [Session] = useState(() => lazy(() => import('./PptxAsyncInThreadSession').then(
    (module) => ({ default: module.PptxAsyncInThreadSession }),
    (cause: unknown) => { throw new PptxWorkerEditorChunkLoadError(cause); })));
  return <Suspense fallback={fallback}><Session {...editorProps} /></Suspense>;
}

function PptxAsyncEditorSession(props: PptxAsyncEditorInternalProps) {
  const inThread = props.experimentalWorkerOpen === false;
  const identity = useMemo(() => ({}), [props.file, inThread]);
  const start = () => ({ identity, reason: inThread ? 'deprecated' as const
    : typeof Worker !== 'function' ? 'no-worker' as const
    : typeof OffscreenCanvas !== 'function' ? 'no-offscreen-canvas' as const : undefined });
  const [choice, setChoice] = useState(start);
  const engine = choice.identity === identity ? choice : start();
  if (choice !== engine) setChoice(engine);
  useEffect(() => { if (engine.reason) warnPptxEngine(engine.reason); }, [engine.reason]);
  const fallback = <PptxEditorFallback {...props} />;
  return <PeerlessEditorBoundary onError={props.onError} className={props.className} style={props.style}>
    {engine.reason ? <PptxInThreadSessionLoader fallback={fallback} editorProps={props} />
      : <PeerlessPptxEditorLoader load={pptxPeerlessEditorLoader.load} identity={identity}
        ErrorType={PptxWorkerEditorChunkLoadError} fallback={fallback} editorProps={props} />}
  </PeerlessEditorBoundary>;
}

function PptxEditorFallback(props: Pick<PptxAsyncEditorProps, 'file' | 'loadingIndicator' | 'className' | 'style'>) {
  const { t } = useTranslation();
  return <div dir="ltr" className={props.className} style={props.style} role="status">
    {props.file && props.loadingIndicator !== undefined ? props.loadingIndicator : t('worker.preparing')}
  </div>;
}

function PptxLegacySessionLoader({ editorProps }: { editorProps: GeneralEditorSessionProps }) {
  const [Session] = useState(() => lazy(() => pptxLegacyEditorLoader.load().then(
    (module) => ({ default: module.PptxEditorSession }),
    (cause: unknown) => { throw new PptxWorkerEditorChunkLoadError(cause); })));
  return <PeerlessEditorBoundary onError={editorProps.onError} className={editorProps.className}>
    <Suspense fallback={<PptxEditorFallback {...editorProps} />}><Session {...editorProps} /></Suspense>
  </PeerlessEditorBoundary>;
}

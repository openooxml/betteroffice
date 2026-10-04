import type { MethodPolicies, SessionState } from '../../../../shared/office-session';
import type {
  PptxEditRequest,
  PptxEditResult,
  PptxFindRequest,
  PptxFindResult,
  PptxReadRequest,
  PptxReadResult,
  PptxValidationResult,
} from '../edits';
import type { OpenPresentationOptions } from '../wasm/loader';

/**
 * A transferable font face.
 * @experimental
 */
export interface PresentationSessionFont {
  family: string;
  bytes: ArrayBuffer;
  bold?: boolean;
  italic?: boolean;
}

/**
 * Structured-cloneable presentation open options.
 * @experimental
 */
export interface PresentationSessionOpenOptions
  extends Omit<OpenPresentationOptions, 'fonts' | 'fallbackFonts'> {
  fonts?: readonly PresentationSessionFont[];
  fallbackFonts?: readonly PresentationSessionFont[];
  wasm?: ArrayBuffer | WebAssembly.Module;
}

/**
 * Slide metadata without layout; indices are zero-based.
 * @experimental
 */
export interface PresentationSlideSummary {
  id: string;
  index: number;
  name: string | null;
  layoutPartPath: string | null;
}

/**
 * Presentation projection; version counts applied batches and size is in EMU.
 * @experimental
 */
export type PresentationSessionState = Omit<SessionState, 'stage'> & {
  format: 'pptx';
  stage: 'preview' | 'ready' | 'failed';
  slides: PresentationSlideSummary[];
  size: { width: number; height: number };
};

/**
 * Presentation change notifications after applied batches.
 * @experimental
 */
export interface PresentationSessionEvents {
  changed: { version: number; dirty: boolean };
}

/**
 * Presentation RPC methods; engine version tokens remain strings.
 * @experimental
 */
export type PresentationSessionMethods = {
  open(bytes: ArrayBuffer, options?: PresentationSessionOpenOptions): PresentationSessionState;
  version(): string;
  readContent(request?: PptxReadRequest): PptxReadResult;
  findText(request: PptxFindRequest): PptxFindResult;
  validateEdits(request: PptxEditRequest): PptxValidationResult;
  applyEdits(request: PptxEditRequest): PptxEditResult;
  slides(): PresentationSlideSummary[];
  slideSize(): { width: number; height: number };
  save(): ArrayBuffer;
  dispose(): void;
};

/** Exhaustive RPC names used by the session client. */
export const PRESENTATION_SESSION_METHODS = {
  open: true,
  version: true,
  readContent: true,
  findText: true,
  validateEdits: true,
  applyEdits: true,
  slides: true,
  slideSize: true,
  save: true,
  dispose: true,
} satisfies { readonly [K in keyof PresentationSessionMethods]-?: true };

/** Non-reorderable policies preserve call order across all lanes. */
export const PRESENTATION_SESSION_POLICIES: MethodPolicies<PresentationSessionMethods> = {
  open: { lane: 'interactive' },
  version: { lane: 'interactive' },
  readContent: { lane: 'interactive' },
  findText: { lane: 'interactive' },
  validateEdits: { lane: 'interactive' },
  applyEdits: { lane: 'input', mutates: true, userInput: true },
  slides: { lane: 'interactive' },
  slideSize: { lane: 'interactive' },
  save: { lane: 'interactive' },
  dispose: { lane: 'interactive' },
};

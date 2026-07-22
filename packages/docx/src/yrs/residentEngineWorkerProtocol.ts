import type {
  YrsEngineApplyProfile,
  YrsResidentCaretSnapshot,
  YrsResidentWorkerSnapshot,
  YrsSelection,
} from './index';
import type { ResidentCaretPaintStyle } from './residentCaret';

export type ResidentEngineWorkerRequest =
  | {
      id: number;
      type: 'bootstrap';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
    }
  | {
      id: number;
      type: 'sync';
      snapshot: YrsResidentWorkerSnapshot;
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'buildFrame';
      extras: string;
      expectedFrameEpoch: number;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyInput';
      text: string;
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyDelete';
      direction: 'backward' | 'forward';
      selection: YrsSelection;
      expectedFrameEpoch: number;
      profile: boolean;
      paintCaret: boolean;
    }
  | {
      id: number;
      type: 'applyUpdate';
      update: Uint8Array;
      selection: YrsSelection | null;
    }
  | {
      id: number;
      type: 'attachCanvases';
      pages: Array<{ pageId: string; canvas: OffscreenCanvas }>;
      activePageIds: string[];
      devicePixelRatio: number;
      zoom: number;
      caretStyle: ResidentCaretPaintStyle;
    }
  | { id: number; type: 'eraseCaret' }
  | { id: number; type: 'destroy' };

export type ResidentEngineWorkerRequestWithoutId = ResidentEngineWorkerRequest extends infer Request
  ? Request extends { id: number }
    ? Omit<Request, 'id'>
    : never
  : never;

export type ResidentEngineWorkerResponse =
  | {
      id: number;
      ok: true;
      frame?: ArrayBuffer;
      updates?: ArrayBuffer[];
      engineMs?: number;
      workerTotalMs?: number;
      engineProfile?: YrsEngineApplyProfile;
      caret?: YrsResidentCaretSnapshot;
      selection?: YrsSelection | null;
      /** The presented frame carries the worker-painted caret line. */
      caretPainted?: boolean;
      replayMs?: number;
      replayedPages?: number;
      layoutRevision?: number;
      /** The worker replica's yrs state vector after this operation, so the
       * next sync can ship a diff instead of the whole document state. */
      stateVector?: ArrayBuffer;
    }
  | {
      id: number;
      ok: false;
      error: string;
      residentUnavailable?: boolean;
    };

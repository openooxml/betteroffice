import { createSessionHost, type SessionHost } from '../../../../shared/office-session/host';
import { transferable } from '../../../../shared/office-session/protocol';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure, type MethodHandlers, type MethodPolicy } from '../../../../shared/office-session/types';
import { wasmAssetUrl } from '../wasm/asset';
import {
  initWasm, openWorkbook, openWorkbookPeer, StaleProposalError, workbookDisplayListJson, workbookPeerHydration,
  workbookPeerSnapshot,
  type SheetInfo, type WorkbookCalculationContext, type WorkbookHandle,
} from '../wasm/loader';
import { localNowSerial } from './calculationClock';
import {
  WORKBOOK_SESSION_POLICIES,
  type WorkbookSessionEvents,
  type WorkbookSessionMethods,
  type WorkbookSheetSummary,
} from './methods';
import {
  applyWorkbookReplayOp,
  validateWorkbookReplayEnvelope,
  workbookReplayRefused,
  WORKBOOK_INTERNAL_SESSION_POLICIES,
  type WorkbookInternalSessionMethods,
  type WorkbookReplayEnvelope,
} from './replay';
import { WorkbookPeerHydrationError } from './peerHydrationError';

type Events = { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] } & {
  peerOpened: { version: string; initialCalculation?: WorkbookCalculationContext | null };
};
type Methods = WorkbookSessionMethods & WorkbookInternalSessionMethods;

function sheets(info: SheetInfo): WorkbookSheetSummary[] {
  return info.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] }));
}

export function createWorkbookSessionHost(
  transport: SessionTransport,
  options: {
    initWasm?: (source?: ArrayBuffer | WebAssembly.Module) => Promise<void>;
    wasmModule?(): WebAssembly.Module | undefined;
  } = {}
): SessionHost<Events> {
  let handle: WorkbookHandle | undefined;
  let disposed = false;
  let version = 0;
  let dirty = false;
  let epoch = 0;
  let sequence = 0;
  let revision = 0;
  let retainedHydration = false;
  let peerAttached = false;
  let previewSource: { bytes: Uint8Array; hydration: string } | undefined;
  let peerSnapshot: {
    version: string; sequence: number; records: number; bytes: number; pending?: Uint8Array;
  } | undefined;
  const committed: WorkbookReplayEnvelope[] = [];
  const encoder = new TextEncoder();

  function workbook(): WorkbookHandle {
    if (disposed) throw new Error('Workbook session is disposed');
    if (!handle) throw new Error('Workbook session is not open');
    return handle;
  }

  function dispose(): void {
    disposed = true;
    const opened = handle;
    handle = undefined;
    opened?.dispose();
    previewSource = undefined;
    peerSnapshot = undefined;
    committed.length = 0;
  }

  function checkSheet(opened: WorkbookHandle, sheet: number): void {
    if (!Number.isInteger(sheet) || sheet < 0 || sheet >= opened.sheetCount()) {
      throw new RangeError('Sheet index is out of range');
    }
  }

  const internalHandlers: MethodHandlers<WorkbookInternalSessionMethods, null> = {
    beginPeerSnapshot(_, records, bytes) {
      const opened = workbook();
      peerSnapshot = undefined;
      if (!retainedHydration || peerAttached) throw new Error('Worker peer snapshot requires pending retained hydration');
      try {
        workbookPeerSnapshot(opened).begin(records, bytes);
        peerSnapshot = { version: opened.version(), sequence, records, bytes };
        return { version: peerSnapshot.version, sequence };
      } catch (error) {
        try { workbookPeerSnapshot(opened).end(); } catch {}
        throw new Error(error instanceof Error ? error.message : String(error));
      }
    },
    pullPeerSnapshot() {
      const opened = workbook();
      try {
        const snapshot = checkPeerSnapshot(opened);
        const buffers: ArrayBuffer[] = [];
        let bytes = 0;
        while (buffers.length < snapshot.records && bytes < snapshot.bytes) {
          const chunk = snapshot.pending ?? workbookPeerSnapshot(opened).next();
          snapshot.pending = undefined;
          if (chunk === undefined) break;
          const buffer = chunk.buffer;
          if (!(buffer instanceof ArrayBuffer) || chunk.byteOffset !== 0 || chunk.byteLength !== buffer.byteLength) {
            throw new Error('Worker peer snapshot chunk is not transferable');
          }
          if (buffers.length > 0 && bytes + chunk.byteLength > snapshot.bytes) {
            snapshot.pending = chunk;
            break;
          }
          buffers.push(buffer);
          bytes += chunk.byteLength;
        }
        return buffers.length === 0 ? undefined : transferable(buffers, buffers);
      } catch (error) {
        peerSnapshot = undefined;
        try { workbookPeerSnapshot(opened).end(); } catch {}
        throw new Error(error instanceof Error ? error.message : String(error));
      }
    },
    endPeerSnapshot(_, discard) {
      const opened = workbook();
      try {
        if (!discard) checkPeerSnapshot(opened);
        workbookPeerSnapshot(opened).end();
      } catch (error) {
        try { workbookPeerSnapshot(opened).end(); } catch {}
        throw new Error(error instanceof Error ? error.message : String(error));
      } finally {
        peerSnapshot = undefined;
      }
    },
    preview(_, viewport, sheet, ops) {
      const opened = workbook();
      checkSheet(opened, sheet);
      if (!previewSource) throw new Error('Worker preview requires retained hydration');
      const speculative = openWorkbookPeer(previewSource.bytes, {}, previewSource.hydration);
      try {
        for (const envelope of committed) {
          speculative.setCalculationContext(envelope.calculation);
          try { applyWorkbookReplayOp(speculative, envelope.op); }
          catch (error) { if (!(error instanceof StaleProposalError) || !envelope.staleProposal) throw error; }
        }
        for (const op of ops) {
          validateWorkbookReplayEnvelope({ sequence: 1, calculation: { nowSerial: 0, randSeed: 0 }, op });
          if (op.calculation) speculative.setCalculationContext(op.calculation);
          const result = applyWorkbookReplayOp(speculative, op);
          if (workbookReplayRefused(result)) throw new Error(`Preview refused: ${JSON.stringify(result)}`);
        }
        const buffer = encoder.encode(workbookDisplayListJson(speculative, viewport, sheet)).buffer;
        return transferable({ displayList: buffer, version: speculative.version(), epoch: 0,
          sequence, sheet, viewport, mergedRanges: speculative.visibleMergedRanges(sheet, viewport) }, [buffer]);
      } finally { speculative.dispose(); }
    },
    attachPeer(_, peerVersion, peerSequence) {
      peerAttached = false;
      if (peerVersion !== workbook().version() || peerSequence !== sequence) {
        throw new WorkbookPeerHydrationError('version-mismatch',
          'Workbook worker state differs from retained peer hydration');
      }
      peerAttached = true;
      previewSource = undefined;
      committed.length = 0;
    },
    detachPeer() {
      peerAttached = false;
    },
    replay(_, envelope) {
      validateWorkbookReplayEnvelope(envelope);
      if (retainedHydration && !peerAttached) {
        throw new WorkbookPeerHydrationError('mutation-outside-replay',
          'Workbook replay requires an attached edit peer');
      }
      const opened = workbook();
      if (envelope.sequence !== sequence + 1) {
        const error = new Error(`Expected workbook replay sequence ${sequence + 1}, got ${envelope.sequence}`);
        error.name = 'WorkbookReplayOrderError';
        throw error;
      }
      const op = envelope.op;
      try {
        const before = opened.version();
        opened.setCalculationContext(envelope.calculation);
        let result: ReturnType<typeof applyWorkbookReplayOp> = undefined;
        try {
          result = applyWorkbookReplayOp(opened, !retainedHydration && op.method === 'applyEdits' ? {
            method: 'applyEdits', args: [{ ...op.args[0], expectVersion: before }],
          } : op);
          if (envelope.staleProposal) throw new Error('Expected a stale proposal refresh');
        } catch (error) {
          if (!(error instanceof StaleProposalError) || !envelope.staleProposal ||
            JSON.stringify({ cells: error.cells, targets: error.targets }) !== JSON.stringify(envelope.staleProposal)) {
            throw error;
          }
        }
        if (workbookReplayRefused(result)) throw new Error(`Engine refused replay: ${JSON.stringify(result)}`);
        const changed = opened.version() !== before;
        sequence = envelope.sequence;
        if (previewSource) committed.push(structuredClone(envelope));
        if (changed) {
          revision += 1;
          version += 1;
          dirty = true;
          host.emit('changed', { version, dirty });
        }
        return { sequence, revision, version, result };
      } catch (error) {
        const message = `Workbook replay diverged at sequence ${envelope.sequence} (${op.method})`;
        throw new SessionFailure(
          error instanceof WebAssembly.RuntimeError ? 'trap' : 'crash', message,
          `${message}: ${error instanceof Error ? error.message : String(error)}`
        );
      }
    },
  };

  function checkPeerSnapshot(opened: WorkbookHandle) {
    if (!peerSnapshot) throw new Error('Worker peer snapshot is not active');
    if (peerSnapshot.version !== opened.version() || peerSnapshot.sequence !== sequence) {
      throw new Error('Workbook changed during peer snapshot');
    }
    return peerSnapshot;
  }

  const handlers: MethodHandlers<WorkbookSessionMethods, null> = {
    async open(_, bytes, input = {}) {
      if (disposed) throw new Error('Workbook session is disposed');
      if (handle) throw new Error('Workbook session is already open');
      const retainPeerHydration = 'retainPeerHydration' in input && input.retainPeerHydration === true;
      retainedHydration = retainPeerHydration;
      let wasm = input.wasm;
      if (retainPeerHydration && wasm instanceof ArrayBuffer) wasm = await WebAssembly.compile(wasm);
      await (options.initWasm ?? initWasm)(wasm);
      wasm ??= options.wasmModule?.();
      if (disposed) throw new Error('Workbook session is disposed');
      const calculation = input.calculation ?? (retainPeerHydration && !input.collaborative ? {
        nowSerial: localNowSerial(Date.now()),
        randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0,
      } : undefined);
      const opened = openWorkbook(new Uint8Array(bytes), {
        collaborative: input.collaborative,
        clientId: input.clientId,
        calculation,
      });
      try {
        let initialCalculation: WorkbookCalculationContext | null | undefined;
        if (retainPeerHydration) previewSource = {
          bytes: new Uint8Array(bytes).slice(), hydration: workbookPeerHydration(opened),
        };
        if (previewSource) {
          const hydration = JSON.parse(previewSource.hydration);
          initialCalculation = hydration.calculation_context === undefined ? undefined :
            hydration.calculation_context === null ? null : {
              nowSerial: hydration.calculation_context.now_serial, randSeed: hydration.workbook.rand_seed,
            };
        }
        if (retainPeerHydration && wasm instanceof WebAssembly.Module) {
          transport.post({
            protocol: 1, kind: 'wasm-module', url: wasmAssetUrl().href, module: wasm,
            hydration: previewSource!.hydration, version: opened.version(), sequence,
          });
        }
        const info = opened.sheetInfo();
        const summaries = sheets(info);
        handle = opened;
        host.emit('peerOpened', { version: opened.version(), initialCalculation });
        return { format: 'xlsx', stage: 'ready', version, dirty,
          sheets: summaries, activeSheet: info.activeSheet };
      } catch (error) {
        opened.dispose();
        throw error;
      }
    },
    version: () => workbook().version(),
    readCells: (_, request) => workbook().readCells(request),
    findText: (_, request) => workbook().findText(request),
    validateEdits: (_, request) => workbook().validateEdits(request),
    applyEdits(_, request) {
      const result = workbook().applyEdits(request);
      if (result.ok && result.applied) {
        version += 1;
        dirty = true;
        host.emit('changed', { version, dirty });
      }
      return result;
    },
    frame(_, viewport, options = {}) {
      const opened = workbook();
      const sheet = options.sheet === undefined ? opened.sheetInfo().activeSheet : options.sheet;
      checkSheet(opened, sheet);
      const json = workbookDisplayListJson(opened, viewport, options.sheet);
      const mergedRanges = opened.visibleMergedRanges(sheet, viewport);
      const buffer = encoder.encode(json).buffer;
      epoch += 1;
      return transferable({
        displayList: buffer, version: opened.version(), epoch, sequence, sheet, viewport, mergedRanges,
      }, [buffer]);
    },
    sheetView(_, sheet) {
      const opened = workbook();
      checkSheet(opened, sheet);
      const { contentWidth, contentHeight, frozenRows, frozenCols, initialScrollX, initialScrollY } =
        opened.sheetInfoFor(sheet);
      const edge = opened.cellRect(sheet, Math.max(0, frozenRows - 1), Math.max(0, frozenCols - 1));
      return {
        sheet, version: opened.version(), contentWidth, contentHeight, frozenRows, frozenCols,
        initialScrollX, initialScrollY,
        frozenWidth: frozenCols === 0 ? 0 : edge.x + edge.w,
        frozenHeight: frozenRows === 0 ? 0 : edge.y + edge.h,
      };
    },
    cellGeometry(_, sheet, row, col) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return {
        sheet, version: opened.version(), rect: opened.cellRect(sheet, row, col),
        scrollPosition: opened.cellPosition(sheet, row, col),
      };
    },
    cellInputs(_, sheet, range) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return { sheet, version: opened.version(), cells: opened.rangeCells(sheet, range) };
    },
    sheets: () => sheets(workbook().sheetInfo()),
    calculationStatus: () => workbook().calculationStatus(),
    save() {
      const bytes = workbook().save();
      const buffer = bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 &&
        bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : new Uint8Array(bytes).buffer;
      return transferable(buffer, [buffer]);
    },
    dispose() {
      workbook();
      dispose();
    },
  };

  for (const method of Object.keys(handlers) as (keyof WorkbookSessionMethods)[]) {
    const configured = WORKBOOK_SESSION_POLICIES[method];
    const handler = handlers[method] as (...args: unknown[]) => unknown;
    Object.defineProperty(handlers, method, { enumerable: true, value: (...args: unknown[]) => {
      const policy = typeof configured === 'function'
        ? (configured as (...args: unknown[]) => MethodPolicy)(...args.slice(1)) : configured;
      if (retainedHydration && policy.mutates) {
        throw new WorkbookPeerHydrationError('mutation-outside-replay',
          `Cannot call ${method} outside workbook edit-peer replay in retained hydration mode`);
      }
      return handler(...args);
    } });
  }

  const host = createSessionHost<Methods, Events, null>(transport, {
    handlers: { ...handlers, ...internalHandlers },
    policies: { ...WORKBOOK_SESSION_POLICIES, ...WORKBOOK_INTERNAL_SESSION_POLICIES },
    context: null, onDispose: dispose,
  });
  return host;
}

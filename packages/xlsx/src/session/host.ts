import { createCollaborationPipeHost } from '../../../../shared/collaboration-pipe/worker';
import type { PipeCommandPayload, PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import { CollaborationProvider, providerArrival, providerLeave } from '../collaboration/provider';
import { CollaborationError, type AwarenessCursor, type AwarenessPeer } from '../collaboration/types';
import { createSessionHost, type SessionHost } from '../../../../shared/office-session/host';
import { deferReply, isHostMessage, transferable } from '../../../../shared/office-session/protocol';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure, type MethodHandlers, type MethodPolicy } from '../../../../shared/office-session/types';
import type { XlsxEditRefusal } from '../edits';
import type { XlsxExportOptions, XlsxExportResult, XlsxStructuredContent } from '../exports';
import { wasmAssetUrl } from '../wasm/asset';
import {
  initWasm, isPngExportAvailable, openWorkbook, openWorkbookPeer, StaleProposalError, workbookDisplayListJson, workbookPeerHydration,
  workbookPeerSnapshot,
  openWorkbookRecoveryCheckpoint, adoptWorkbookPeerVersion,
  type SheetInfo, type WorkbookCalculationContext, type WorkbookHandle,
  type Viewport,
} from '../wasm/loader';
import { localNowSerial } from './calculationClock';
import { createWorkbookCommandState, projectEditState, translateWorkbookSheetIds, type WorkbookCommandState } from './commandHandlers';
import { createWorkbookCommandHost } from './commandHost';
import { isWorkbookCommandFrameRequest, WorkbookCommandRefusal } from './commandProtocol';
import {
  WORKBOOK_SESSION_POLICIES,
  WorkbookPngUnavailableError,
  type WorkbookCollaborationPipeReply,
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
import { decodeWorkbookRecovery, isCollaborativeWorkbookRecovery, restoreWorkbookCommandState, type WorkbookRecovery } from './recovery';

type Events = { [K in keyof WorkbookSessionEvents]: WorkbookSessionEvents[K] } & {
  peerOpened: { version: string; initialCalculation?: WorkbookCalculationContext | null };
  collaborationPipe: PipeOutputBatch<AwarenessPeer>[];
};
type Methods = WorkbookSessionMethods & WorkbookInternalSessionMethods;
type WorkerCollaborationProvider = CollaborationProvider & { onPeers: CollaborationProvider['onAwareness'] };

function sheets(info: SheetInfo): WorkbookSheetSummary[] {
  return info.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] }));
}

function commandId(): string {
  const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
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
  const commandIdentity = { session: commandId(), incarnation: commandId() };
  let commands: WorkbookCommandState | undefined;
  let commandHost: ReturnType<typeof createWorkbookCommandHost> | undefined;
  let collaborative = false;
  let offUpdates = () => {};
  let pipe: ReturnType<typeof createCollaborationPipeHost<WorkerCollaborationProvider, AwarenessCursor, AwarenessPeer>> | undefined;
  let pipeOutputTimer: ReturnType<typeof setTimeout> | undefined;
  let previewSource: { bytes: Uint8Array; hydration: string } | undefined;
  let peerSnapshot: {
    version: string; sequence: number; records: number; bytes: number; pending?: Uint8Array;
  } | undefined;
  const committed: WorkbookReplayEnvelope[] = [];
  const encoder = new TextEncoder();

  function schedulePipeOutput(): void {
    if (disposed || pipeOutputTimer !== undefined) return;
    pipeOutputTimer = setTimeout(() => {
      pipeOutputTimer = undefined;
      const output = pipe?.drain();
      if (output) host.emit('collaborationPipe', output.batches, output.transfer);
    }, 0);
  }

  function applyPipe(payload: PipeCommandPayload<AwarenessCursor>): void {
    if (!pipe) throw new WorkbookCommandRefusal('unsupported');
    pipe.run(payload);
  }

  function workbook(): WorkbookHandle {
    if (disposed) throw new Error('Workbook session is disposed');
    if (!handle) throw new Error('Workbook session is not open');
    return handle;
  }

  function dispose(): void {
    disposed = true;
    if (pipeOutputTimer !== undefined) clearTimeout(pipeOutputTimer);
    pipeOutputTimer = undefined;
    pipe?.dispose();
    pipe = undefined;
    const opened = handle;
    handle = undefined;
    offUpdates();
    offUpdates = () => {};
    opened?.dispose();
    previewSource = undefined;
    peerSnapshot = undefined;
    committed.length = 0;
    commands = undefined;
    commandHost = undefined;
  }

  function checkSheet(opened: WorkbookHandle, sheet: number): void {
    if (!Number.isInteger(sheet) || sheet < 0 || sheet >= opened.sheetCount()) {
      throw new RangeError('Sheet index is out of range');
    }
  }

  function orderedRequest<T, R>(request: T, read: (request: T) => R): R | XlsxEditRefusal {
    const opened = workbook();
    if (!commands?.throughSeq) return read(request);
    const positional = opened.sheetInfo().sheetIds;
    let resolved: T;
    try { resolved = translateWorkbookSheetIds(request, commands.sheetIds, positional, true); }
    catch (error) {
      if (!(error instanceof WorkbookCommandRefusal)) throw error;
      return { ok: false, version: opened.version(),
        failure: { code: 'missing-target', message: 'The command sheet no longer exists' } };
    }
    return translateWorkbookSheetIds(read(resolved), positional, commands.sheetIds);
  }

  function frame(viewport: Viewport, sheet: number) {
    const opened = workbook();
    checkSheet(opened, sheet);
    const buffer = encoder.encode(workbookDisplayListJson(opened, viewport, sheet)).buffer;
    epoch += 1;
    return { displayList: buffer, version: opened.version(), epoch,
      sequence: commands?.throughSeq || sequence, sheet, viewport,
      mergedRanges: opened.visibleMergedRanges(sheet, viewport) };
  }

  function renderPng(sheet: number, render: (opened: WorkbookHandle) => Uint8Array) {
    const opened = workbook();
    if (!isPngExportAvailable()) throw new WorkbookPngUnavailableError();
    checkSheet(opened, sheet);
    const active = opened.sheetInfo().activeSheet;
    try {
      if (sheet !== active) opened.setActiveSheet(sheet);
      return transferBytes(render(opened));
    } finally {
      if (sheet !== active) opened.setActiveSheet(active);
    }
  }

  function transferBytes(bytes: Uint8Array) {
    const buffer = bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 &&
      bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : new Uint8Array(bytes).buffer;
    return transferable(buffer, [buffer]);
  }

  function exported(options?: XlsxExportOptions): XlsxExportResult<XlsxStructuredContent> {
    const opened = workbook();
    const result = opened.exportStructured(options);
    if (!commands?.throughSeq) return result;
    const requestedMaxBytes = options?.maxBytes;
    const maxBytes = requestedMaxBytes !== undefined && Number.isFinite(requestedMaxBytes)
      ? requestedMaxBytes : 8 * 1024 * 1024;
    const translated = translateWorkbookSheetIds(result, opened.sheetInfo().sheetIds, commands.sheetIds);
    if (translated.ok && encoder.encode(JSON.stringify(translated.content)).byteLength > maxBytes) {
      return { ok: false, version: translated.version,
        failure: { code: 'limit-exceeded', target: null, message: 'The exported content exceeds maxBytes' } };
    }
    return translated;
  }

  function sheetView(sheet: number) {
    const opened = workbook();
    checkSheet(opened, sheet);
    const { contentWidth, contentHeight, frozenRows, frozenCols, initialScrollX, initialScrollY } =
      opened.sheetInfoFor(sheet);
    const edge = opened.cellRect(sheet, Math.max(0, frozenRows - 1), Math.max(0, frozenCols - 1));
    return { sheet, version: opened.version(), contentWidth, contentHeight, frozenRows, frozenCols,
      initialScrollX, initialScrollY,
      frozenWidth: frozenCols === 0 ? 0 : edge.x + edge.w,
      frozenHeight: frozenRows === 0 ? 0 : edge.y + edge.h };
  }

  function installCommandHost(recovery?: WorkbookRecovery): void {
    commandHost = createWorkbookCommandHost({ handle: workbook(), state: commands!, applyPipe,
      ...(recovery ? { revision: recovery.revision, versions: recovery.versions, refusedSeqs: recovery.refusedSeqs,
        refusedRanges: recovery.refusedRanges, refusedOutcomes: recovery.refusedOutcomes } : {}),
      sessionState: () => ({ version, dirty }),
      frame: (request, coversThroughSeq) => ({ ...frame(request.viewport, request.sheet), ...commandIdentity,
        sequence: coversThroughSeq, coversThroughSeq, requestGeneration: request.requestGeneration ?? request.generation,
        zoom: request.zoom, dpr: request.dpr }),
      onCommit() {
        version += 1;
        dirty = true;
        host.emit('changed', { version, dirty });
      },
    });
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
    async openRecovery(_, checkpoint, input = {}) {
      if (disposed) throw new Error('Workbook session is disposed');
      if (handle) throw new Error('Workbook session is already open');
      const recovery = decodeWorkbookRecovery(checkpoint.state);
      if (checkpoint.epoch !== recovery.epoch || checkpoint.version !== recovery.revision ||
        checkpoint.throughSeq !== recovery.throughSeq) throw new TypeError('Workbook recovery envelope mismatch');
      if (isCollaborativeWorkbookRecovery(recovery.engine)) throw new Error('Recovery is unavailable for collaborative workbooks');
      await (options.initWasm ?? initWasm)(input.wasm);
      if (disposed) throw new Error('Workbook session is disposed');
      const opened = openWorkbookRecoveryCheckpoint(recovery.engine, recovery.throughSeq);
      try {
        commandIdentity.session = recovery.session;
        commands = restoreWorkbookCommandState(opened, recovery, commandIdentity);
        version = recovery.hostVersion;
        dirty = recovery.dirty;
        handle = opened;
        installCommandHost(recovery);
        const info = opened.sheetInfo();
        return { format: 'xlsx' as const, stage: 'ready' as const, version, dirty,
          sheets: commands.sheetIds.map((id, index) => ({ id, index, name: info.sheetNames[index] })),
          activeSheet: info.activeSheet, commandContext: { ...commandIdentity,
            basis: { epoch: commands.epoch, version: opened.version() } },
          throughSeq: commands.throughSeq, revision: recovery.revision, historyReset: commands.historyReset === true, pngExport: isPngExportAvailable(),
          selection: commands.selection, editState: commands.editState };
      } catch (error) {
        handle = undefined;
        commands = undefined;
        commandHost = undefined;
        opened.dispose();
        throw error;
      }
    },
    async open(_, bytes, input = {}) {
      if (disposed) throw new Error('Workbook session is disposed');
      if (handle) throw new Error('Workbook session is already open');
      const initial = input.initialViewport;
      if (initial && !isWorkbookCommandFrameRequest({ sheet: initial.sheet ?? 0,
        viewport: { x: 0, y: 0, width: initial.width, height: initial.height },
        zoom: initial.zoom, dpr: initial.dpr, generation: initial.requestGeneration })) {
        throw new TypeError('Invalid workbook initial paint request');
      }
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
      collaborative = input.collaborative === true;
      try {
        if (input.initialUpdate) {
          if (!collaborative || retainPeerHydration) throw new Error('Initial updates require a collaborative worker replica');
          opened.applyUpdate(input.initialUpdate);
        }
        if (input.recoveryVersion !== undefined) adoptWorkbookPeerVersion(opened, input.recoveryVersion);
        if (initial?.sheet !== undefined) {
          checkSheet(opened, initial.sheet);
          opened.setActiveSheet(initial.sheet);
        }
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
        if (collaborative && !retainPeerHydration) offUpdates = opened.onUpdate((update, origin) => {
          host.emit('collaborationUpdate', { update, origin });
        });
        if (collaborative && !retainPeerHydration) {
          pipe = createCollaborationPipeHost<WorkerCollaborationProvider, AwarenessCursor, AwarenessPeer>({
            clientId: opened.clientId,
            read: { encodeStateVector: () => opened.encodeStateVector(),
              encodeStateAsUpdate: (vector) => opened.encodeStateAsUpdate(vector) },
            onUpdate: (listener) => opened.onUpdate(listener),
            merge(updates) {
              for (const update of updates) {
                const before = opened.version();
                try { opened.applyUpdate(update); }
                catch (error) {
                  if (!(error instanceof Error) || error instanceof WebAssembly.RuntimeError || opened.version() !== before) throw error;
                  return { failed: error };
                }
              }
              return { applied: null };
            },
            createProvider: (replica, transport, providerOptions) => {
              const provider = new CollaborationProvider(replica, transport, providerOptions);
              return Object.assign(provider, { onPeers: (listener: Parameters<CollaborationProvider['onAwareness']>[0]) =>
                provider.onAwareness(listener) });
            },
            arrive: providerArrival,
            leave: providerLeave,
            transportError: (error) => new CollaborationError(error.code, error.message),
            onOutput: schedulePipeOutput,
          });
        }
        commands = createWorkbookCommandState(commandIdentity, { epoch: 1, version: opened.version() },
          { sheetId: info.sheetIds[info.activeSheet], range: { kind: 'a1', a1: 'A1' } }, info.sheetIds, calculation ?? null);
        installCommandHost();
        if (retainPeerHydration) host.emit('peerOpened', { version: opened.version(), initialCalculation });
        const projection = { ...commands };
        if (initial || !retainPeerHydration && input.includeEditState) projectEditState(opened, projection);
        const state = { format: 'xlsx' as const, stage: 'ready' as const, version, dirty,
          ...(collaborative && !retainPeerHydration ? { clientId: opened.clientId } : {}),
          sheets: summaries, activeSheet: info.activeSheet, pngExport: isPngExportAvailable(),
          commandContext: { ...commandIdentity, basis: { epoch: commands.epoch, version: opened.version() } },
          ...(!retainPeerHydration && projection.editState ? { editState: projection.editState, selection: commands.selection } : {}) };
        if (!initial) return state;
        const view = sheetView(info.activeSheet);
        const paint = { ...frame({ x: view.initialScrollX, y: view.initialScrollY,
          width: initial.width / initial.zoom, height: initial.height / initial.zoom }, info.activeSheet),
          ...commandIdentity, coversThroughSeq: 0, requestGeneration: initial.requestGeneration,
          zoom: initial.zoom, dpr: initial.dpr };
        commandHost!.rememberFrame({ sheet: info.activeSheet, viewport: paint.viewport,
          zoom: initial.zoom, dpr: initial.dpr, generation: initial.requestGeneration });
        return transferable({ ...state, initialPaint: {
          sheetView: view, frame: paint, editState: projection.editState!,
        } }, [paint.displayList]);
      } catch (error) {
        handle = undefined;
        commands = undefined;
        commandHost = undefined;
        previewSource = undefined;
        pipe?.dispose();
        pipe = undefined;
        offUpdates();
        offUpdates = () => {};
        opened.dispose();
        throw error;
      }
    },
    version: () => workbook().version(),
    collaborationStateVector() {
      checkCollaboration();
      return workbook().encodeStateVector();
    },
    collaborationStateAsUpdate(_, remoteStateVector) {
      checkCollaboration();
      return workbook().encodeStateAsUpdate(remoteStateVector);
    },
    listDefinedNames: (_, request) => workbook().listDefinedNames(request),
    readCells: (_, request) => orderedRequest(request, (resolved) => workbook().readCells(resolved)),
    findText: (_, request) => orderedRequest(request, (resolved) => workbook().findText(resolved)),
    validateEdits: (_, request) => request.expectVersion !== workbook().version()
      ? workbook().validateEdits(request) : orderedRequest(request, (resolved) => workbook().validateEdits(resolved)),
    applyEdits(_, request) {
      if (commands?.throughSeq) throw new Error('Workbook command sessions require command batches');
      commandHost?.validateDirectEdits();
      const result = workbook().applyEdits(request);
      if (result.ok && result.applied) {
        commandHost?.synchronizeDirectEdits(result.version);
        version += 1;
        dirty = true;
        host.emit('changed', { version, dirty });
      }
      return result;
    },
    async command(_, command) {
      checkCommands();
      if (command.kind === 'applyUpdate' || command.kind === 'collabPipe') checkCollaboration();
      const before = workbook().version();
      const reply = await commandHost!.command(command);
      if (workbook().version() !== before) host.scheduler.bump({ version: true, generation: true });
      return reply;
    },
    edit(_, command) {
      checkCommands();
      if (command.kind === 'applyUpdate' || command.kind === 'collabPipe') checkCollaboration();
      return commandHost!.handler(null, command);
    },
    checkpoint(_, preserveHistory = false) {
      checkCommands();
      return commandHost!.checkpoint(null, preserveHistory);
    },
    commandFrame(_, request) {
      checkCommands();
      const rendered = request ? commandHost!.requestFrame(request) : commandHost!.frame();
      return deferReply(rendered.catch((error: unknown) => {
        if (error instanceof Error && !(error instanceof SessionFailure) &&
          error.message === 'No frame is available at the settled command boundary') return undefined;
        throw error;
      }));
    },
    frame(_, viewport, options = {}) {
      const paint = frame(viewport, options.sheet ?? workbook().sheetInfo().activeSheet);
      return transferable(paint, [paint.displayList]);
    },
    sheetView(_, sheet) {
      return sheetView(sheet);
    },
    cellGeometry(_, sheet, row, col) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return {
        sheet, version: opened.version(), rect: opened.cellRect(sheet, row, col),
        scrollPosition: opened.cellPosition(sheet, row, col),
      };
    },
    selectionData(_, sheet, row, col, direction) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return opened.selectionData(sheet, row, col, direction);
    },
    cellInputs(_, sheet, range) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return { sheet, version: opened.version(), cells: opened.rangeCells(sheet, range) };
    },
    captureFormat(_, sheet, range) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return opened.captureFormat(sheet, range);
    },
    sheets: () => sheets({ ...workbook().sheetInfo(),
      ...(commands?.throughSeq ? { sheetIds: commands.sheetIds } : {}) }),
    calculationStatus: () => workbook().calculationStatus(),
    renderPng(_, viewport, options = {}) {
      return renderPng(options.sheet ?? workbook().sheetInfo().activeSheet, (opened) => opened.renderPng(viewport));
    },
    renderRangePng(_, options, frameOptions = {}) {
      return renderPng(frameOptions.sheet ?? workbook().sheetInfo().activeSheet, (opened) => opened.renderRangePng(options));
    },
    exportStructured: (_, options) => exported(options),
    exportMarkdown(_, options, markdownOptions) {
      const opened = workbook();
      const result = opened.exportMarkdown(options, markdownOptions);
      if (!commands?.throughSeq) return result;
      if (result.ok) {
        const structured = exported(options);
        if (!structured.ok) return structured;
      }
      return translateWorkbookSheetIds(result, opened.sheetInfo().sheetIds, commands.sheetIds);
    },
    printDisplayList(_, sheet, range, metrics, gridlines) {
      const opened = workbook();
      checkSheet(opened, sheet);
      return transferBytes(encoder.encode(JSON.stringify(opened.printDisplayList(sheet, range, metrics, gridlines))));
    },
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

  function checkCommands(): void {
    workbook();
    if (!commandHost) throw new Error('Workbook commands are not open');
    if (retainedHydration || peerAttached || sequence > 0) {
      throw new Error('Workbook commands require a session without an edit peer');
    }
  }

  function checkCollaboration(): void {
    checkCommands();
    if (!collaborative) throw new Error('Workbook session is not collaborative');
  }

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

  const replyTransport: SessionTransport = { ...transport, post(message, transfer) {
    const output = isHostMessage(message) && message.kind === 'reply' ? pipe?.drain() : undefined;
    const collaboration: WorkbookCollaborationPipeReply | undefined = output ? { collabPipe: output.batches } : undefined;
    transport.post(collaboration ? { ...(message as object), ...collaboration } : message,
      output ? [...(transfer ?? []), ...output.transfer] : transfer);
  } };
  const host = createSessionHost<Methods, Events, null>(replyTransport, {
    handlers: { ...handlers, ...internalHandlers },
    policies: { ...WORKBOOK_SESSION_POLICIES, ...WORKBOOK_INTERNAL_SESSION_POLICIES,
      edit: (command) => commandHost?.policy(command) ?? { lane: 'input', reorderable: false },
      checkpoint: () => commandHost?.checkpointPolicy() ?? { lane: 'input', reorderable: false },
    },
    context: null, onDispose: dispose,
  });
  return host;
}

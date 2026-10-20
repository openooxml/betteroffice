import { StrictMode } from 'react';
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { XlsxEditor, defineXlsxPlugin, type XlsxAsyncEditorApi, type XlsxPluginContext, type XlsxPluginSelection, type XlsxWorkerEditorApi } from '@betteroffice/xlsx-react';
import { editableWorkbookSessionBackend } from '../../../packages/xlsx-react/src/worker/useEditableSessionWorkbook';
import { cellRect, openWorkbookSession, rangeRect, type WorkbookEditPeer, type WorkbookFrame, type WorkbookSession } from '@betteroffice/xlsx';
import {
  CollaborationProvider,
  createAsyncCollaborationReplica,
  type CollaborationTransportEvent,
  type CollaborationTransport,
} from '@betteroffice/xlsx/collaboration';
import { decodeMessages } from '../../../packages/xlsx/src/collaboration/protocol';
import { workbookSessionInternals, WORKBOOK_REPLAY_MUTATORS } from '../../../packages/xlsx/src/session/replay';
import type { WorkerEditorProbe } from './xlsx-worker-editor-probe';
import { addPatternSheet } from './xlsx-pattern-workbook';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

const parameters = new URLSearchParams(location.search);
const variant = parameters.get('variant') ?? 'plain';
const collaboration = parameters.has('collaboration');
const peerless = parameters.has('asyncApi') || collaboration;
let mainThreadEngines = 0;
globalThis.__xlsxMainEngineProbe = { capture() { mainThreadEngines += 1; } };

async function workbook() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId3"/></sheets>
</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`);
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="sans-serif"/></font><font><sz val="16"/><name val="Georgia"/><b/><i/><u/><strike/><color rgb="FF884422"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEECC"/></patternFill></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="2" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="bottom"/></xf></cellXfs>
</styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${variant === 'merged-grid' ? 'A1:D4' : 'A1:B2'}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="24"/>
  <cols><col min="1" max="2" width="24" customWidth="1"/></cols>
  <sheetData><row r="1"><c r="A1" s="${variant === 'formatted' ? 2 : variant === 'styled' ? 1 : 0}" t="inlineStr"><is><t>initial</t></is></c>${variant === 'merged' || variant === 'overflow' ? '' : '<c r="B1"><v>1</v></c>'}</row></sheetData>
${variant === 'merged' ? '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' : ''}
${variant === 'merged-grid' ? '<mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>' : ''}
<drawing r:id="rId1"/>
</worksheet>`);
  zip.file('xl/worksheets/sheet2.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>other</t></is></c></row></sheetData></worksheet>');
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  zip.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="1714500" cy="1143000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Route chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
  zip.file('xl/charts/chart1.xml', `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Route</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`);
  if (variant === 'patterns') addPatternSheet(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

const paints = new WeakMap<HTMLCanvasElement, Set<string>>();
const nativeClear = CanvasRenderingContext2D.prototype.clearRect;
CanvasRenderingContext2D.prototype.clearRect = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeClear>) {
  nativeClear.apply(this, args);
  if (this.canvas.closest('#worker-editor-root')) paints.set(this.canvas, new Set());
};
const nativeText = CanvasRenderingContext2D.prototype.fillText;
CanvasRenderingContext2D.prototype.fillText = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeText>) {
  nativeText.apply(this, args);
  paints.get(this.canvas)?.add(args[0]);
};

const rootElement = document.getElementById('worker-editor-root')!;
const root = createRoot(rootElement);
const commitOrder: WorkerEditorProbe['commitOrder'] = [];
let holdingHydration = parameters.has('holdHydration');
let hydrationWaiting = false;
const hydrationRelease = deferred<void>();
const hydrate = editableWorkbookSessionBackend.hydrate;
editableWorkbookSessionBackend.hydrate = async (...args) => {
  if (holdingHydration) {
    hydrationWaiting = true;
    if (parameters.get('holdHydration') === 'during') {
      const peer = await hydrate(...args);
      await hydrationRelease.promise;
      hydrationWaiting = false;
      return peer;
    }
    await hydrationRelease.promise;
    hydrationWaiting = false;
  }
  return hydrate(...args);
};
const peerEntries: WorkerEditorProbe['peerEntries'] = [];
const replayEntries: WorkerEditorProbe['replayEntries'] = [];
const peers = new Map<number, WorkbookEditPeer>();
const sessionGenerations = new WeakMap<WorkbookSession, number>();
const initialVersions = new Map<number, string>();
const workers: Worker[] = [];
const NativeWorker = globalThis.Worker;
globalThis.Worker = new Proxy(NativeWorker, {
  construct(target, args) {
    const worker = Reflect.construct(target, args) as Worker;
    workers.push(worker);
    return worker;
  },
});
type PeerlessSlot = {
  api: XlsxAsyncEditorApi | null;
  ready: ReturnType<typeof deferred<XlsxAsyncEditorApi>>;
  synced: ReturnType<typeof deferred<void>>;
  session: WorkbookSession | null;
  frame: WorkbookFrame | null;
  generation: number;
  context: XlsxPluginContext<null> | null;
  provider: CollaborationProvider | null;
};

function createSlot(): PeerlessSlot {
  return {
    api: null,
    ready: deferred<XlsxAsyncEditorApi>(),
    synced: deferred<void>(),
    session: null,
    frame: null,
    generation: 0,
    context: null,
    provider: null,
  };
}

const slots: PeerlessSlot[] = Array.from(
  { length: collaboration ? 2 : 1 },
  () => createSlot(),
);
const eventListeners = [
  new Set<(event: CollaborationTransportEvent) => void>(),
  new Set<(event: CollaborationTransportEvent) => void>(),
];
const connected = [false, false];
const applicationQueue: ReturnType<typeof deferred<void>>[][] = [[], []];
const applicationReceipts: Promise<void>[] = [];
let deliveryPaused = false;
const deliveryQueue: { sender: number; data: Uint8Array }[] = [];

function flushDeliveryQueue() {
  if (deliveryPaused) return;
  let index = 0;
  while (index < deliveryQueue.length) {
    const entry = deliveryQueue[index]!;
    const target = 1 - entry.sender;
    if (!connected[target]) {
      index += 1;
      continue;
    }
    deliveryQueue.splice(index, 1);
    for (const message of decodeMessages(entry.data)) {
      if (message.type !== 'update' && message.type !== 'sync-step-2') continue;
      const receipt = deferred<void>();
      applicationQueue[target]!.push(receipt);
      applicationReceipts.push(receipt.promise);
      void receipt.promise.catch(() => {});
    }
    for (const listener of eventListeners[target]!) {
      listener({ type: 'message', data: entry.data.slice() });
    }
  }
}

function createLoopbackTransport(index: number): CollaborationTransport {
  const listeners = eventListeners[index]!;
  return {
    connect() {
      connected[index] = true;
      for (const listener of listeners) listener({ type: 'open' });
      flushDeliveryQueue();
    },
    disconnect() {
      connected[index] = false;
      for (const listener of listeners) listener({ type: 'close' });
    },
    send(data) {
      deliveryQueue.push({ sender: index, data: data.slice() });
      flushDeliveryQueue();
      return true;
    },
    onEvent(callback) {
      listeners.add(callback);
      return () => { listeners.delete(callback); };
    },
  };
}

function editorRoot(index = 0): ParentNode {
  return (
    rootElement.querySelector<HTMLElement>(
      `[data-peerless-editor="${index}"]`,
    ) ?? rootElement
  );
}

let generation = 0;
let activeSession: WorkbookSession | null = null;
let latestFrame: WorkbookFrame | null = null;
let cellInputHold: ReturnType<typeof deferred<void>> | null = null;
let heldCellInputs = 0;
const backendOpen = editableWorkbookSessionBackend.open;
editableWorkbookSessionBackend.open = async (...args) => {
  const token = ++generation;
  const slotIndex =
    args[1]?.clientId !== undefined ? args[1].clientId - 4200 : 0;
  const slot = slots[slotIndex] ?? slots[0]!;
  const slotGeneration = ++slot.generation;
  const session = await backendOpen(...args);
  if (token === generation) activeSession = session;
  if (slotGeneration === slot.generation) slot.session = session;
  sessionGenerations.set(session, token);
  const internal = workbookSessionInternals.get(session)!;
  if (internal.initialVersion) initialVersions.set(token, internal.initialVersion);
  const replay = internal.replay;
  internal.replay = (envelope) => {
    replayEntries.push({ generation: token, sequence: envelope.sequence, method: envelope.op.method, args: structuredClone(envelope.op.args) });
    return replay(envelope);
  };
  const frame = session.call.frame;
  session.call.frame = async (...input) => {
    const value = await frame(...input);
    if (token === generation) latestFrame = value;
    if (slotGeneration === slot.generation) slot.frame = value;
    return value;
  };
  const cellInputs = session.call.cellInputs;
  session.call.cellInputs = async (...input) => {
    const hold = cellInputHold;
    const value = await cellInputs(...input);
    if (!hold) return value;
    heldCellInputs += 1;
    try { await hold.promise; } finally { heldCellInputs -= 1; }
    return value;
  };
  return session;
};
const attach = editableWorkbookSessionBackend.attach;
editableWorkbookSessionBackend.attach = (options) => {
  const token = sessionGenerations.get(options.session)!;
  for (const method of Object.keys(WORKBOOK_REPLAY_MUTATORS) as (keyof typeof WORKBOOK_REPLAY_MUTATORS)[]) {
    const native = options.peer[method] as (...args: unknown[]) => unknown;
    Object.defineProperty(options.peer, method, { value: (...args: unknown[]) => {
      const entry = { generation: token, method, args: structuredClone(args) };
      peerEntries.push(entry);
      if (method === 'editCell') commitOrder.push({ kind: 'mutator-entry', text: args[3] as string });
      try {
        const result = native(...args);
        if (result && typeof result === 'object' && 'ok' in result && result.ok === false) {
          peerEntries.splice(peerEntries.indexOf(entry), 1);
        }
        return result;
      } catch (error) {
        peerEntries.splice(peerEntries.indexOf(entry), 1);
        throw error;
      }
    } });
  }
  const edits = attach(options);
  peers.set(token, edits);
  return edits;
};
const errors: string[] = [];
const previews: string[] = [];
const previewFrames: string[] = [];
const observer = new MutationObserver(() => {
  const preview = rootElement.querySelector('[data-testid="xlsx-commit-preview"]');
  if (!preview?.textContent) return;
  const text = preview.textContent;
  previews.push(text);
  requestAnimationFrame(() => {
    if (preview.isConnected && preview.textContent === text) previewFrames.push(text);
  });
});
observer.observe(rootElement, { subtree: true, childList: true, characterData: true });
let current = deferred<XlsxWorkerEditorApi>();
let api: XlsxWorkerEditorApi | null = null;
let peerlessCurrent = deferred<XlsxAsyncEditorApi>();
let peerlessApi: XlsxAsyncEditorApi | null = null;
let savedBytes: Uint8Array | null = null;
let context: XlsxPluginContext<null> | null = null;
const plugin = defineXlsxPlugin({ id: 'routes', createState: () => null, initialize(value) { context = value; } });
let source: Uint8Array;
let readOnly = false;
let readyCount = 0;

function renderPeerlessEditor(bytes: Uint8Array, index: number) {
  const slot = slots[index]!;
  const slotPlugin = defineXlsxPlugin({
    id: 'routes',
    createState: () => null,
    initialize(value) {
      slot.context = value;
      if (index === 0) context = value;
    },
  });
  const transport = collaboration ? createLoopbackTransport(index) : null;
  return (
    <XlsxEditor
      file={bytes}
      readOnly={readOnly}
      plugins={[slotPlugin]}
      pluginGrants={{ routes: { document: 'write', editBatches: true } }}
      collaboration={
        collaboration
          ? {
              clientId: 4200 + index,
              onReplica(replica) {
                slot.provider?.destroy();
                slot.provider = null;
                if (!replica) return;
                const asyncReplica = slot.session!.replica!.async;
                const tracked = createAsyncCollaborationReplica({
                  ...asyncReplica,
                  async applyUpdate(update) {
                    const receipt = applicationQueue[index]!.shift();
                    if (!receipt) throw new Error('Remote application has no delivery receipt');
                    try {
                      const result = await asyncReplica.applyUpdate(update);
                      receipt.resolve();
                      return result;
                    } catch (error) {
                      receipt.reject(error);
                      throw error;
                    }
                  },
                });
                const provider = new CollaborationProvider(tracked, transport!);
                slot.provider = provider;
                provider.onError((error) => errors.push(error.message));
                const off = provider.onStatus(({ synced }) => {
                  if (synced) { off(); slot.synced.resolve(); }
                });
                provider.connect();
              },
            }
          : undefined
      }
      onError={(error) => errors.push(error.message)}
      onSave={(nextBytes) => {
        savedBytes = nextBytes;
      }}
      onReady={(value) => {
        readyCount += 1;
        slot.api = value;
        slot.ready.resolve(value);
        if (index === 0) {
          peerlessApi = value;
          peerlessCurrent.resolve(value);
        }
      }}
    />
  );
}

const open = (bytes: Uint8Array) => {
  source = bytes;
  const editor = collaboration ? (
    <div style={{ display: 'flex', width: '100%', height: '100%' }}>
      {slots.map((_, index) => (
        <div
          key={index}
          data-editor
          data-peerless-editor={index}
          style={{ width: '50%', height: '100%', minWidth: 0 }}
        >
          {renderPeerlessEditor(bytes, index)}
        </div>
      ))}
    </div>
  ) : peerless ? (
    renderPeerlessEditor(bytes, 0)
  ) : (
    <XlsxEditor legacySyncApi file={bytes} experimentalWorkerOpen readOnly={readOnly} plugins={[plugin]}
      pluginGrants={{ routes: { document: 'write', editBatches: true } }} onError={(error) => errors.push(error.message)}
      onReady={(value) => {
        readyCount += 1;
        if (!('whenHydrated' in value)) throw new Error('Expected worker editor session');
        api = value; current.resolve(value);
      }} />
  );
  root.render(parameters.has('strict') ? <StrictMode>{editor}</StrictMode> : editor);
};
const ready = (async () => {
  await document.fonts.ready;
  open(await workbook());
  if (peerless) {
    const editors = await Promise.all(slots.map((slot) => slot.ready.promise));
    if (!holdingHydration) {
      await Promise.all(editors.map((editor) => editor.whenHydrated()));
    }
    if (collaboration) {
      await Promise.all(slots.map((slot) => slot.synced.promise));
    }
    return;
  }
  const editor = await current.promise;
  if (!holdingHydration) await editor.whenHydrated();
})();

async function reopenPeerless(bytes: Uint8Array) {
  peerlessApi = null;
  peerlessCurrent = deferred<XlsxAsyncEditorApi>();
  context = null;
  slots[0]!.api = null;
  slots[0]!.ready = deferred<XlsxAsyncEditorApi>();
  slots[0]!.context = null;
  open(bytes);
  const editor = await peerlessCurrent.promise;
  await editor.whenHydrated();
  return bytes.byteLength;
}

window.__xlsxWorkerEditor = {
  ready, errors, previews, previewFrames, commitOrder, peerEntries, replayEntries,
  mainThreadEngines: () => mainThreadEngines,
  peerless: peerless ? {
    get editor() { return peerlessApi!; },
    editorAt(index = 0) { return slots[index]!.api!; },
    async cells(sheet, range, index = 0) {
      const session = slots[index]?.session ?? activeSession!;
      const result = await session.call.cellInputs(sheet, range);
      return result.cells.map((row) => row.map((cell) => cell.input));
    },
    async selection(index = 0) {
      const target = slots[index]?.context ?? context;
      if (!target) throw new Error('Plugin is not active');
      let selected: XlsxPluginSelection = null;
      await target.run((current) => { selected = structuredClone(current.snapshot.selection); });
      return selected;
    },
    failure: (index = 0) => (slots[index]?.api ?? peerlessApi)?.failure?.message ?? null,
    async cellRects(cells, index = 0) {
      const target = slots[index]?.context ?? context;
      if (!target) throw new Error('Plugin is not active');
      const scope = editorRoot(index);
      let placed: Awaited<ReturnType<NonNullable<WorkerEditorProbe['peerless']>['cellRects']>> = null;
      await target.run(({ geometry }) => {
        const canvas = scope.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]');
        const layer = scope.querySelector<HTMLElement>('[data-testid="plugin-overlays"]');
        if (!geometry || !canvas || !layer) return;
        const canvasBounds = canvas.getBoundingClientRect();
        const layerBounds = layer.getBoundingClientRect();
        const dx = layerBounds.left + layer.clientLeft - layer.scrollLeft - canvasBounds.left;
        const dy = layerBounds.top + layer.clientTop - layer.scrollTop - canvasBounds.top;
        const rects = cells.map(([row, col]) => geometry.getCellRect({ sheetId: geometry.layout.sheetId, row, col }));
        if (rects.some((rect) => !rect)) return;
        placed = { zoom: geometry.layout.zoom, rects: rects.map((rect) => ({ ...rect!, x: rect!.x + dx, y: rect!.y + dy })) };
      });
      return placed;
    },
    formulaBar(index = 0) {
      return (
        editorRoot(index).querySelector<HTMLInputElement>(
          '[data-testid="xlsx-formula-input"]',
        )?.value ?? ''
      );
    },
    workerCount: () => workers.length,
    async saveAndReopen() {
      savedBytes = null;
      const result = await peerlessApi!.commands.execute('save', null);
      if (!result.ok || !savedBytes) throw new Error('Workbook save returned no bytes');
      return reopenPeerless(savedBytes);
    },
    async recoverAndReopen() {
      const result = await peerlessApi!.recoverySave();
      if (!result.recovery) throw new Error('Expected recovery bytes');
      return reopenPeerless(result.bytes);
    },
  } : null,
  generation: (index = 0) => (collaboration ? slots[index]!.generation : generation),
  readyCount: () => readyCount,
  setReadOnly(value) { readOnly = value; flushSync(() => open(source)); },
  readOnlyEditRefusal() {
    try { api!.editCell(0, 0, 0, 'blocked'); return ''; }
    catch (error) { return (error as Error).message; }
  },
  async savedInputs() {
    const bytes = await api!.saveAsync();
    if (!bytes) throw new Error('Workbook save returned no bytes');
    const session = await openWorkbookSession(bytes);
    try {
      const result = await session.call.cellInputs(0, 'A1:B2');
      return [result.cells[0][0].input, result.cells[1][1].input];
    } finally { await session.dispose(); }
  },
  releaseHydration() { holdingHydration = false; hydrationRelease.resolve(); },
  holdCellInputs() { cellInputHold = deferred<void>(); },
  heldCellInputs: () => heldCellInputs,
  releaseCellInputs() { cellInputHold?.resolve(); cellInputHold = null; },
  async flush() { await api!.flush(); },
  async zoom(scale) { await (peerless ? peerlessApi : api)!.commands.execute('zoom', { scale }); },
  cellClip() {
    const grid = latestFrame?.displayList.grid;
    const canvas = rootElement.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]');
    if (!grid || !canvas) throw new Error('Worker cell geometry unavailable');
    const rect = variant === 'merged' ? rangeRect(grid, { top: 0, left: 0, bottom: 0, right: 1 }) : cellRect(grid, 0, 0);
    if (!rect) throw new Error('Edited cell is not visible');
    const bounds = canvas.getBoundingClientRect();
    const zoom = canvas.clientWidth / latestFrame!.viewport.width;
    return { x: bounds.x + rect.x * zoom, y: bounds.y + rect.y * zoom, width: rect.w * zoom, height: rect.h * zoom };
  },
  mergedGridPoints() {
    const grid = latestFrame?.displayList.grid;
    const canvas = rootElement.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]');
    if (!grid || !canvas) throw new Error('Worker grid geometry unavailable');
    const merge = rangeRect(grid, { top: 1, left: 1, bottom: 2, right: 2 });
    const first = cellRect(grid, 1, 1);
    const last = cellRect(grid, 2, 2);
    const below = cellRect(grid, 3, 2);
    if (!merge || !first || !last || !below) throw new Error('Merged grid is not visible');
    const zoom = canvas.clientWidth / latestFrame!.viewport.width;
    const point = (x: number, y: number) => ({ x: x * zoom, y: y * zoom });
    return {
      background: point(first.x + first.w / 2, first.y + first.h / 2),
      interiors: [point(last.x, first.y + first.h / 2), point(first.x + first.w / 2, last.y)],
      edges: [point(merge.x, first.y + first.h / 2), point(merge.x + merge.w, first.y + first.h / 2),
        point(first.x + first.w / 2, merge.y), point(first.x + first.w / 2, merge.y + merge.h)],
      neighbour: point(last.x, below.y + below.h / 2),
    };
  },
  chartClip() {
    const chart = latestFrame?.displayList.charts?.[0];
    const canvas = rootElement.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]');
    if (!chart || !canvas) throw new Error('Worker chart geometry unavailable');
    const bounds = canvas.getBoundingClientRect();
    const zoom = canvas.clientWidth / latestFrame!.viewport.width;
    return { x: bounds.x + chart.rect.x * zoom, y: bounds.y + chart.rect.y * zoom,
      width: chart.rect.w * zoom, height: chart.rect.h * zoom };
  },
  queueNavigation() { void api!.selectCellsAsync(1, { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 } }).catch(() => {}); },
  adoptedSequence() {
    return Number(rootElement.querySelector<HTMLCanvasElement>('[data-paint-source="worker"]')?.dataset.workerSequence ?? -1);
  },
  async hostEdit(value) { await api!.editCellAsync(0, 1, 1, value); },
  queueHostEdit(value) { void api!.editCellAsync(0, 1, 1, value).catch(() => {}); },
  queueCellHostEdit(value) { void api!.editCellAsync(0, 0, 0, value).catch(() => {}); },
  queueStyledBatch() {
    const version = initialVersions.get(generation);
    if (!version) throw new Error('Missing edit version');
    const target = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };
    void api!.applyEdits({ expectVersion: version, steps: [
      { op: 'setCellInputs', target, inputs: [['styled batch']] },
      { op: 'patchStyle', target, patch: { fontSize: 24 } },
    ] }).catch((error) => errors.push(error.message));
  },
  queueBulkFill() {
    const version = peers.get(generation) ? undefined : initialVersions.get(generation);
    if (!version) throw new Error('Queue a bulk fill before hydration');
    void api!.applyEdits({ expectVersion: version, steps: [{ op: 'setCellInputs',
      target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A3:B3' } }, inputs: [['queued fill one', 'queued fill two']] }] }).catch(() => {});
  },
  async bulkFill() {
    const version = await api!.version();
    if (!version) throw new Error('Missing edit version');
    const result = await api!.applyEdits({ expectVersion: version, steps: [{ op: 'setCellInputs',
      target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A3:B3' } }, inputs: [['fill one', 'fill two']] }] });
    if (!result?.ok) throw new Error('Bulk fill refused');
  },
  async formatCells() {
    for (const [id, args] of [['fillColor', { color: '#ffeecc' }], ['textColor', { color: '#442211' }], ['numberFormat', { value: 'number' }]] as const) {
      const result = await api!.commands.execute(id, args as never);
      if (!result.ok) throw new Error(`${id} refused`);
    }
    const copied = await api!.commands.execute('paintFormat', null);
    if (!copied.ok) throw new Error('Paint format refused');
    const scroll = rootElement.querySelector<HTMLElement>('[data-testid="xlsx-scroll"]')!;
    scroll.focus();
    scroll.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    await api!.flush();
  },
  async pluginRoutes() {
    if (!context) throw new Error('Plugin is not active');
    await context.run(async (current) => {
      const version = await current.read.version();
      if (!version.ok) throw new Error('Plugin version refused');
      const batch = await current.edits!.applyEdits({ expectVersion: version.version, steps: [{ op: 'setCellInputs',
        target: { sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A4' } }, inputs: [['plugin']] }] });
      if (!batch.ok) throw new Error('Plugin batch refused');
      const next = await current.read.version();
      if (!next.ok) throw new Error('Plugin navigation version refused');
      const navigated = await current.navigation.selectCells({ sheetId: 'sheet:0', selection: {
        anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 },
      } }, { expectVersion: next.version, focus: true });
      if (!navigated.ok) throw new Error('Plugin navigation refused');
    });
  },
  async proposals() {
    const edits = peers.get(generation)!;
    const proposal = edits.propose('browser', null, [{ sheet: 0, row: 4, col: 0, input: 'proposal' }]);
    api!.refreshProposals();
    const accepted = await api!.commands.execute('proposalAccept', { proposalId: proposal.id });
    if (!accepted.ok) throw new Error('Proposal refused');
  },
  async history() {
    for (const id of ['undo', 'redo'] as const) {
      const result = await api!.commands.execute(id, null);
      if (!result.ok) throw new Error(`${id} refused`);
    }
  },
  fail() {
    const worker = workers.at(-1)!;
    worker.dispatchEvent(new ErrorEvent('error', { message: 'Browser worker failure', error: new Error('Browser worker failure') }));
    worker.terminate();
  },
  async recover() { return (await api!.recoverySave()).bytes.byteLength; },
  async replace() {
    api = null; current = deferred<XlsxWorkerEditorApi>(); context = null;
    open(await workbook());
    await current.promise;
  },
  dispose() { root.unmount(); },
  hydrated(index = 0) {
    return (peerless ? slots[index]?.api ?? peerlessApi : api)?.hydrated ?? false;
  },
  holdPreview() {
    if (!holdingHydration || !hydrationWaiting || api?.hydrated !== false) {
      throw new Error('Hydration must be held before opening the editor');
    }
    commitOrder.length = 0;
  },
  previewHeld() {
    return holdingHydration && hydrationWaiting && api?.hydrated === false &&
      rootElement.querySelector('[data-testid="xlsx-commit-preview"]')?.getAttribute('data-preview-ready') === 'true';
  },
  releasePreview(text) {
    if (!holdingHydration || !hydrationWaiting || api?.hydrated !== false) {
      throw new Error('Hydration preceded the preview screenshot');
    }
    if (commitOrder.length > 0) throw new Error('Mutation preceded the preview screenshot');
    if (rootElement.querySelector('[data-testid="xlsx-commit-preview"]')?.textContent !== text) {
      throw new Error('Pending text changed before the preview screenshot');
    }
    commitOrder.push({ kind: 'painted-preview', text });
    holdingHydration = false;
    hydrationRelease.resolve();
  },
  paintedTexts(index = 0) {
    const canvas = editorRoot(index).querySelector<HTMLCanvasElement>('[data-testid="xlsx-scroll"] canvas');
    return canvas ? [...(paints.get(canvas) ?? [])] : [];
  },
  formulaBar(index = 0) {
    return (
      editorRoot(index).querySelector<HTMLInputElement>(
        '[data-testid="xlsx-formula-input"]',
      )?.value ?? ''
    );
  },
  async cell(index = 0) {
    if (peerless) {
      const session = slots[index]?.session ?? activeSession;
      if (!session) return null;
      const result = await session.call.cellInputs(0, 'A1');
      return result.cells[0]?.[0]?.input ?? null;
    }
    return (await api?.cellAsync(0, 0, 0))?.input ?? null;
  },
  async saveAndReopen() {
    const bytes = await api!.save();
    if (!bytes) throw new Error('Workbook save returned no bytes');
    api = null;
    current = deferred<XlsxWorkerEditorApi>();
    open(bytes);
    const editor = await current.promise;
    await editor.whenHydrated();
    return bytes.byteLength;
  },
  async undo(index = 0) {
    const target = peerless ? slots[index]?.api ?? peerlessApi! : api!;
    const result = await target.commands.execute('undo', null);
    return result.ok;
  },
  pauseDelivery() {
    deliveryPaused = true;
  },
  async resumeDelivery() {
    deliveryPaused = false;
    flushDeliveryQueue();
    await Promise.all(applicationReceipts);
  },
  pendingUpdates(sender) {
    return deliveryQueue.filter((entry) => entry.sender === sender)
      .flatMap((entry) => decodeMessages(entry.data))
      .filter((message) => message.type === 'update').length;
  },
};

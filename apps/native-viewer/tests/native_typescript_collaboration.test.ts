import { expect, test } from 'bun:test';
import { createHash, randomUUID } from 'node:crypto';
import { readFile, rm } from 'node:fs/promises';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';

import { createRoomTransport } from '../../demo/app/collab/createRoomTransport';
import { CollaborationProvider } from '../../../packages/docx/src/collaboration/provider';
import type { PresentationHandle } from '../../../packages/pptx/src';
import type { YrsSession } from '../../../packages/docx/src/yrs';
import type { WorkbookHandle } from '../../../packages/xlsx/src';

const ROOT = resolve(import.meta.dir, '../../..');
const NATIVE_VIEWER = resolve(ROOT, 'apps/native-viewer');
const RELAY = resolve(ROOT, 'apps/relay');
const DOCUMENT = resolve(ROOT, 'apps/demo/public/betteroffice-demo.docx');
const SEED = resolve(ROOT, 'apps/demo/public/seeds/docx.bin');
const WASM = resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const XLSX_DOCUMENT = resolve(ROOT, 'apps/demo/public/showcase.xlsx');
const XLSX_SEED = resolve(ROOT, 'apps/demo/public/seeds/xlsx.bin');
const XLSX_WASM = resolve(ROOT, 'packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm');
const PPTX_DOCUMENT = resolve(ROOT, 'apps/demo/public/betteroffice-demo.pptx');
const PPTX_SEED = resolve(ROOT, 'apps/demo/public/seeds/pptx.bin');
const PPTX_WASM = resolve(ROOT, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm');
const WRANGLER_HOME = resolve(tmpdir(), `betteroffice-wrangler-${process.pid}`);

interface StoryState {
  storyId: string;
  checksum: string;
}

interface ParagraphState {
  paraId: string;
  text: string;
}

interface NativeSnapshot {
  format: 'docx' | 'xlsx' | 'pptx';
  connected: boolean;
  synced: boolean;
  stateVector: string;
  canonicalChecksum: string;
  stories: StoryState[];
  paragraphs: ParagraphState[];
  cells: XlsxCellState[];
  pptxStories: PptxStoryState[];
}

interface XlsxCellState {
  sheet: number;
  row: number;
  col: number;
  a1: string;
  input: string;
}

interface PptxStoryState {
  slide: number;
  shapeId: string;
  storyId: string;
  length: number;
  text: string;
}

interface NativeResponse {
  id: number;
  ok: boolean;
  snapshot?: NativeSnapshot;
  error?: string;
}

interface PendingRequest {
  resolve(snapshot: NativeSnapshot): void;
  reject(error: Error): void;
  timer: ReturnType<typeof setTimeout>;
}

class NativePeer {
  private readonly process;
  private readonly pending = new Map<number, PendingRequest>();
  private readonly stderrChunks: string[] = [];
  private nextId = 1;
  private stopped = false;

  constructor(origin: string, room: string, document = DOCUMENT) {
    this.process = Bun.spawn({
      cmd: [
        'cargo',
        'run',
        '--quiet',
        '--manifest-path',
        resolve(NATIVE_VIEWER, 'Cargo.toml'),
        '--',
        '--document',
        document,
        '--room',
        room,
        '--relay-origin',
        origin,
        '--collaboration-test-peer',
      ],
      cwd: ROOT,
      env: process.env,
      stdin: 'pipe',
      stdout: 'pipe',
      stderr: 'pipe',
    });
    void this.readResponses();
    void this.readStderr();
    void this.process.exited.then((code) => {
      this.stopped = true;
      if (code === 0) return;
      const error = new Error(
        `native peer exited with ${code}: ${this.stderrChunks.join('').slice(-4000)}`,
      );
      for (const request of this.pending.values()) {
        clearTimeout(request.timer);
        request.reject(error);
      }
      this.pending.clear();
    });
  }

  async command(
    command: 'snapshot' | 'disconnect' | 'reconnect' | 'shutdown',
  ): Promise<NativeSnapshot>;
  async command(
    command: 'insert',
    fields: { paraId: string; offset: number; text: string },
  ): Promise<NativeSnapshot>;
  async command(
    command: 'setCell',
    fields: { sheet: number; row: number; col: number; input: string },
  ): Promise<NativeSnapshot>;
  async command(
    command: 'insertPptx',
    fields: { storyId: string; index: number; text: string },
  ): Promise<NativeSnapshot>;
  async command(
    command: string,
    fields: Record<string, unknown> = {},
  ): Promise<NativeSnapshot> {
    if (this.stopped) throw new Error('native peer is stopped');
    const id = this.nextId++;
    const response = new Promise<NativeSnapshot>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(
          new Error(
            `native peer timed out on ${command}: ${this.stderrChunks.join('').slice(-4000)}`,
          ),
        );
      }, 240_000);
      this.pending.set(id, { resolve, reject, timer });
    });
    this.process.stdin.write(`${JSON.stringify({ id, command, ...fields })}\n`);
    await this.process.stdin.flush();
    return response;
  }

  async stop(): Promise<void> {
    if (this.stopped) return;
    try {
      await this.command('shutdown');
      await Promise.race([this.process.exited, Bun.sleep(5_000)]);
    } finally {
      if (!this.stopped) this.process.kill('SIGKILL');
    }
  }

  private async readResponses(): Promise<void> {
    const reader = this.process.stdout.getReader();
    const decoder = new TextDecoder();
    let buffered = '';
    while (true) {
      const { done, value } = await reader.read();
      buffered += decoder.decode(value, { stream: !done });
      const lines = buffered.split('\n');
      buffered = lines.pop() ?? '';
      for (const line of lines) {
        if (!line) continue;
        let response: NativeResponse;
        try {
          response = JSON.parse(line) as NativeResponse;
        } catch {
          this.stderrChunks.push(`unexpected native stdout: ${line}\n`);
          continue;
        }
        if (response.id === 0 && !response.ok) {
          const error = new Error(response.error ?? 'native peer rejected a command');
          for (const pending of this.pending.values()) {
            clearTimeout(pending.timer);
            pending.reject(error);
          }
          this.pending.clear();
          continue;
        }
        const request = this.pending.get(response.id);
        if (!request) continue;
        this.pending.delete(response.id);
        clearTimeout(request.timer);
        if (response.ok && response.snapshot) {
          request.resolve(response.snapshot);
        } else {
          request.reject(new Error(response.error ?? 'native peer command failed'));
        }
      }
      if (done) return;
    }
  }

  private async readStderr(): Promise<void> {
    const reader = this.process.stderr.getReader();
    const decoder = new TextDecoder();
    while (true) {
      const { done, value } = await reader.read();
      if (value) this.stderrChunks.push(decoder.decode(value, { stream: !done }));
      if (done) return;
    }
  }
}

function streamText(stream: ReadableStream<Uint8Array>): Promise<string> {
  return new Response(stream).text();
}

async function runChecked(command: string[], cwd: string): Promise<string> {
  const process = Bun.spawn({
    cmd: command,
    cwd,
    env: processEnv(),
    stdout: 'pipe',
    stderr: 'pipe',
  });
  const stdout = streamText(process.stdout);
  const stderr = streamText(process.stderr);
  const code = await process.exited;
  const [output, error] = await Promise.all([stdout, stderr]);
  if (code !== 0) {
    throw new Error(`${command.join(' ')} exited with ${code}\n${output}\n${error}`);
  }
  return output;
}

function processEnv(): Record<string, string | undefined> {
  return {
    ...process.env,
    CI: '1',
    NO_UPDATE_NOTIFIER: '1',
    WRANGLER_LOG_PATH: resolve(WRANGLER_HOME, 'logs'),
    WRANGLER_SEND_METRICS: 'false',
    XDG_CACHE_HOME: resolve(WRANGLER_HOME, 'cache'),
    XDG_CONFIG_HOME: resolve(WRANGLER_HOME, 'config'),
  };
}

async function availablePort(): Promise<number> {
  return new Promise((resolvePort, reject) => {
    const server = createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      if (!address || typeof address === 'string') {
        server.close();
        reject(new Error('failed to allocate relay port'));
        return;
      }
      server.close((error) => {
        if (error) reject(error);
        else resolvePort(address.port);
      });
    });
  });
}

async function startRelay(port: number) {
  const relay = Bun.spawn({
    cmd: [
      process.execPath,
      'run',
      'dev',
      '--',
      '--port',
      String(port),
      '--compatibility-date',
      '2026-07-15',
    ],
    cwd: RELAY,
    env: processEnv(),
    stdout: 'pipe',
    stderr: 'pipe',
  });
  const stdout = streamText(relay.stdout);
  const stderr = streamText(relay.stderr);
  let exitCode: number | null = null;
  void relay.exited.then((code) => {
    exitCode = code;
  });
  const origin = `http://127.0.0.1:${port}`;
  const deadline = Date.now() + 30_000;
  while (Date.now() < deadline) {
    if (exitCode !== null) {
      throw new Error(
        `relay exited with ${exitCode}\n${await stdout}\n${await stderr}`,
      );
    }
    try {
      const response = await fetch(origin);
      if (response.ok && (await response.text()) === 'ok') {
        return { origin, relay, stderr, stdout };
      }
    } catch {}
    await Bun.sleep(50);
  }
  relay.kill('SIGKILL');
  throw new Error(`relay did not start\n${await stdout}\n${await stderr}`);
}

async function stopRelay(relay: ReturnType<typeof Bun.spawn>): Promise<void> {
  relay.kill('SIGTERM');
  const stopped = await Promise.race([
    relay.exited.then(() => true),
    Bun.sleep(5_000).then(() => false),
  ]);
  if (!stopped) relay.kill('SIGKILL');
}

function hex(bytes: Uint8Array): string {
  return [...bytes].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

function canonicalStories(session: YrsSession): StoryState[] {
  return session.storyIds().map((storyId) => ({
    storyId,
    checksum: session.storyChecksum(storyId).toString(),
  }));
}

function canonicalChecksum(stories: StoryState[]): string {
  const checksum = createHash('sha256');
  checksum.update('canonical-document-checksums-v1\n');
  for (const story of stories) {
    checksum.update(story.storyId);
    checksum.update(Uint8Array.of(0));
    checksum.update(story.checksum);
    checksum.update('\n');
  }
  return checksum.digest('hex');
}

function typescriptSnapshot(session: YrsSession) {
  const stories = canonicalStories(session);
  return {
    stateVector: hex(session.encodeStateVector()),
    canonicalChecksum: canonicalChecksum(stories),
    stories,
  };
}

function typescriptText(session: YrsSession): string {
  return session
    .paragraphs('body')
    .map((paragraph) => paragraph.text)
    .join('\n');
}

function nativeText(snapshot: NativeSnapshot): string {
  return snapshot.paragraphs.map((paragraph) => paragraph.text).join('\n');
}

function nativeCell(
  snapshot: NativeSnapshot,
  sheet: number,
  row: number,
  col: number,
): string | undefined {
  return snapshot.cells.find((cell) => cell.sheet === sheet && cell.row === row && cell.col === col)
    ?.input;
}

function xlsxCanonicalChecksum(workbook: WorkbookHandle): string {
  const checksum = createHash('sha256');
  checksum.update('betteroffice-native-xlsx-workbook-v1\0');
  checksum.update(workbook.save());
  return checksum.digest('hex');
}

function pptxStoryText(presentation: PresentationHandle, storyId: string): string {
  return presentation
    .story(storyId)
    .paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

function pptxCanonicalChecksum(presentation: PresentationHandle): string {
  const checksum = createHash('sha256');
  checksum.update('betteroffice-native-pptx-deck-v1\0');
  checksum.update(presentation.save());
  return checksum.digest('hex');
}

async function waitUntil(
  assertion: () => boolean | Promise<boolean>,
  label: string,
  timeout = 20_000,
): Promise<void> {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    if (await assertion()) return;
    await Bun.sleep(50);
  }
  throw new Error(`timed out waiting for ${label}`);
}

async function waitForNative(
  peer: NativePeer,
  assertion: (snapshot: NativeSnapshot) => boolean,
  label: string,
): Promise<NativeSnapshot> {
  let snapshot = await peer.command('snapshot');
  await waitUntil(async () => {
    snapshot = await peer.command('snapshot');
    return assertion(snapshot);
  }, label);
  return snapshot;
}

async function assertConverged(
  peer: NativePeer,
  session: YrsSession,
  label: string,
): Promise<NativeSnapshot> {
  let native = await peer.command('snapshot');
  await waitUntil(async () => {
    native = await peer.command('snapshot');
    const typescript = typescriptSnapshot(session);
    return (
      native.canonicalChecksum === typescript.canonicalChecksum &&
      native.stateVector === typescript.stateVector
    );
  }, `${label} convergence`);
  const typescript = typescriptSnapshot(session);
  expect(native.stories).toEqual(typescript.stories);
  expect(native.canonicalChecksum).toBe(typescript.canonicalChecksum);
  expect(native.stateVector).toBe(typescript.stateVector);
  console.log(
    `${label}: checksum ${native.canonicalChecksum}, ${native.stories.length} stories, identical state vector`,
  );
  return native;
}

async function assertXlsxConverged(
  peer: NativePeer,
  workbook: WorkbookHandle,
  label: string,
): Promise<NativeSnapshot> {
  let native = await peer.command('snapshot');
  await waitUntil(async () => {
    native = await peer.command('snapshot');
    return (
      native.canonicalChecksum === xlsxCanonicalChecksum(workbook) &&
      native.stateVector === hex(workbook.encodeStateVector())
    );
  }, `${label} convergence`);
  expect(native.canonicalChecksum).toBe(xlsxCanonicalChecksum(workbook));
  expect(native.stateVector).toBe(hex(workbook.encodeStateVector()));
  console.log(
    `${label}: checksum ${native.canonicalChecksum}, ${native.cells.length} populated cells, identical state vector`,
  );
  return native;
}

async function assertPptxConverged(
  peer: NativePeer,
  presentation: PresentationHandle,
  label: string,
): Promise<NativeSnapshot> {
  let native = await peer.command('snapshot');
  await waitUntil(async () => {
    native = await peer.command('snapshot');
    return (
      native.canonicalChecksum === pptxCanonicalChecksum(presentation) &&
      native.stateVector === hex(presentation.encodeStateVector())
    );
  }, `${label} convergence`);
  expect(native.canonicalChecksum).toBe(pptxCanonicalChecksum(presentation));
  expect(native.stateVector).toBe(hex(presentation.encodeStateVector()));
  console.log(
    `${label}: checksum ${native.canonicalChecksum}, ${native.pptxStories.length} text stories, identical state vector`,
  );
  return native;
}

test(
  'native and the production TypeScript provider converge through the relay',
  async () => {
    await runChecked([process.execPath, resolve(ROOT, 'scripts/build-docx-wasm.ts')], ROOT);
    const [{ createYrsSession }, { preloadEditWasm }] = await Promise.all([
      import('../../../packages/docx/src/yrs/index'),
      import('../../../packages/docx/src/wasm/edit'),
    ]);
    await preloadEditWasm(new Uint8Array(await readFile(WASM)));

    const port = await availablePort();
    const relay = await startRelay(port);
    const room = `native-typescript-${randomUUID()}`;
    let native: NativePeer | null = null;
    let session: YrsSession | null = null;
    let provider: CollaborationProvider | null = null;
    let transport: ReturnType<typeof createRoomTransport> | null = null;
    const providerErrors: string[] = [];

    try {
      native = new NativePeer(relay.origin, room);
      let nativeState = await waitForNative(
        native,
        (snapshot) => snapshot.connected,
        'native relay connection',
      );
      const firstParagraph = nativeState.paragraphs[0];
      const nativeBeforeJoin = ' [native before join]';
      nativeState = await native.command('insert', {
        paraId: firstParagraph.paraId,
        offset: firstParagraph.text.length,
        text: nativeBeforeJoin,
      });
      expect(nativeText(nativeState)).toContain(nativeBeforeJoin);
      nativeState = await native.command('disconnect');
      expect(nativeState.connected).toBe(false);
      await Bun.sleep(100);

      session = await createYrsSession({ clientId: 2_147_400_001 });
      session.openDocx(new Uint8Array(await readFile(DOCUMENT)), false);
      session.loadState(new Uint8Array(await readFile(SEED)));
      transport = createRoomTransport(relay.origin, room);
      provider = new CollaborationProvider(session, transport);
      provider.onError((error) => providerErrors.push(error.message));
      provider.connect();

      await waitUntil(
        () => typescriptText(session!).includes(nativeBeforeJoin),
        'joining TypeScript peer to receive populated room state',
      );
      expect(provider.synced).toBe(false);
      expect(providerErrors).toEqual([]);
      console.log('join replay: TypeScript model received native pre-join text');

      await native.command('reconnect');
      nativeState = await waitForNative(
        native,
        (snapshot) => snapshot.connected && snapshot.synced,
        'native handshake after retained-state replay',
      );
      provider.disconnect();
      provider.connect();
      await waitUntil(() => provider?.synced === true, 'TypeScript handshake with native peer');

      const nativeLive = ' [native live]';
      const typescriptLive = ' [typescript live]';
      const latestNativeFirst = nativeState.paragraphs[0];
      const typescriptLast = session.paragraphs('body').at(-1);
      if (!typescriptLast) throw new Error('TypeScript replica has no body paragraphs');
      const nativeEdit = native.command('insert', {
        paraId: latestNativeFirst.paraId,
        offset: latestNativeFirst.text.length,
        text: nativeLive,
      });
      session.insertText(
        {
          story: 'body',
          paraId: typescriptLast.paraId,
          offset: typescriptLast.text.length,
        },
        typescriptLive,
      );
      await nativeEdit;
      await waitUntil(
        () => typescriptText(session!).includes(nativeLive),
        'TypeScript model to receive native live text',
      );
      nativeState = await waitForNative(
        native,
        (snapshot) => nativeText(snapshot).includes(typescriptLive),
        'native model to receive TypeScript live text',
      );
      expect(nativeText(nativeState)).toContain(nativeBeforeJoin);
      expect(typescriptText(session)).toContain(typescriptLive);
      await assertConverged(native, session, 'live edits');

      nativeState = await native.command('disconnect');
      expect(nativeState.connected).toBe(false);
      const nativeOffline = ' [native offline]';
      const typescriptOffline = ' [typescript while native offline]';
      const nativeOfflineTarget = nativeState.paragraphs[0];
      nativeState = await native.command('insert', {
        paraId: nativeOfflineTarget.paraId,
        offset: nativeOfflineTarget.text.length,
        text: nativeOffline,
      });
      expect(nativeText(nativeState)).toContain(nativeOffline);
      expect(typescriptText(session)).not.toContain(nativeOffline);

      const typescriptOfflineTarget = session.paragraphs('body')[0];
      session.insertText(
        {
          story: 'body',
          paraId: typescriptOfflineTarget.paraId,
          offset: typescriptOfflineTarget.text.length,
        },
        typescriptOffline,
      );
      await Bun.sleep(100);
      nativeState = await native.command('snapshot');
      expect(nativeText(nativeState)).not.toContain(typescriptOffline);

      await native.command('reconnect');
      nativeState = await waitForNative(
        native,
        (snapshot) =>
          snapshot.connected &&
          snapshot.synced &&
          nativeText(snapshot).includes(typescriptOffline),
        'native reconnect and offline TypeScript edit',
      );
      await waitUntil(
        () => typescriptText(session!).includes(nativeOffline),
        'TypeScript model to receive reconnected native edit',
      );
      expect(nativeText(nativeState)).toContain(nativeOffline);
      expect(typescriptText(session)).toContain(typescriptOffline);
      await assertConverged(native, session, 'reconnected edits');
      expect(providerErrors).toEqual([]);
    } finally {
      provider?.destroy();
      transport?.disconnect();
      session?.destroy();
      await native?.stop();
      await stopRelay(relay.relay);
      await rm(WRANGLER_HOME, { recursive: true, force: true });
    }
  },
  600_000,
);

test('native XLSX and the production TypeScript provider converge through the relay', async () => {
  await runChecked([process.execPath, resolve(ROOT, 'scripts/build-xlsx-wasm.ts')], ROOT);
  const [{ initWasm, openWorkbook }, { CollaborationProvider: XlsxCollaborationProvider }] =
    await Promise.all([
      import('../../../packages/xlsx/src/index'),
      import('../../../packages/xlsx/src/collaboration/provider'),
    ]);
  await initWasm(new Uint8Array(await readFile(XLSX_WASM)));

  const port = await availablePort();
  const relay = await startRelay(port);
  const room = `native-typescript-xlsx-${randomUUID()}`;
  let native: NativePeer | null = null;
  let workbook: WorkbookHandle | null = null;
  let provider: InstanceType<typeof XlsxCollaborationProvider> | null = null;
  let transport: ReturnType<typeof createRoomTransport> | null = null;
  const providerErrors: string[] = [];

  try {
    native = new NativePeer(relay.origin, room, XLSX_DOCUMENT);
    let nativeState = await waitForNative(
      native,
      (snapshot) => snapshot.connected,
      'native XLSX relay connection'
    );
    expect(nativeState.format).toBe('xlsx');
    const nativeBeforeJoin = 'native before join';
    nativeState = await native.command('setCell', {
      sheet: 0,
      row: 4,
      col: 1,
      input: nativeBeforeJoin,
    });
    expect(nativeCell(nativeState, 0, 4, 1)).toBe(nativeBeforeJoin);
    nativeState = await native.command('disconnect');
    expect(nativeState.connected).toBe(false);
    await Bun.sleep(100);

    workbook = openWorkbook(new Uint8Array(await readFile(XLSX_DOCUMENT)), {
      collaborative: true,
      clientId: 2_147_400_002,
    });
    workbook.applyUpdate(new Uint8Array(await readFile(XLSX_SEED)));
    transport = createRoomTransport(relay.origin, room);
    provider = new XlsxCollaborationProvider(workbook, transport);
    provider.onError((error) => providerErrors.push(error.message));
    provider.connect();

    await waitUntil(
      () => workbook?.cell(0, 4, 1).input === nativeBeforeJoin,
      'joining TypeScript XLSX peer to receive populated room state'
    );
    expect(provider.synced).toBe(false);
    expect(providerErrors).toEqual([]);
    console.log('XLSX join replay: TypeScript model received native pre-join cell');

    await native.command('reconnect');
    nativeState = await waitForNative(
      native,
      (snapshot) => snapshot.connected && snapshot.synced,
      'native XLSX handshake after retained-state replay'
    );
    provider.disconnect();
    provider.connect();
    await waitUntil(() => provider?.synced === true, 'TypeScript XLSX handshake with native peer');

    const nativeLive = 'native live';
    const typescriptLive = 'typescript live';
    const nativeEdit = native.command('setCell', {
      sheet: 0,
      row: 4,
      col: 2,
      input: nativeLive,
    });
    workbook.editCell(0, 4, 3, typescriptLive);
    await nativeEdit;
    await waitUntil(
      () => workbook?.cell(0, 4, 2).input === nativeLive,
      'TypeScript XLSX model to receive native live cell'
    );
    nativeState = await waitForNative(
      native,
      (snapshot) => nativeCell(snapshot, 0, 4, 3) === typescriptLive,
      'native XLSX model to receive TypeScript live cell'
    );
    expect(nativeCell(nativeState, 0, 4, 1)).toBe(nativeBeforeJoin);
    await assertXlsxConverged(native, workbook, 'XLSX live edits');

    nativeState = await native.command('disconnect');
    expect(nativeState.connected).toBe(false);
    const nativeOffline = 'native offline';
    const typescriptOffline = 'typescript while native offline';
    nativeState = await native.command('setCell', {
      sheet: 0,
      row: 4,
      col: 4,
      input: nativeOffline,
    });
    expect(nativeCell(nativeState, 0, 4, 4)).toBe(nativeOffline);
    expect(workbook.cell(0, 4, 4).input).not.toBe(nativeOffline);

    workbook.editCell(0, 4, 5, typescriptOffline);
    await Bun.sleep(100);
    nativeState = await native.command('snapshot');
    expect(nativeCell(nativeState, 0, 4, 5)).not.toBe(typescriptOffline);

    await native.command('reconnect');
    nativeState = await waitForNative(
      native,
      (snapshot) =>
        snapshot.connected &&
        snapshot.synced &&
        nativeCell(snapshot, 0, 4, 5) === typescriptOffline,
      'native XLSX reconnect and offline TypeScript cell'
    );
    await waitUntil(
      () => workbook?.cell(0, 4, 4).input === nativeOffline,
      'TypeScript XLSX model to receive reconnected native cell'
    );
    expect(nativeCell(nativeState, 0, 4, 4)).toBe(nativeOffline);
    expect(workbook.cell(0, 4, 5).input).toBe(typescriptOffline);
    await assertXlsxConverged(native, workbook, 'XLSX reconnected edits');
    expect(providerErrors).toEqual([]);
  } finally {
    provider?.destroy();
    transport?.disconnect();
    workbook?.dispose();
    await native?.stop();
    await stopRelay(relay.relay);
    await rm(WRANGLER_HOME, { recursive: true, force: true });
  }
}, 600_000);

test('native PPTX and the production TypeScript provider converge through the relay', async () => {
  await runChecked([process.execPath, resolve(ROOT, 'scripts/build-pptx-wasm.ts')], ROOT);
  const [{ initWasm, openPresentation }, { CollaborationProvider: PptxCollaborationProvider }] =
    await Promise.all([
      import('../../../packages/pptx/src/index'),
      import('../../../packages/pptx/src/collaboration/provider'),
    ]);
  await initWasm(new Uint8Array(await readFile(PPTX_WASM)));

  const port = await availablePort();
  const relay = await startRelay(port);
  const room = `native-typescript-pptx-${randomUUID()}`;
  let native: NativePeer | null = null;
  let presentation: PresentationHandle | null = null;
  let provider: InstanceType<typeof PptxCollaborationProvider> | null = null;
  let transport: ReturnType<typeof createRoomTransport> | null = null;
  const providerErrors: string[] = [];

  try {
    native = new NativePeer(relay.origin, room, PPTX_DOCUMENT);
    let nativeState = await waitForNative(
      native,
      (snapshot) => snapshot.connected,
      'native PPTX relay connection',
    );
    expect(nativeState.format).toBe('pptx');
    const target = nativeState.pptxStories.find((story) => story.length > 1);
    if (!target) throw new Error('native PPTX has no editable text story');
    const typescriptTarget = nativeState.pptxStories.find(
      (story) => story.storyId !== target.storyId && story.length > 1,
    );
    if (!typescriptTarget) throw new Error('native PPTX has no second editable text story');
    const nativeBeforeJoin = ' [native before join]';
    nativeState = await native.command('insertPptx', {
      storyId: target.storyId,
      index: target.text.length,
      text: nativeBeforeJoin,
    });
    expect(nativeState.pptxStories.find((story) => story.storyId === target.storyId)?.text).toContain(
      nativeBeforeJoin,
    );
    nativeState = await native.command('disconnect');
    expect(nativeState.connected).toBe(false);
    await Bun.sleep(100);

    presentation = openPresentation(new Uint8Array(await readFile(PPTX_DOCUMENT)), {
      clientId: 2_147_400_003,
      initialUpdate: new Uint8Array(await readFile(PPTX_SEED)),
    });
    transport = createRoomTransport(relay.origin, room);
    provider = new PptxCollaborationProvider(presentation, transport);
    provider.onError((error) => providerErrors.push(error.message));
    provider.connect();

    await waitUntil(
      () => pptxStoryText(presentation!, target.storyId).includes(nativeBeforeJoin),
      'joining TypeScript PPTX peer to receive populated room state',
    );
    expect(provider.synced).toBe(false);
    expect(providerErrors).toEqual([]);
    console.log('PPTX join replay: TypeScript model received native pre-join text');

    await native.command('reconnect');
    nativeState = await waitForNative(
      native,
      (snapshot) => snapshot.connected && snapshot.synced,
      'native PPTX handshake after retained-state replay',
    );
    provider.disconnect();
    provider.connect();
    await waitUntil(() => provider?.synced === true, 'TypeScript PPTX handshake with native peer');

    const nativeLive = ' [native live]';
    const typescriptLive = ' [typescript live]';
    const nativeTarget = nativeState.pptxStories.find(
      (story) => story.storyId === target.storyId,
    );
    if (!nativeTarget) throw new Error('native PPTX target story disappeared');
    const nativeEdit = native.command('insertPptx', {
      storyId: target.storyId,
      index: nativeTarget.text.length,
      text: nativeLive,
    });
    presentation.insertText(typescriptTarget.storyId, 0, typescriptLive);
    await nativeEdit;
    await waitUntil(
      () => pptxStoryText(presentation!, target.storyId).includes(nativeLive),
      'TypeScript PPTX model to receive native live text',
    );
    nativeState = await waitForNative(
      native,
      (snapshot) =>
        snapshot.pptxStories
          .find((story) => story.storyId === typescriptTarget.storyId)
          ?.text.includes(typescriptLive) === true,
      'native PPTX model to receive TypeScript live text',
    );
    expect(pptxStoryText(presentation, target.storyId)).toContain(nativeBeforeJoin);
    await assertPptxConverged(native, presentation, 'PPTX live edits');

    nativeState = await native.command('disconnect');
    expect(nativeState.connected).toBe(false);
    const nativeOffline = ' [native offline]';
    const typescriptOffline = ' [typescript while native offline]';
    const offlineTarget = nativeState.pptxStories.find(
      (story) => story.storyId === target.storyId,
    );
    if (!offlineTarget) throw new Error('native PPTX target story disappeared while offline');
    nativeState = await native.command('insertPptx', {
      storyId: target.storyId,
      index: offlineTarget.text.length,
      text: nativeOffline,
    });
    expect(
      nativeState.pptxStories.find((story) => story.storyId === target.storyId)?.text,
    ).toContain(nativeOffline);
    expect(pptxStoryText(presentation, target.storyId)).not.toContain(nativeOffline);

    presentation.insertText(typescriptTarget.storyId, 0, typescriptOffline);
    await Bun.sleep(100);
    nativeState = await native.command('snapshot');
    expect(
      nativeState.pptxStories.find((story) => story.storyId === typescriptTarget.storyId)?.text,
    ).not.toContain(typescriptOffline);

    await native.command('reconnect');
    nativeState = await waitForNative(
      native,
      (snapshot) =>
        snapshot.connected &&
        snapshot.synced &&
        snapshot.pptxStories
          .find((story) => story.storyId === typescriptTarget.storyId)
          ?.text.includes(typescriptOffline) === true,
      'native PPTX reconnect and offline TypeScript text',
    );
    await waitUntil(
      () => pptxStoryText(presentation!, target.storyId).includes(nativeOffline),
      'TypeScript PPTX model to receive reconnected native text',
    );
    expect(
      nativeState.pptxStories.find((story) => story.storyId === target.storyId)?.text,
    ).toContain(nativeOffline);
    expect(pptxStoryText(presentation, typescriptTarget.storyId)).toContain(typescriptOffline);
    await assertPptxConverged(native, presentation, 'PPTX reconnected edits');
    expect(providerErrors).toEqual([]);
  } finally {
    provider?.destroy();
    transport?.disconnect();
    presentation?.dispose();
    await native?.stop();
    await stopRelay(relay.relay);
    await rm(WRANGLER_HOME, { recursive: true, force: true });
  }
}, 600_000);

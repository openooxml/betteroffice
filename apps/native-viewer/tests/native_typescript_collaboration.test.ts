import { expect, test } from 'bun:test';
import { createHash, randomUUID } from 'node:crypto';
import { readFile, rm } from 'node:fs/promises';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';

import { createRoomTransport } from '../../demo/app/collab/createRoomTransport';
import { CollaborationProvider } from '../../../packages/docx/src/collaboration/provider';
import type { YrsSession } from '../../../packages/docx/src/yrs';

const ROOT = resolve(import.meta.dir, '../../..');
const NATIVE_VIEWER = resolve(ROOT, 'apps/native-viewer');
const RELAY = resolve(ROOT, 'apps/relay');
const DOCUMENT = resolve(ROOT, 'apps/demo/public/betteroffice-demo.docx');
const SEED = resolve(ROOT, 'apps/demo/public/seeds/docx.bin');
const WASM = resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm');
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
  connected: boolean;
  synced: boolean;
  stateVector: string;
  canonicalChecksum: string;
  stories: StoryState[];
  paragraphs: ParagraphState[];
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

  constructor(origin: string, room: string) {
    this.process = Bun.spawn({
      cmd: [
        'cargo',
        'run',
        '--quiet',
        '--manifest-path',
        resolve(NATIVE_VIEWER, 'Cargo.toml'),
        '--',
        '--document',
        DOCUMENT,
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

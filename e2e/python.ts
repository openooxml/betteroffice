import { resolve } from 'node:path';

import type { Detail, ScenarioRecorder, StageProfile } from './harness';
import { errorText } from './harness';
import { pythonWithBindings } from './python-env';

const WORKER = resolve(import.meta.dir, 'python', 'worker.py');
let availability: ReturnType<typeof pythonWithBindings> | undefined;

/** Why the cross-SDK scenarios must skip, or undefined when Python is ready. */
export function pythonMissing(): string | undefined {
  availability ??= pythonWithBindings();
  return 'missing' in availability ? availability.missing : undefined;
}

interface Response {
  id: number;
  ok: boolean;
  result?: unknown;
  timings?: StageProfile;
  error?: string;
}

type Worker = Bun.Subprocess<'pipe', 'pipe', 'inherit'>;

export class PythonWorker {
  private readonly process: Worker;
  private readonly reader: ReadableStreamDefaultReader<Uint8Array>;
  private readonly decoder = new TextDecoder();
  private buffer = '';
  private nextId = 1;
  private readonly spawnedAt: number;

  constructor(
    private readonly recorder: ScenarioRecorder,
    readonly actor = 'python'
  ) {
    const ready = pythonWithBindings();
    if ('missing' in ready) throw new Error(ready.missing);
    this.spawnedAt = performance.now();
    this.process = Bun.spawn([ready.python, '-u', WORKER], {
      stdin: 'pipe',
      stdout: 'pipe',
      stderr: 'inherit',
    }) as Worker;
    this.reader = this.process.stdout.getReader();
  }

  /** Waits for the interpreter; recorded as `python:start` from spawn to first reply. */
  async start(): Promise<void> {
    try {
      const response = await this.exchange(
        'def run(state, input, timed):\n    return "ready"\n',
        {}
      );
      if (!response.ok)
        throw new Error(response.error ?? 'Python worker startup failed');
      this.recorder.record({
        op: 'python:start',
        actor: this.actor,
        e2eMs: performance.now() - this.spawnedAt,
      });
    } catch (error) {
      this.recorder.record({
        op: 'python:start',
        actor: this.actor,
        e2eMs: performance.now() - this.spawnedAt,
        error: errorText(error),
      });
      throw error;
    }
  }

  /** Runs `script` (defining `run(state, input, timed)`) and records it as `op`. */
  async call<T>(
    op: string,
    script: string,
    input: unknown = {},
    detail?: Detail
  ): Promise<T> {
    const started = performance.now();
    let response: Response | undefined;
    try {
      response = await this.exchange(script, input);
      if (!response.ok)
        throw new Error(op + ' failed in Python:\n' + response.error);
      this.recorder.record({
        op,
        actor: this.actor,
        e2eMs: performance.now() - started,
        internal: response.timings,
        detail,
      });
      return response.result as T;
    } catch (error) {
      this.recorder.record({
        op,
        actor: this.actor,
        e2eMs: performance.now() - started,
        internal: response?.timings,
        detail,
        error: errorText(error),
      });
      throw error;
    }
  }

  async close(): Promise<void> {
    this.process.kill();
    await this.process.exited;
    await this.reader.cancel().catch(() => undefined);
  }

  private async exchange(script: string, input: unknown): Promise<Response> {
    const id = this.nextId++;
    this.process.stdin.write(JSON.stringify({ id, script, input }) + '\n');
    this.process.stdin.flush();
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try {
      return await Promise.race([
        this.readResponse(id),
        new Promise<never>((_resolve, reject) => {
          timeout = setTimeout(() => {
            this.process.kill();
            reject(
              new Error('Python worker did not respond within 60 seconds')
            );
          }, 60_000);
        }),
      ]);
    } finally {
      clearTimeout(timeout);
    }
  }

  private async readResponse(id: number): Promise<Response> {
    for (;;) {
      const newline = this.buffer.indexOf('\n');
      if (newline >= 0) {
        const line = this.buffer.slice(0, newline);
        this.buffer = this.buffer.slice(newline + 1);
        const response = JSON.parse(line) as Response;
        if (response.id !== id)
          throw new Error('Python response id does not match request');
        return response;
      }
      const { value, done } = await this.reader.read();
      if (done) throw new Error('the Python worker exited');
      this.buffer += this.decoder.decode(value, { stream: true });
    }
  }
}

export function toBase64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64');
}

export function fromBase64(text: string): Uint8Array {
  return new Uint8Array(Buffer.from(text, 'base64'));
}

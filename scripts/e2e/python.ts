/**
 * Drives the Python bindings from a scenario: a long-lived interpreter runs
 * one script per call, keeps state between calls, and every call is recorded
 * as an op with the round trip as e2e latency and Python's own stage timings
 * as the internal breakdown.
 */

import { resolve } from 'node:path';

import type { Detail, ScenarioRecorder, StageProfile } from './harness';
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

  constructor(private readonly recorder: ScenarioRecorder, readonly actor = 'python') {
    const ready = pythonWithBindings();
    if ('missing' in ready) throw new Error(ready.missing);
    this.spawnedAt = performance.now();
    this.process = Bun.spawn([ready.python, '-u', WORKER], { stdin: 'pipe', stdout: 'pipe', stderr: 'inherit' }) as Worker;
    this.reader = this.process.stdout.getReader();
  }

  /** Waits for the interpreter; recorded as `python:start` from spawn to first reply. */
  async start(): Promise<void> {
    await this.exchange('def run(state, input, timed):\n    return "ready"\n', {});
    this.recorder.record({ op: 'python:start', actor: this.actor, e2eMs: performance.now() - this.spawnedAt });
  }

  /** Runs `script` (defining `run(state, input, timed)`) and records it as `op`. */
  async call<T>(op: string, script: string, input: unknown = {}, detail?: Detail): Promise<T> {
    const started = performance.now();
    const response = await this.exchange(script, input);
    const e2eMs = performance.now() - started;
    if (!response.ok) throw new Error(`${op} failed in Python:\n${response.error}`);
    this.recorder.record({ op, actor: this.actor, e2eMs, internal: response.timings, detail });
    return response.result as T;
  }

  close(): void {
    try {
      this.process.stdin.end();
    } catch {}
    this.process.kill();
  }

  private async exchange(script: string, input: unknown): Promise<Response> {
    const id = this.nextId++;
    this.process.stdin.write(JSON.stringify({ id, script, input }) + '\n');
    this.process.stdin.flush();
    for (;;) {
      const newline = this.buffer.indexOf('\n');
      if (newline >= 0) {
        const line = this.buffer.slice(0, newline);
        this.buffer = this.buffer.slice(newline + 1);
        const response = JSON.parse(line) as Response;
        if (response.id === id) return response;
        continue;
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

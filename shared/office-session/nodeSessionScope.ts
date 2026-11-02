import type { SessionScope } from './transport';
import { missingWebWorker } from './webWorker';
import { compileWasm } from './wasm';
import { readWasmSync } from './readWasmAsset';

export { readWasmSync };

export async function compileSessionWasm(url: URL): Promise<WebAssembly.Module> {
  const bytes = readWasmSync(url);
  return bytes ? WebAssembly.compile(bytes as BufferSource) : compileWasm(url);
}

type WorkerEvent = 'message' | 'error' | 'messageerror';

interface ThreadWorker {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): unknown;
  on(type: WorkerEvent, listener: (value: unknown) => void): void;
}

interface WorkerThreads {
  Worker: new (url: URL, options: { name?: string; credentials?: RequestCredentials }) => ThreadWorker;
  parentPort: SessionScope | null;
}

export function workerThreads(inThread: string): WorkerThreads {
  const proc = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') throw missingWebWorker(inThread);
  const threads = proc.getBuiltinModule('node:worker_threads') as WorkerThreads | undefined;
  if (!threads) throw missingWebWorker(inThread);
  return threads;
}

export function sessionWorkerScope(): SessionScope {
  const scope = workerThreads('the in-thread API').parentPort;
  if (!scope) throw missingWebWorker('the in-thread API');
  return scope;
}

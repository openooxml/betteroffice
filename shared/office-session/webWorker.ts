/** The error worker sessions raise where the runtime has no Web Worker (Node.js), naming the in-thread API. */
export function missingWebWorker(inThread: string): Error {
  return new Error(`Worker sessions need a Web Worker (browsers, Bun, Deno), which this runtime lacks; use ${inThread} in-thread or pass the worker option`);
}

export function requireWebWorker(inThread: string): void {
  if (typeof Worker !== 'function') throw missingWebWorker(inThread);
}

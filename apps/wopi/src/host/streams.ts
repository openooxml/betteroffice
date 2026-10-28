export function cancelStream(
  stream: ReadableStream<Uint8Array> | ReadableStreamDefaultReader<Uint8Array> | null | undefined,
  abort?: () => void,
  reason?: unknown,
): void {
  try {
    stream?.cancel(reason).catch(() => {});
  } catch {}
  abort?.();
}

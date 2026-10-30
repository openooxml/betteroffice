export interface SocketSaveBarrier {
  connectionId: string;
  seq: number;
}

export type SaveBarrier = SocketSaveBarrier;

const SEQUENCE_FRAME_TYPE = 4;
const SEQUENCE_ACK_FRAME_TYPE = 5;
const CONNECTION_ID_BYTES = 36;
const HEADER_BYTES = 1 + CONNECTION_ID_BYTES + 8;
const CONNECTION_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function isSocketSaveBarrier(value: unknown): value is SocketSaveBarrier {
  if (typeof value !== "object" || value === null) return false;
  const barrier = value as Partial<SocketSaveBarrier>;
  return typeof barrier.connectionId === "string" && CONNECTION_ID_PATTERN.test(barrier.connectionId) &&
    typeof barrier.seq === "number" && Number.isSafeInteger(barrier.seq) && barrier.seq >= 0;
}

export function encodeSequencedFrame(barrier: SocketSaveBarrier, payload: Uint8Array = new Uint8Array()): Uint8Array {
  if (!isSocketSaveBarrier(barrier) || (barrier.seq === 0) !== (payload.byteLength === 0)) {
    throw new Error("invalidSaveBarrier");
  }
  const frame = new Uint8Array(HEADER_BYTES + payload.byteLength);
  frame[0] = SEQUENCE_FRAME_TYPE;
  frame.set(new TextEncoder().encode(barrier.connectionId), 1);
  new DataView(frame.buffer).setBigUint64(1 + CONNECTION_ID_BYTES, BigInt(barrier.seq));
  frame.set(payload, HEADER_BYTES);
  return frame;
}

export function decodeSequencedFrame(frame: Uint8Array): (SocketSaveBarrier & { payload: Uint8Array }) | null {
  if (frame[0] !== SEQUENCE_FRAME_TYPE) return null;
  if (frame.byteLength < HEADER_BYTES) throw new Error("invalidSaveBarrier");
  const connectionId = new TextDecoder().decode(frame.subarray(1, 1 + CONNECTION_ID_BYTES));
  const seq = Number(new DataView(frame.buffer, frame.byteOffset, frame.byteLength).getBigUint64(1 + CONNECTION_ID_BYTES));
  const payload = frame.subarray(HEADER_BYTES);
  const barrier = { connectionId, seq };
  if (!isSocketSaveBarrier(barrier) || (seq === 0) !== (payload.byteLength === 0)) throw new Error("invalidSaveBarrier");
  return { ...barrier, payload };
}

export function encodeSequenceAck(barrier: SocketSaveBarrier): Uint8Array {
  if (!isSocketSaveBarrier(barrier)) throw new Error("invalidSaveBarrier");
  const frame = new Uint8Array(HEADER_BYTES);
  frame[0] = SEQUENCE_ACK_FRAME_TYPE;
  frame.set(new TextEncoder().encode(barrier.connectionId), 1);
  new DataView(frame.buffer).setBigUint64(1 + CONNECTION_ID_BYTES, BigInt(barrier.seq));
  return frame;
}

export function decodeSequenceAck(frame: Uint8Array): SocketSaveBarrier | null {
  if (frame[0] !== SEQUENCE_ACK_FRAME_TYPE || frame.byteLength !== HEADER_BYTES) return null;
  const barrier = {
    connectionId: new TextDecoder().decode(frame.subarray(1, 1 + CONNECTION_ID_BYTES)),
    seq: Number(new DataView(frame.buffer, frame.byteOffset, frame.byteLength).getBigUint64(1 + CONNECTION_ID_BYTES)),
  };
  return isSocketSaveBarrier(barrier) ? barrier : null;
}

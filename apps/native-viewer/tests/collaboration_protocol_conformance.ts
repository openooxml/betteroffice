import {
  decodeMessages,
  encodeEmptyAwarenessUpdate,
  encodeQueryAwareness,
  encodeSyncStep1,
  encodeSyncStep2,
  encodeUpdate,
  encodeVarUint,
} from '../../../packages/docx/src/collaboration/protocol';

function concat(...parts: Uint8Array[]): Uint8Array {
  const output = new Uint8Array(parts.reduce((length, part) => length + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.byteLength;
  }
  return output;
}

function encodeAuth(reason: string): Uint8Array {
  const bytes = new TextEncoder().encode(reason);
  return concat(encodeVarUint(2), encodeVarUint(0), encodeVarUint(bytes.byteLength), bytes);
}

function normalize(message: ReturnType<typeof decodeMessages>[number]): unknown {
  switch (message.type) {
    case 'sync-step-1':
      return { type: message.type, stateVector: [...message.stateVector] };
    case 'sync-step-2':
    case 'update':
    case 'awareness':
      return { type: message.type, update: [...message.update] };
    case 'auth':
      return { type: message.type, reason: message.reason };
    case 'query-awareness':
      return { type: message.type };
  }
}

const request = (await Bun.stdin.json()) as { rustFrames: number[][] };
const decodedRustFrames = request.rustFrames.map((frame) => {
  const [message] = decodeMessages(Uint8Array.from(frame));
  return normalize(message);
});
const typescriptFrames = [
  encodeSyncStep1(Uint8Array.of(1, 2, 3)),
  encodeSyncStep2(new Uint8Array()),
  encodeUpdate(new Uint8Array(130).fill(7)),
  encodeEmptyAwarenessUpdate(),
  encodeAuth('denied 雪'),
  encodeQueryAwareness(),
].map((frame) => [...frame]);

process.stdout.write(JSON.stringify({ decodedRustFrames, typescriptFrames }));

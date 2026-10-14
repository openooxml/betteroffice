import { transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import type { BuildCommandFrameInput, DocxCommandFrame } from './residentCommandProtocol';

export function packResidentFrame(
  bytes: Uint8Array,
  precedingPageFrames: readonly Uint8Array[] = []
): TransferResult<{ frame: ArrayBuffer; pageFrames?: ArrayBuffer[] }> {
  const frame = exactFrameBuffer(bytes);
  const pageFrames = precedingPageFrames.map(exactFrameBuffer);
  return transferable({ frame, ...(pageFrames.length ? { pageFrames } : {}) },
    [...new Set([frame, ...pageFrames])]);
}

export function exactFrameBuffer(bytes: Uint8Array): ArrayBuffer {
  if (bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes.buffer;
  }
  return new Uint8Array(bytes).buffer;
}

export function buildCommandFrame(input: BuildCommandFrameInput): TransferResult<DocxCommandFrame> {
  const { frame: bytes, metadata: delta, pageFrames, coversThroughSeq, ...metadata } = input;
  const packed = packResidentFrame(bytes, pageFrames);
  const transfer = [...packed.transfer];
  if (metadata.layoutMeta) transfer.push(metadata.layoutMeta.pageSizes.buffer as ArrayBuffer);
  for (const page of metadata.stagedPresentation?.pages ?? []) {
    transfer.push(page.bitmap);
    for (const background of page.backgrounds ?? []) transfer.push(background.bitmap);
  }
  return transferable({
    coversThroughSeq,
    revision: delta.frameEpoch,
    ...(delta.full ? {} : { baseRevision: delta.baseFrameEpoch }),
    payload: { ...metadata, ...packed.value },
  }, [...new Set(transfer)]);
}

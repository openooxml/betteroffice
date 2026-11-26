import { shouldSlicePresentationFonts } from '../../../../shared/office-session/render/pptxSlicedFonts';
import type { preparePresentationFontDecoder } from './fontWorker';

type FontMode = 'disabled' | 'whole' | 'sliced';

interface PreparedFontWorkerOwner {
  retired(): boolean;
  shared(): boolean;
  take(): Worker;
  handle(): ReturnType<typeof preparePresentationFontDecoder>;
}

const preparedWorkers = new WeakMap<object, Map<FontMode, PreparedFontWorkerOwner>>();

export function fontMode(defaultFonts?: boolean, slicedFonts = shouldSlicePresentationFonts()): FontMode {
  return defaultFonts === false ? 'disabled' : slicedFonts ? 'sliced' : 'whole';
}

export function getOrCreatePreparedFontWorker(source: Uint8Array | ArrayBuffer, mode: FontMode,
  create: () => PreparedFontWorkerOwner): PreparedFontWorkerOwner {
  let workers = preparedWorkers.get(source);
  if (!workers) preparedWorkers.set(source, workers = new Map());
  let owner = workers.get(mode);
  if (!owner || owner.retired()) {
    owner = create();
    workers.set(mode, owner);
  }
  return owner;
}

export function takePreparedPresentationFontWorker(source: Uint8Array | ArrayBuffer, mode: FontMode): Worker | undefined {
  const workers = preparedWorkers.get(source);
  const owner = workers?.get(mode);
  if (owner?.shared() && !owner.retired()) return undefined;
  workers?.delete(mode);
  return owner && !owner.retired() ? owner.take() : undefined;
}

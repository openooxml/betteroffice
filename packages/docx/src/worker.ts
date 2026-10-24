export { SessionFailure } from '../../../shared/office-session/types';
export { createDocxCommandConnection } from './yrs/residentCommandConnection';
export { createDocxWorkerEditSession } from './yrs/createDocxWorkerEditSession';
export { discardCommandPresentation, closeCommandBitmap } from './yrs/residentCommandPresentation';
/** @internal */
export { indexEchoGlyphSet, leadingWordWidth, placeEchoGlyphs, type EchoGlyphIndex } from './yrs/residentEchoGlyphs';
export type * from './yrs/residentCommandProtocol';
export { DocxSourceWorkerClient, takePreloadedDocxSourceWorker } from './yrs/residentSourceWorkerClient';
export type { DocxEngineWorkerSourceOpened } from './yrs/residentSourceWorkerClient';

export { createResidentCommandConnection } from './yrs/residentCommandConnection';
export { createDocxWorkerEditSession } from './yrs/createDocxWorkerEditSession';
export { discardCommandPresentation, closeCommandBitmap } from './yrs/residentCommandPresentation';
/** @internal */
export { indexEchoGlyphSet, leadingWordWidth, placeEchoGlyphs, type EchoGlyphIndex } from './yrs/residentEchoGlyphs';
export type * from './yrs/residentCommandProtocol';
export { ResidentSourceWorkerClient, takePreloadedResidentSourceWorker } from './yrs/residentSourceWorkerClient';
export type { ResidentEngineWorkerSourceOpened } from './yrs/residentSourceWorkerClient';

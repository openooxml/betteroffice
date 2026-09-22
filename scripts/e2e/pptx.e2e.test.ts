/**
 * pptx end-to-end: pinned corpus decks through the scenarios under ./pptx,
 * every operation timed across the wasm boundary together with the stage
 * breakdown the renderer and the edit boundary report.
 */

import { defineSuite } from './suite';
import { context, setup } from './pptx/context';
import { scenarios } from './pptx/index';

defineSuite('pptx', scenarios, { setup, context });

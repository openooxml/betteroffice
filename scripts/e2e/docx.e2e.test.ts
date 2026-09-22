/**
 * docx end-to-end: pinned corpus documents through the scenarios under ./docx,
 * every operation timed across the wasm boundary together with the stage
 * breakdown the resident layout engine reports for it.
 */

import { defineSuite } from './suite';
import { context, setup } from './docx/context';
import { scenarios } from './docx/index';

defineSuite('docx', scenarios, { setup, context, timeoutMs: 180_000 });

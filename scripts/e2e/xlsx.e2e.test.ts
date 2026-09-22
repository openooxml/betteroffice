/**
 * xlsx end-to-end: pinned corpus workbooks through the scenarios under
 * ./xlsx, every operation timed across the wasm boundary together with the
 * stage breakdown the core reports for it.
 */

import { defineSuite } from './suite';
import { context, setup } from './xlsx/context';
import { scenarios } from './xlsx/index';

defineSuite('xlsx', scenarios, { setup, context });

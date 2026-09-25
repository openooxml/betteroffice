import { defineSuite } from './suite';
import { context, setup } from './docx/context';
import { scenarios } from './docx/index';

defineSuite('docx', scenarios, { setup, context, timeoutMs: 180_000 });

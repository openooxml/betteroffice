import { defineSuite } from './suite';
import { context, setup } from './pptx/context';
import { scenarios } from './pptx/index';

defineSuite('pptx', scenarios, { setup, context });

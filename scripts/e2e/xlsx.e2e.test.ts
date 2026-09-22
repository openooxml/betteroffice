import { defineSuite } from './suite';
import { context, setup } from './xlsx/context';
import { scenarios } from './xlsx/index';

defineSuite('xlsx', scenarios, { setup, context });

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { setGoogleFontsEnabled } from '../packages/docx/src/utils/fontLoader';
import { type ActScopeLog, logTestStart, openSummaryLine } from './react-act-scope';

// Unit tests stay offline: a Google Fonts stylesheet fetch that fails after its file's DOM teardown is an unhandled error.
setGoogleFontsEnabled(false);

// React flushes passive effects, unmount cleanups included, in scheduler macrotasks; they need the DOM globals to run.
const unregister = GlobalRegistrator.unregister.bind(GlobalRegistrator);
GlobalRegistrator.unregister = async () => {
  for (let turn = 0; turn < 3; turn++) await new Promise(resolve => setTimeout(resolve, 0));
  return unregister();
};

const ACT_SCOPE_LOG = Symbol.for('betteroffice.test.actScopeLog');
const writeStderr = (line: string) => {
  process.stderr.write(`${line}\n`);
};
const actScopeLog = (): ActScopeLog => {
  const existing: unknown = Reflect.get(process, ACT_SCOPE_LOG);
  if (typeof existing === 'object' && existing !== null) return existing as ActScopeLog;
  const log: ActScopeLog = { started: 0, open: 0 };
  Reflect.set(process, ACT_SCOPE_LOG, log);
  try {
    process.on('exit', () => {
      try {
        const line = openSummaryLine(log);
        if (line) writeStderr(line);
      } catch {}
    });
  } catch {}
  return log;
};
try {
  const { beforeEach }: typeof import('bun:test') = require('bun:test');
  beforeEach(() => {
    try {
      logTestStart(actScopeLog(), require.cache, writeStderr);
    } catch {}
  });
} catch {}

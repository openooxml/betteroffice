import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { setGoogleFontsEnabled } from '../packages/docx/src/utils/fontLoader';

// Unit tests stay offline: a Google Fonts stylesheet fetch that fails after its file's DOM teardown is an unhandled error.
setGoogleFontsEnabled(false);

// React flushes passive effects, unmount cleanups included, in scheduler macrotasks; they need the DOM globals to run.
const unregister = GlobalRegistrator.unregister.bind(GlobalRegistrator);
GlobalRegistrator.unregister = async () => {
  for (let turn = 0; turn < 3; turn++) await new Promise(resolve => setTimeout(resolve, 0));
  return unregister();
};

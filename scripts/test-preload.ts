import { setGoogleFontsEnabled } from '../packages/docx/src/utils/fontLoader';

// Unit tests stay offline: a Google Fonts stylesheet fetch that fails after its file's DOM teardown is an unhandled error.
setGoogleFontsEnabled(false);

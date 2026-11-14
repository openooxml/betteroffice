import { defineConfig } from 'tsup';
import { library } from './tsup.config';

/**
 * The `./shortcuts` subpath, built after the library so each entry's declarations stand alone and
 * the library's clean step cannot remove them.
 */
export default defineConfig({ ...library, entry: { shortcuts: 'src/shortcuts.ts' }, clean: false });

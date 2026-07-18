import path from 'path';
import { fileURLToPath } from 'url';

const __configDir = path.dirname(fileURLToPath(import.meta.url));

/**
 * Shared color/theme palette lives in the core preset (single source of truth);
 * this config only adds the React-specific content glob + .oox-root scoping.
 * @type {import('tailwindcss').Config}
 */
export default {
  presets: [require('../docx/tailwind-preset.cjs')],
  // Scope all utilities under .oox-root to avoid clashing with host app CSS
  important: '.oox-root',
  // Only scan library source files, not demo. Absolute path so this works no
  // matter where `tailwindcss -c` is invoked from (see issue #340).
  content: [path.join(__configDir, 'src/**/*.{ts,tsx}')],
};

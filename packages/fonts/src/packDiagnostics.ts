import intlSizes from './intl-sizes.json';
import latinManifest from '../../fonts-latin/assets/manifest.json';
import { BUNDLED_FONTS } from './manifest';

/** An installed font pack that cannot be used, so its faces are substituted. */
export interface FontPackDiagnostic {
  pack: 'latin' | 'intl' | 'cjk';
  /** `version-mismatch`: the pack is another release; `missing-file`: files are absent; `invalid-manifest`: its manifest cannot be read. */
  reason: 'version-mismatch' | 'missing-file' | 'invalid-manifest';
  detail: string;
}

type Pack = FontPackDiagnostic['pack'];
type AssetUrls = Record<string, () => URL>;

const PACK_FILES: Record<Pack, string[]> = {
  latin: latinManifest.faces.map((face) => face.file),
  intl: Object.keys(intlSizes),
  cjk: BUNDLED_FONTS.filter((face) => face.script?.startsWith('cjk-')).map((face) => face.file),
};
const PACK_VERSION = latinManifest.version;

interface NodeFs {
  readFileSync(path: string, encoding: 'utf8'): string;
  statSync(path: string, options: { throwIfNoEntry: false }): { isFile(): boolean } | undefined;
}

function nodeModule<T>(name: string): T | undefined {
  const proc = (globalThis as { process?: { getBuiltinModule?: (id: string) => unknown } }).process;
  if (typeof proc?.getBuiltinModule !== 'function') return undefined;
  try {
    return proc.getBuiltinModule(name) as T;
  } catch {
    return undefined;
  }
}

const message = (error: unknown) => (error instanceof Error ? error.message : String(error));

function diagnostic(pack: Pack, reason: FontPackDiagnostic['reason'], detail: string): FontPackDiagnostic {
  return { pack, reason, detail: `@betteroffice/fonts-${pack} ${detail}; its faces are substituted` };
}

function missingFiles(pack: Pack, files: string[]): FontPackDiagnostic | undefined {
  const sorted = [...files].sort();
  if (!sorted.length) return undefined;
  const more = sorted.length > 1 ? ` and ${sorted.length - 1} more files` : '';
  return diagnostic(pack, 'missing-file', `is missing ${sorted[0]}${more}`);
}

/** Checks an installed pack from its directory on disk, as the native engines do; undefined outside Node. */
function checkOnDisk(pack: Pack, asset: URL): FontPackDiagnostic | null | undefined {
  const fs = nodeModule<NodeFs>('node:fs');
  const url = nodeModule<{ fileURLToPath(url: string): string }>('node:url');
  if (asset.protocol !== 'file:' || !fs || !url) return undefined;
  const path = (file: string) => url.fileURLToPath(new URL(file, asset).href);
  const absent = (files: string[]) => files.filter((file) => !fs.statSync(path(file), { throwIfNoEntry: false })?.isFile());
  let text: string;
  try {
    text = fs.readFileSync(path('manifest.json'), 'utf8');
  } catch (error) {
    if ((error as { code?: unknown }).code === 'ENOENT') {
      return missingFiles(pack, [...absent(PACK_FILES[pack]), 'manifest.json']) ?? null;
    }
    return diagnostic(pack, 'invalid-manifest', `has an unreadable manifest (${message(error)})`);
  }
  let version: unknown;
  try {
    version = (JSON.parse(text) as { version?: unknown } | null)?.version;
  } catch (error) {
    return diagnostic(pack, 'invalid-manifest', `has an unreadable manifest (${message(error)})`);
  }
  if (version !== PACK_VERSION) {
    const found = typeof version === 'string' ? version : 'of unknown version';
    return diagnostic(pack, 'version-mismatch', `${found} does not match the ${PACK_VERSION} release the engine needs`);
  }
  return missingFiles(pack, absent(PACK_FILES[pack])) ?? null;
}

/**
 * The installed font packs that cannot be used, and why. A pack that is not installed is not reported. Under Node
 * the pack's manifest and files are checked on disk; in a browser only its asset list is compared with this
 * release, since checking the files would fetch them.
 */
export async function checkInstalledPacks(importers: Record<Pack, () => Promise<AssetUrls | undefined>>): Promise<FontPackDiagnostic[]> {
  const diagnostics: FontPackDiagnostic[] = [];
  for (const pack of ['latin', 'intl', 'cjk'] as const) {
    let urls: AssetUrls | undefined;
    try {
      urls = await importers[pack]();
    } catch (error) {
      diagnostics.push(diagnostic(pack, 'invalid-manifest', `has an unreadable asset list (${message(error)})`));
      continue;
    }
    if (!urls) continue;
    const listed = PACK_FILES[pack].filter((file) => file in urls);
    const first = listed[0] ?? Object.keys(urls)[0];
    const onDisk = first === undefined ? undefined : checkOnDisk(pack, urls[first]!());
    if (onDisk !== undefined) {
      if (onDisk) diagnostics.push(onDisk);
      continue;
    }
    if (listed.length < PACK_FILES[pack].length) {
      diagnostics.push(diagnostic(pack, 'version-mismatch', `of unknown version does not match the ${PACK_VERSION} release the engine needs`));
    }
  }
  return diagnostics;
}

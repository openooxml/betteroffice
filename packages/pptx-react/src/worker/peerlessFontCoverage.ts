import type { PaintFontCoverage } from '@betteroffice/pptx';
import type { PresentationCommandClient } from '@betteroffice/pptx/worker';
import { readFontTraits } from '../fontTraits';

const KEPT_REVISIONS = 2;
let sessions = 0;

interface PageFace {
  revision: number;
  family: string;
  face: FontFace;
}

interface Family {
  faces: PageFace[];
}

interface Batch {
  texts: Set<string>;
  flush?: Promise<string>;
}

const keyOf = (family: string) => family.trim().toLowerCase();
const pageFonts = () => typeof document === 'undefined' ? undefined : document.fonts;

/**
 * The page's side of sliced faces: the worker owns them and fetches their chunks; the page keeps a copy of each
 * revision a paint needed, under its own family name. Text from one paint's tiles is asked for together, and
 * thumbnails batch apart, so the slide being edited never waits for a chunk only a thumbnail needs.
 */
export function createPeerlessFontCoverage(
  client: Pick<PresentationCommandClient, 'readCoverageFont' | 'state'>, current: () => boolean,
): PaintFontCoverage & { readonly thumbnails: PaintFontCoverage; dispose(): void } {
  const families = new Map<string, Family>();
  const session = ++sessions;
  const isSliced = (key: string) => client.state.slicedFamilies?.some((family) => keyOf(family) === key);

  async function request(name: string, entry: Family, batch: Batch): Promise<string> {
    const key = keyOf(name);
    const currentEntry = () => current() && families.get(key) === entry && isSliced(key);
    await new Promise((resolve) => setTimeout(resolve, 0));
    const text = [...batch.texts].join('\n');
    batch.texts.clear();
    batch.flush = undefined;
    const have = entry.faces[entry.faces.length - 1]?.revision ?? null;
    const read = await client.readCoverageFont?.(name, text, have).catch(() => null);
    if (!read || !currentEntry()) return name;
    if (!read.bytes) return entry.faces.find(({ revision }) => revision === read.revision)?.family ?? name;
    const family = `${name} s${session} r${read.revision}`;
    const fonts = pageFonts();
    if (!fonts || typeof FontFace !== 'function') return name;
    try {
      const traits = readFontTraits(new Uint8Array(read.bytes));
      const face = await new FontFace(family, read.bytes, traits ? {
        weight: String(traits.weight), style: traits.italic ? 'italic' : traits.oblique ? 'oblique' : 'normal',
      } : {}).load();
      if (!currentEntry()) return name;
      const newest = entry.faces[entry.faces.length - 1];
      if (newest && newest.revision >= read.revision) return newest.family;
      fonts.add(face);
      entry.faces.push({ revision: read.revision, family, face });
      while (entry.faces.length > KEPT_REVISIONS) fonts.delete(entry.faces.shift()!.face);
      return family;
    } catch { return name; }
  }

  const latest = (name: string) => {
    const entry = families.get(keyOf(name));
    return entry?.faces[entry.faces.length - 1]?.family;
  };

  function lane(): PaintFontCoverage {
    const batches = new Map<string, Batch>();
    return {
      prepare(name, text) {
        const key = keyOf(name);
        const fonts = pageFonts();
        if (!isSliced(key) || !client.readCoverageFont || !fonts || typeof FontFace !== 'function') {
          const replaced = families.get(key);
          if (replaced) for (const { face } of replaced.faces) fonts?.delete(face);
          families.delete(key);
          return Promise.resolve(name);
        }
        let entry = families.get(key);
        if (!entry) families.set(key, entry = { faces: [] });
        let batch = batches.get(name);
        if (!batch) batches.set(name, batch = { texts: new Set() });
        batch.texts.add(document.documentElement.dir === 'rtl' ? `${text}\u200f` : text);
        return batch.flush ??= request(name, entry, batch);
      },
      latest,
      covers: (name, bold, italic) => !client.state.slicedPlainStyles?.includes(`${keyOf(name)}|${bold ? 1 : 0}|${italic ? 1 : 0}`),
    };
  }

  return {
    ...lane(),
    thumbnails: lane(),
    dispose() {
      const fonts = pageFonts();
      for (const entry of families.values()) for (const { face } of entry.faces) fonts?.delete(face);
      families.clear();
    },
  };
}


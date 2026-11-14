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

  async function request(name: string, entry: Family, batch: Batch): Promise<string> {
    await new Promise((resolve) => setTimeout(resolve, 0));
    const text = [...batch.texts].join('\n');
    batch.texts.clear();
    batch.flush = undefined;
    const latest = entry.faces[entry.faces.length - 1];
    const read = await client.readCoverageFont?.(name, text, latest?.revision ?? null).catch(() => null);
    if (!read) return name;
    if (!read.bytes) return latest?.family ?? name;
    const family = `${name} s${session} r${read.revision}`;
    const traits = readFontTraits(new Uint8Array(read.bytes));
    const face = await new FontFace(family, read.bytes, traits ? {
      weight: String(traits.weight), style: traits.italic ? 'italic' : traits.oblique ? 'oblique' : 'normal',
    } : {}).load();
    if (!current() || families.get(keyOf(name)) !== entry) return name;
    const newest = entry.faces[entry.faces.length - 1];
    if (newest && newest.revision >= read.revision) return newest.family;
    document.fonts.add(face);
    entry.faces.push({ revision: read.revision, family, face });
    while (entry.faces.length > KEPT_REVISIONS) document.fonts.delete(entry.faces.shift()!.face);
    return family;
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
        if (!client.state.slicedFamilies?.some((family) => keyOf(family) === key) || !client.readCoverageFont ||
          typeof document === 'undefined') {
          const replaced = families.get(key);
          if (replaced) for (const { face } of replaced.faces) document.fonts.delete(face);
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
      for (const entry of families.values()) for (const { face } of entry.faces) document.fonts?.delete(face);
      families.clear();
    },
  };
}

/** Installs complete faces for paints that do not prepare their text; callers during a load share it, and it reads again for them. */
export function createWholeFontInstaller<Face extends { family: string }>(
  read: (have: string[]) => Promise<readonly Face[]>,
  install: (faces: Face[]) => Promise<readonly { family: string }[]>,
  supplied: (family: string) => boolean,
): () => Promise<void> {
  const installed = new Set<string>();
  let loading: Promise<void> | undefined;
  let askedAgain = false;
  return () => {
    if (loading) {
      askedAgain = true;
      return loading;
    }
    return loading = (async () => {
      do {
        askedAgain = false;
        const faces = (await read([...installed])).filter((face) => !supplied(face.family));
        if (faces.length) for (const face of await install(faces)) installed.add(face.family);
      } while (askedAgain);
    })().finally(() => { loading = undefined; });
  };
}


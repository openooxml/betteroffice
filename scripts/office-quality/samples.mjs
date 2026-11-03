export const CORPUS_ORIGIN = 'https://corpus.betteroffice.dev';
export const DEFAULT_COLLECTION = 'office-quality';
export const MAX_SAMPLES = 8000;

const LOOPBACK = new Set(['127.0.0.1', 'localhost', '[::1]']);

function isFolderName(value) {
  return typeof value === 'string' && value.length > 0 && !/[^a-z0-9-]/.test(value);
}

export function corpusOrigin(environment = process.env) {
  const value = environment.QUALITY_CORPUS_ORIGIN?.trim();
  if (!value) return CORPUS_ORIGIN;
  let url;
  try {
    url = new URL(value);
  } catch {
    url = null;
  }
  if (
    !url ||
    url.origin !== value.replace(/\/$/, '') ||
    !(url.protocol === 'https:' || (url.protocol === 'http:' && LOOPBACK.has(url.hostname)))
  )
    throw new Error('QUALITY_CORPUS_ORIGIN must be an https origin or a loopback http origin');
  return url.origin;
}

function validateSamples(ids, label) {
  if (
    !Array.isArray(ids) ||
    !ids.length ||
    ids.length > MAX_SAMPLES ||
    ids.some((id) => !isFolderName(id)) ||
    new Set(ids).size !== ids.length
  )
    throw new Error(`${label} must contain 1–${MAX_SAMPLES} unique sample folder names`);
  return ids;
}

export async function selectSamples(environment, download) {
  const origin = corpusOrigin(environment);
  const id = environment.QUALITY_COLLECTION?.trim() || DEFAULT_COLLECTION;
  if (!isFolderName(id))
    throw new Error('QUALITY_COLLECTION must be a collection folder name');
  if (environment.QUALITY_SAMPLES?.trim())
    return validateSamples(JSON.parse(environment.QUALITY_SAMPLES), 'QUALITY_SAMPLES');

  const manifest = JSON.parse(
    await download(`${origin}/collections/${id}.json`, 2 * 1024 * 1024)
  );
  if (manifest?.schema_version !== 1 || manifest.id !== id)
    throw new Error(`Invalid corpus collection: ${id}`);
  return validateSamples(manifest.samples, `Collection ${id}`);
}

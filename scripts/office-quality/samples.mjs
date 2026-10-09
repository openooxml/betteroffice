export const CORPUS_ORIGIN = 'https://corpus.betteroffice.dev';
export const MAX_SAMPLES = 1000;

function isFolderName(value) {
  return typeof value === 'string' && value.length > 0 && !/[^a-z0-9-]/.test(value);
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
  const collection = environment.QUALITY_COLLECTION?.trim();
  if (collection && collection !== 'office-quality')
    throw new Error(
      'QUALITY_COLLECTION is fixed to office-quality; use QUALITY_FORMAT or QUALITY_SAMPLES'
    );
  if (environment.QUALITY_SAMPLES?.trim())
    return validateSamples(JSON.parse(environment.QUALITY_SAMPLES), 'QUALITY_SAMPLES');

  const id = 'office-quality';
  const manifest = JSON.parse(
    await download(`${CORPUS_ORIGIN}/collections/${id}.json`, 2 * 1024 * 1024)
  );
  if (manifest?.schema_version !== 1 || manifest.id !== id)
    throw new Error(`Invalid corpus collection: ${id}`);
  return validateSamples(manifest.samples, `Collection ${id}`);
}

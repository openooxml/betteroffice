export const CORPUS_ORIGIN = 'https://corpus.betteroffice.dev';

function isFolderName(value) {
  return typeof value === 'string' && value.length > 0 && !/[^a-z0-9-]/.test(value);
}

function validateSamples(ids, label) {
  if (
    !Array.isArray(ids) ||
    !ids.length ||
    ids.length > 100 ||
    ids.some((id) => !isFolderName(id)) ||
    new Set(ids).size !== ids.length
  )
    throw new Error(`${label} must contain 1–100 unique sample folder names`);
  return ids;
}

export async function selectSamples(environment, download) {
  if (environment.QUALITY_SAMPLES?.trim())
    return validateSamples(JSON.parse(environment.QUALITY_SAMPLES), 'QUALITY_SAMPLES');

  const id = environment.QUALITY_COLLECTION || 'office-quality';
  if (!isFolderName(id))
    throw new Error('QUALITY_COLLECTION must be a collection folder name');
  const manifest = JSON.parse(
    await download(`${CORPUS_ORIGIN}/collections/${id}.json`, 2 * 1024 * 1024)
  );
  if (manifest?.schema_version !== 1 || manifest.id !== id)
    throw new Error(`Invalid corpus collection: ${id}`);
  return validateSamples(manifest.samples, `Collection ${id}`);
}

import { stripVTControlCharacters } from 'node:util';

export function validateComparison(result) {
  if (
    result.source_verified !== true ||
    result.reference?.status !== 'ok' ||
    result.actual?.status !== 'ok' ||
    typeof result.reference.sha256 !== 'string' ||
    !result.reference.sha256 ||
    result.reference.sha256 !== result.actual.sha256 ||
    !Number.isFinite(result.penalized_ssim) ||
    Math.abs(result.penalized_ssim) > 1 ||
    result.resized === true
  )
    throw new Error('Cannot publish an incomplete or mismatched comparison');
}

export function summarizeError(error) {
  let message;
  try {
    const output = JSON.parse(String(error?.stdout ?? ''));
    if (typeof output.error === 'string') message = output.error;
  } catch {}
  if (!message) {
    const lines = stripVTControlCharacters(String(error?.stderr ?? ''))
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);
    message =
      lines.findLast((line) => /\berror:|\b\w*Error:/i.test(line)) ??
      lines.at(-1) ??
      String(error?.message ?? error ?? 'Unknown error');
  }
  if (message.includes('Unexpected external requests'))
    return 'Unexpected external requests during capture';
  return (
    stripVTControlCharacters(message)
      .split('\n')[0]
      .replace(/https?:\/\/[^\s"'<>]+/g, '[url]')
      .replace(/(?<![\w@])(?:file:\/\/)?(?:\/|[a-z]:[\\/])[^\s"'<>|]+/gi, '[path]')
      .replace(/[\u0000-\u001f\u007f]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 240) || 'Unknown error'
  );
}

export async function measureSamples(
  samples,
  identity,
  { capture, compare, log = console.log }
) {
  for (const sample of samples) {
    let stage = 'capture';
    let result;
    try {
      await capture(sample);
      stage = 'compare';
      const comparison = await compare(sample);
      validateComparison(comparison);
      result = { ...comparison, ...identity, status: 'ok' };
    } catch (error) {
      result = {
        ...identity,
        status: 'failed',
        stage,
        error: summarizeError(error),
      };
    }
    sample.comparisons.push(result);
    log(
      `${sample.id} ${identity.channel}: ${
        result.status === 'ok'
          ? result.penalized_ssim.toFixed(4)
          : `FAILED (${result.stage}): ${result.error}`
      }`
    );
  }
}

import { createHash } from 'node:crypto';

export async function r2Options(environment, request = fetch) {
  const accountId = environment.CLOUDFLARE_ACCOUNT_ID?.trim();
  const token = environment.CLOUDFLARE_API_TOKEN?.trim();
  if (!/^[a-f0-9]{32}$/.test(accountId ?? '') || !token)
    throw new Error('CLOUDFLARE_ACCOUNT_ID and CLOUDFLARE_API_TOKEN are required');

  const endpoints = [`accounts/${accountId}/tokens/verify`, 'user/tokens/verify'];
  for (const [index, endpoint] of endpoints.entries()) {
    const response = await request(`https://api.cloudflare.com/client/v4/${endpoint}`, {
      headers: { Authorization: `Bearer ${token}` },
      redirect: 'error',
      signal: AbortSignal.timeout(30_000),
    });
    if (index === 0 && [401, 403].includes(response.status)) {
      await response.body?.cancel();
      continue;
    }
    if (!response.ok)
      throw new Error(`Cloudflare token verification failed (HTTP ${response.status})`);
    const result = await response.json();
    if (!result.success || result.result?.status !== 'active' ||
        !/^[a-f0-9]{32}$/.test(result.result?.id ?? ''))
      throw new Error('Cloudflare token verification did not return an active token ID');
    return {
      accessKeyId: result.result.id,
      secretAccessKey: createHash('sha256').update(token).digest('hex'),
      endpoint: `https://${accountId}.r2.cloudflarestorage.com`,
      region: 'auto',
      bucket: environment.QUALITY_RENDER_BUCKET,
    };
  }
  throw new Error('Cloudflare token verification failed');
}

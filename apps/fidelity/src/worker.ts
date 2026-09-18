export interface Env {
  ASSETS: Fetcher;
  RENDERS?: R2Bucket;
}

const PREFIX = 'renders/';
const IMMUTABLE = 'public, max-age=31536000, immutable';
const POINTER = 'public, max-age=0, must-revalidate';

function key(pathname: string): string | null {
  const decoded = decodeURIComponent(pathname.slice(1));
  if (!decoded.startsWith(PREFIX) || decoded.includes('..') || decoded.includes('//')) return null;
  return decoded;
}

async function renders(request: Request, env: Env, path: string): Promise<Response> {
  if (!env.RENDERS) return new Response('Renders are not configured', { status: 503 });
  const object = await env.RENDERS.get(path, { onlyIf: request.headers });
  if (!object) return new Response('Not found', { status: 404 });
  const headers = new Headers();
  object.writeHttpMetadata(headers);
  headers.set('etag', object.httpEtag);
  headers.set('cache-control', path.endsWith('.json') ? POINTER : IMMUTABLE);
  if (!('body' in object)) return new Response(null, { status: 304, headers });
  return new Response(object.body, { headers });
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (request.method !== 'GET' && request.method !== 'HEAD')
      return new Response('Method not allowed', { status: 405 });
    const { pathname } = new URL(request.url);
    if (!pathname.startsWith(`/${PREFIX}`)) return env.ASSETS.fetch(request);
    const path = key(pathname);
    return path ? renders(request, env, path) : new Response('Not found', { status: 404 });
  },
};

export interface Env {
  ASSETS: Fetcher;
  RENDERS?: R2Bucket;
}

const HOST = 'benchmarks.betteroffice.dev';
const LEGACY_HOST = 'fidelity.betteroffice.dev';
const PREFIXES = ['renders/', 'e2e/'];
const IMMUTABLE = 'public, max-age=31536000, immutable';
const POINTER = 'public, max-age=0, must-revalidate';

function allowed(path: string): boolean {
  return (
    PREFIXES.some((prefix) => path.startsWith(prefix)) && !path.includes('..') && !path.includes('//')
  );
}

/** The old host served the page viewer at its root, so its links land on the viewer. */
function moved(url: URL): Response {
  const target = new URL(`https://${HOST}`);
  target.pathname = url.pathname === '/' ? '/compare' : url.pathname;
  target.search = url.search;
  return Response.redirect(target.toString(), 301);
}

async function stored(request: Request, env: Env, path: string): Promise<Response> {
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
    const url = new URL(request.url);
    if (url.hostname === LEGACY_HOST) return moved(url);
    const { pathname } = url;
    if (!PREFIXES.some((prefix) => pathname.startsWith(`/${prefix}`))) return env.ASSETS.fetch(request);
    let path: string;
    try {
      path = decodeURIComponent(pathname.slice(1));
    } catch {
      return new Response('Invalid path', { status: 400 });
    }
    return allowed(path) ? stored(request, env, path) : new Response('Not found', { status: 404 });
  },
};

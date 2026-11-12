import { createHash, randomBytes } from 'node:crypto';
import { existsSync } from 'node:fs';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createServer, type Server, type Socket } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, expect, test } from 'bun:test';

const script = new URL('./fetch-libreoffice.sh', import.meta.url).pathname;
const pin = new URL('./libreoffice-pin.sh', import.meta.url).pathname;
const install = new URL('./install-libreoffice.sh', import.meta.url).pathname;
const payload = randomBytes(256 * 1024);
const sha = createHash('sha256').update(payload).digest('hex');

type Hit = { path: string; range?: number };
type Route = (socket: Socket, hit: Hit, count: number) => void;

function head(socket: Socket, status: number, length: number, extra: string[] = []) {
  socket.write([`HTTP/1.1 ${status} X`, `Content-Length: ${length}`, 'Connection: close', ...extra, '', ''].join('\r\n'));
}

function send(socket: Socket, status: number, body: Buffer, extra: string[] = []) {
  head(socket, status, body.length, extra);
  socket.end(body);
}

function file(body: Buffer): Route {
  return (socket, { range }) => {
    if (range === undefined) return send(socket, 200, body);
    if (range >= body.length) return send(socket, 416, Buffer.alloc(0), [`Content-Range: bytes */${body.length}`]);
    send(socket, 206, body.subarray(range), [`Content-Range: bytes ${range}-${body.length - 1}/${body.length}`]);
  };
}

function cut(body: Buffer, at: number): Route {
  const whole = file(body);
  return (socket, hit, count) => {
    if (count > 1) return whole(socket, hit, count);
    head(socket, 200, body.length);
    socket.end(body.subarray(0, at));
  };
}

function trickle(body: Buffer, bytes: number, every: number): Route {
  return (socket, { range = 0 }) => {
    const rest = body.subarray(range);
    head(socket, range ? 206 : 200, rest.length, range ? [`Content-Range: bytes ${range}-${body.length - 1}/${body.length}`] : []);
    let sent = 0;
    const timer = setInterval(() => {
      if (socket.destroyed) return clearInterval(timer);
      socket.write(rest.subarray(sent, sent + bytes));
      sent += bytes;
      if (sent >= rest.length) {
        clearInterval(timer);
        socket.end();
      }
    }, every);
    socket.on('close', () => clearInterval(timer));
  };
}

const missing: Route = (socket) => send(socket, 404, Buffer.from('not found'));

let server: Server | undefined;
let root: string | undefined;
const sockets = new Set<Socket>();

afterEach(async () => {
  for (const socket of sockets) socket.destroy();
  sockets.clear();
  await new Promise((resolve) => (server ? server.close(resolve) : resolve(undefined)));
  server = undefined;
  if (root) await rm(root, { recursive: true, force: true });
  root = undefined;
});

async function serve(routes: Record<string, Route>) {
  const hits: Hit[] = [];
  server = createServer((socket) => {
    let text = '';
    sockets.add(socket);
    socket.on('close', () => sockets.delete(socket));
    socket.on('error', () => {});
    socket.on('data', (chunk) => {
      if (text.includes('\r\n\r\n')) return;
      text += chunk.toString('latin1');
      if (!text.includes('\r\n\r\n')) return;
      const [line, ...headers] = text.split('\r\n');
      const range = headers.map((header) => /^range:\s*bytes=(\d+)-/i.exec(header)).find(Boolean);
      const hit = { path: line.split(' ')[1], ...(range ? { range: Number(range[1]) } : {}) };
      hits.push(hit);
      (routes[hit.path] ?? missing)(socket, hit, hits.filter(({ path }) => path === hit.path).length);
    });
  });
  await new Promise<void>((resolve) => server!.listen(0, '127.0.0.1', resolve));
  const { port } = server.address() as { port: number };
  return { hits, url: (path: string) => `http://127.0.0.1:${port}${path}` };
}

async function fetchArchive(urls: string[], env: Record<string, string> = {}) {
  root ??= await mkdtemp(join(tmpdir(), 'fetch-libreoffice-'));
  const dest = join(root, 'cache', 'LibreOffice.tar.gz');
  const child = Bun.spawn(['bash', script, dest, sha, ...urls], {
    env: { PATH: process.env.PATH!, HOME: root, NO_PROXY: '*', no_proxy: '*', LO_FETCH_PAUSE: '0', ...env },
    stdout: 'pipe',
    stderr: 'pipe',
  });
  const [status, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()]);
  return { status, log: stdout + stderr, dest };
}

async function digestOf(path: string) {
  return createHash('sha256').update(await readFile(path)).digest('hex');
}

test('the pin is one version, URL pair and hash shared by the fetch helper and the installer', async () => {
  const child = Bun.spawn(['bash', '-c', `HOME=/home/runner; source "${pin}"; printf '%s\\n' "$LO_VERSION" "$LO_SHA256" "$LO_ARCHIVE_PATH" "\${LO_MIRRORS[@]}"`], { stdout: 'pipe' });
  expect((await new Response(child.stdout).text()).trim().split('\n')).toEqual([
    '26.2.3.2',
    '18838cb9d028b664a9d0e966cd4c8ca47ca3ea363c393b41d1b5124740b121a5',
    '/home/runner/.cache/libreoffice-archive/LibreOffice_26.2.3.2_Linux_x86-64_deb.tar.gz',
    'https://download.documentfoundation.org/libreoffice/stable/26.2.3/deb/x86_64/LibreOffice_26.2.3.2_Linux_x86-64_deb.tar.gz',
    'https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.3.2/deb/x86_64/LibreOffice_26.2.3.2_Linux_x86-64_deb.tar.gz',
  ]);
  for (const path of [script, install]) expect(await readFile(path, 'utf8')).toContain('libreoffice-pin.sh');
  const installer = await readFile(install, 'utf8');
  expect(installer).not.toContain('curl');
  const check = installer.indexOf('sha256sum --check');
  const extract = installer.indexOf('tar -xzf');
  expect(check).toBeGreaterThanOrEqual(0);
  expect(extract).toBeGreaterThanOrEqual(0);
  expect(check).toBeLessThan(extract);
});

test('a good first mirror delivers the archive in one request', async () => {
  const { hits, url } = await serve({ '/good': file(payload), '/spare': file(payload) });
  const { status, log, dest } = await fetchArchive([url('/good'), url('/spare')]);
  expect({ status, log }).toEqual({ status: 0, log: expect.stringContaining(`mirror ${url('/good')}: ok`) });
  expect(await digestOf(dest)).toBe(sha);
  expect(existsSync(`${dest}.part`)).toBe(false);
  expect(hits.map(({ path }) => path)).toEqual(['/good']);
}, 30_000);

test('a wrong hash from the first mirror is deleted and the second mirror delivers', async () => {
  const { url } = await serve({ '/wrong': file(randomBytes(payload.length)), '/good': file(payload) });
  const { status, log, dest } = await fetchArchive([url('/wrong'), url('/good')]);
  expect(status).toBe(0);
  expect(log).toContain(`mirror ${url('/wrong')}: sha256 mismatch`);
  expect(log).toContain(`mirror ${url('/good')}: ok`);
  expect(await digestOf(dest)).toBe(sha);
}, 30_000);

test('a 404 from the first mirror falls through to the second', async () => {
  const { url } = await serve({ '/good': file(payload) });
  const { status, log, dest } = await fetchArchive([url('/gone'), url('/good')]);
  expect(status).toBe(0);
  expect(log).toContain(`mirror ${url('/gone')}: http 404`);
  expect(await digestOf(dest)).toBe(sha);
}, 30_000);

test('an interrupted download resumes with a Range request and completes', async () => {
  const at = 100_000;
  const { hits, url } = await serve({ '/cut': cut(payload, at) });
  const { status, log, dest } = await fetchArchive([url('/cut')]);
  expect({ status, log }).toEqual({ status: 0, log: expect.stringContaining(`mirror ${url('/cut')}: ok`) });
  expect(hits.length).toBe(2);
  expect(hits[0].range).toBeUndefined();
  expect(hits[1].range).toBeGreaterThan(0);
  expect(hits[1].range).toBeLessThanOrEqual(at);
  expect(await digestOf(dest)).toBe(sha);
}, 30_000);

test('a mirror under the speed floor is abandoned for the next, which resumes the partial file', async () => {
  const { hits, url } = await serve({ '/slow': trickle(payload, 1024, 100), '/good': file(payload) });
  const { status, log, dest } = await fetchArchive([url('/slow'), url('/good')], { LO_FETCH_FLOOR: '100000', LO_FETCH_FLOOR_SECONDS: '1' });
  expect(status).toBe(0);
  expect(log).toContain(`mirror ${url('/slow')}: timed out (below the floor of 100000 B/s`);
  expect(log).toMatch(/: timed out \(below the floor of 100000 B\/s for 1s, or a connect\/TLS stall\) after [1-9]\d* bytes at an average \d+ B\/s/);
  expect(log).toContain(`mirror ${url('/good')}: ok`);
  expect(hits.find(({ path }) => path === '/good')!.range).toBeGreaterThan(0);
  expect(await digestOf(dest)).toBe(sha);
}, 30_000);

test('when every mirror is under the floor the archive host is resumed without one inside the budget', async () => {
  const small = payload.subarray(0, 96 * 1024);
  const smallSha = createHash('sha256').update(small).digest('hex');
  const { url } = await serve({ '/slow-a': trickle(small, 1024, 100), '/slow-b': trickle(small, 1024, 100) });
  root = await mkdtemp(join(tmpdir(), 'fetch-libreoffice-'));
  const dest = join(root, 'LibreOffice.tar.gz');
  const child = Bun.spawn(['bash', script, dest, smallSha, url('/slow-a'), url('/slow-b')], {
    env: { PATH: process.env.PATH!, HOME: root, NO_PROXY: '*', no_proxy: '*', LO_FETCH_PAUSE: '0',
      LO_FETCH_FLOOR: '100000', LO_FETCH_FLOOR_SECONDS: '1', LO_FETCH_BUDGET: '40' },
    stderr: 'pipe',
  });
  const [status, log] = await Promise.all([child.exited, new Response(child.stderr).text()]);
  expect(status).toBe(0);
  expect(log).toContain(`mirror ${url('/slow-a')}: timed out (below the floor`);
  expect(log).toContain(`mirror ${url('/slow-b')}: timed out (below the floor`);
  expect(log).toContain(`mirror ${url('/slow-b')}: last resort without a floor: ok`);
  expect(await digestOf(dest)).toBe(smallSha);
}, 60_000);

test('when no mirror yields the hash the helper fails with one line per mirror and keeps no file', async () => {
  const { url } = await serve({ '/wrong': file(randomBytes(1024)) });
  const { status, log, dest } = await fetchArchive([url('/gone'), url('/wrong')]);
  expect(status).not.toBe(0);
  expect(log).toContain(`mirror ${url('/gone')}: http 404`);
  expect(log).toContain(`mirror ${url('/wrong')}: sha256 mismatch`);
  expect(log).toContain('::error::');
  expect(existsSync(dest)).toBe(false);
  expect(existsSync(`${dest}.part`)).toBe(false);
}, 30_000);

test('a slow host that cannot finish inside the budget fails naming the budget', async () => {
  const { url } = await serve({ '/slow': trickle(payload, 1024, 100) });
  const { status, log } = await fetchArchive([url('/slow')], { LO_FETCH_FLOOR: '100000', LO_FETCH_FLOOR_SECONDS: '1', LO_FETCH_BUDGET: '5' });
  expect(status).not.toBe(0);
  expect(log).toContain(`mirror ${url('/slow')}: timed out (below the floor`);
  expect(log).toContain(`mirror ${url('/slow')}: last resort without a floor: budget of 5s exhausted`);
}, 30_000);

test('an archive already in place with the right hash is not fetched again', async () => {
  const { hits, url } = await serve({ '/good': file(payload) });
  root = await mkdtemp(join(tmpdir(), 'fetch-libreoffice-'));
  await fetchArchive([url('/good')]);
  hits.length = 0;
  const { status, log, dest } = await fetchArchive([url('/good')]);
  expect(status).toBe(0);
  expect(log).toContain('already present');
  expect(hits).toEqual([]);
  expect(await digestOf(dest)).toBe(sha);
}, 30_000);

test('an archive in place with the wrong hash is replaced', async () => {
  const { url } = await serve({ '/good': file(payload) });
  root = await mkdtemp(join(tmpdir(), 'fetch-libreoffice-'));
  const dest = join(root, 'cache', 'LibreOffice.tar.gz');
  await mkdir(join(root, 'cache'), { recursive: true });
  await writeFile(dest, 'stale');
  const { status } = await fetchArchive([url('/good')]);
  expect(status).toBe(0);
  expect(await digestOf(dest)).toBe(sha);
}, 30_000);

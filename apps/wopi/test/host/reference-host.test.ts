import { describe, expect, test } from "bun:test";
import { createProofSigner, dotnetTicks, generateKeyPair, proofKeyAttributes, type ProofSigner } from "../../src/proof";
import { createHostClient, type HostFetch } from "../../src/host/client";
import { WopiHostError } from "../../src/host/errors";
import { encodeUtf7 } from "../../src/host/utf7";
import { createReferenceHost } from "./reference-host";

const SRC = "https://cloud.example.com/wopi/files/file1";
const TOKEN = "refhosttoken123";

function setup(quirks?: { timestampConflict?: boolean; tokenExpired?: boolean; slowBody?: boolean; noContentLength?: boolean }) {
  const pair = generateKeyPair();
  const signer = createProofSigner(pair);
  let nowMs = 1700000000000;
  const host = createReferenceHost({
    verifyProof: proofKeyAttributes(pair.publicKey),
    now: () => nowMs,
    quirks,
  });
  host.tokens.set(TOKEN, { userId: "user1", userName: "User One", canWrite: true });
  host.setFile("file1", new TextEncoder().encode("hello"), { name: "hello.docx" });
  const recorded: Array<{ url: string; method: string; headers: Headers }> = [];
  const fetchFn: HostFetch = async (input, init) => {
    const request = new Request(input as string, init as RequestInit);
    recorded.push({ url: request.url, method: request.method, headers: request.headers });
    return host.fetch(request);
  };
  const client = createHostClient({ signer, maxFileBytes: 10 * 1024 * 1024, fetch: fetchFn, now: () => nowMs });
  return {
    host,
    client,
    signer,
    fetchFn,
    recorded,
    nowMs: () => nowMs,
    advance: (ms: number) => {
      nowMs += ms;
    },
  };
}

function rawRequest(signer: ProofSigner, url: string, token: string, nowMs: number, init: RequestInit = {}): Request {
  const full = new URL(url);
  full.searchParams.set("access_token", token);
  const signed = signer.signProof({ accessToken: token, url: full.toString(), ticks: dotnetTicks(new Date(nowMs)) });
  const headers = new Headers(init.headers);
  headers.set("X-WOPI-Proof", signed.proof);
  headers.set("X-WOPI-ProofOld", signed.proofOld);
  headers.set("X-WOPI-TimeStamp", signed.timestamp);
  return new Request(full.toString(), { ...init, headers });
}

async function kindOf(promise: Promise<unknown>): Promise<WopiHostError> {
  try {
    await promise;
  } catch (err) {
    expect(err).toBeInstanceOf(WopiHostError);
    return err as WopiHostError;
  }
  throw new Error("expected the operation to fail");
}

describe("reference host and host client", () => {
  test("checkFileInfo returns the stored file and proof headers are sent", async () => {
    const { client, recorded } = setup();
    const info = await client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN });
    expect(info.BaseFileName).toBe("hello.docx");
    expect(info.Size).toBe(5);
    expect(info.UserId).toBe("user1");
    expect(info.UserCanWrite).toBe(true);
    expect(info.SupportsLocks).toBe(true);
    const last = recorded[recorded.length - 1];
    expect(last?.headers.get("X-WOPI-Proof")).toBeTruthy();
    expect(last?.headers.get("X-WOPI-TimeStamp")).toBeTruthy();
  });

  test("getFile and readFile return the bytes, putFile increments the item version", async () => {
    const { client, host } = setup();
    const data = await client.readFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(new TextDecoder().decode(data)).toBe("hello");
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    const first = await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new TextEncoder().encode("world!"), lock: "lock-a" });
    expect(host.files.get("file1")?.bytes.length).toBe(6);
    const secondVersion = Number(first.itemVersion);
    expect(secondVersion).toBeGreaterThan(1);
    expect(new TextDecoder().decode(host.files.get("file1")?.bytes)).toBe("world!");
  });

  test("every lock op runs against the host with the right override", async () => {
    const { client, recorded } = setup();
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    expect((await client.getLock({ wopiSrc: SRC, accessToken: TOKEN })).lock).toBe("lock-a");
    await client.refreshLock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    await client.unlockAndRelock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-b", oldLock: "lock-a" });
    expect((await client.getLock({ wopiSrc: SRC, accessToken: TOKEN })).lock).toBe("lock-b");
    await client.unlock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-b" });
    expect((await client.getLock({ wopiSrc: SRC, accessToken: TOKEN })).lock).toBeNull();
    const overrides = recorded.map((entry) => entry.headers.get("X-WOPI-Override")).filter(Boolean);
    expect(overrides).toContain("LOCK");
    expect(overrides).toContain("REFRESH_LOCK");
    expect(overrides).toContain("UNLOCK");
    expect(overrides).toContain("GET_LOCK");
  });

  test("putRelativeFile and renameFile round trip UTF-7 names", async () => {
    const { client, host, recorded } = setup();
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    const copy = await client.putRelativeFile({
      wopiSrc: SRC,
      accessToken: TOKEN,
      suggestedTarget: "Kopie München.docx",
      body: new TextEncoder().encode("copy"),
    });
    expect(copy.name).toContain("Kopie");
    const suggested = recorded.map((entry) => entry.headers.get("X-WOPI-SuggestedTarget")).find(Boolean);
    expect(suggested).toBe(encodeUtf7("Kopie München.docx"));
    const renamed = await client.renameFile({ wopiSrc: SRC, accessToken: TOKEN, requestedName: "Umbenannt", lock: "lock-a" });
    expect(renamed.name).toContain("Umbenannt");
    expect(host.files.get("file1")?.name).toContain("Umbenannt");
  });

  test("a bad proof is rejected with unauthorized", async () => {
    const { host } = setup();
    const otherSigner = createProofSigner(generateKeyPair());
    const badClient = createHostClient({
      signer: otherSigner,
      maxFileBytes: 1024,
      fetch: async (input, init) => host.fetch(new Request(input as string, init as RequestInit)),
      now: () => 1700000000000,
    });
    expect((await kindOf(badClient.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("unauthorized");
  });
});

describe("lock lifecycle", () => {
  test("mismatch reports the current lock, expiry frees the file after 30 minutes", async () => {
    const { client, advance } = setup();
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    const mismatch = await kindOf(client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-b" }));
    expect(mismatch.kind).toBe("lockMismatch");
    expect(mismatch.currentLock).toBe("lock-a");
    const refreshMismatch = await kindOf(client.refreshLock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-b" }));
    expect(refreshMismatch.currentLock).toBe("lock-a");
    advance(31 * 60 * 1000);
    expect((await client.getLock({ wopiSrc: SRC, accessToken: TOKEN })).lock).toBeNull();
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-b" });
    expect((await client.getLock({ wopiSrc: SRC, accessToken: TOKEN })).lock).toBe("lock-b");
  });

  test("putFile without a lock fails on a non empty file and succeeds on an empty one", async () => {
    const { client, host } = setup();
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]) }))).kind).toBe("lockMismatch");
    host.setFile("empty1", new Uint8Array(0), { name: "empty.docx" });
    const ok = await client.putFile({
      wopiSrc: "https://cloud.example.com/wopi/files/empty1",
      accessToken: TOKEN,
      body: new Uint8Array([1, 2]),
    });
    expect(ok.itemVersion).toBeTruthy();
  });

  test("putFile with the holding lock succeeds and with another lock fails", async () => {
    const { client } = setup();
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([5]), lock: "lock-a" });
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([5]), lock: "nope" }))).kind).toBe("lockMismatch");
  });
});

describe("reference host quirks", () => {
  test("timestampConflict surfaces as conflict", async () => {
    const { client } = setup({ timestampConflict: true });
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    const err = await kindOf(
      client.putFile({
        wopiSrc: SRC,
        accessToken: TOKEN,
        body: new Uint8Array([1]),
        lock: "lock-a",
        extraHeaders: { "X-COOL-WOPI-Timestamp": "2000-01-01T00:00:00.000Z" },
      }),
    );
    expect(err.kind).toBe("conflict");
  });

  test("tokenExpired surfaces as unauthorized", async () => {
    const { client } = setup({ tokenExpired: true });
    expect((await kindOf(client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("unauthorized");
  });

  test("noContentLength still reads the file and slowBody still completes", async () => {
    const noLength = setup({ noContentLength: true });
    expect(new TextDecoder().decode(await noLength.client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).toBe("hello");
    const slow = setup({ slowBody: true });
    expect(new TextDecoder().decode(await slow.client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).toBe("hello");
  });
});

describe("reference host strictness", () => {
  test("a proof signed 21 minutes off the host clock is rejected", async () => {
    const { host, signer } = setup();
    const staleClient = createHostClient({
      signer,
      maxFileBytes: 1024,
      fetch: async (input, init) => host.fetch(new Request(input as string, init as RequestInit)),
      now: () => 1700000000000 + 21 * 60 * 1000,
    });
    expect((await kindOf(staleClient.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("unauthorized");
  });

  test("a rotated key is accepted through X-WOPI-ProofOld", async () => {
    const current = generateKeyPair();
    const old = generateKeyPair();
    const untrusted = generateKeyPair();
    const host = createReferenceHost({
      verifyProof: { current: proofKeyAttributes(current.publicKey), old: proofKeyAttributes(old.publicKey) },
      now: () => 1700000000000,
    });
    host.tokens.set(TOKEN, { userId: "user1", canWrite: true });
    host.setFile("file1", new TextEncoder().encode("hi"), { name: "hi.docx" });
    const rotatingSigner = createProofSigner(untrusted, old);
    const client = createHostClient({
      signer: rotatingSigner,
      maxFileBytes: 1024,
      fetch: async (input, init) => host.fetch(new Request(input as string, init as RequestInit)),
      now: () => 1700000000000,
    });
    const info = await client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN });
    expect(info.BaseFileName).toBe("hi.docx");
  });

  test("a viewer token can read but every mutating and lock op is unauthorized", async () => {
    const { host, client, signer, fetchFn } = setup();
    host.tokens.set("viewer-token", { userId: "viewer", userName: "Viewer", canWrite: false });
    const viewer = createHostClient({ signer, maxFileBytes: 1024 * 1024, fetch: fetchFn, now: () => 1700000000000 });
    const info = await viewer.checkFileInfo({ wopiSrc: SRC, accessToken: "viewer-token" });
    expect(info.UserCanWrite).toBe(false);
    expect((await kindOf(viewer.lock({ wopiSrc: SRC, accessToken: "viewer-token", lock: "x" }))).kind).toBe("unauthorized");
    expect((await kindOf(viewer.refreshLock({ wopiSrc: SRC, accessToken: "viewer-token", lock: "x" }))).kind).toBe("unauthorized");
    expect((await kindOf(viewer.unlock({ wopiSrc: SRC, accessToken: "viewer-token", lock: "x" }))).kind).toBe("unauthorized");
    expect((await kindOf(viewer.unlockAndRelock({ wopiSrc: SRC, accessToken: "viewer-token", lock: "y", oldLock: "x" }))).kind).toBe("unauthorized");
    expect((await kindOf(viewer.putFile({ wopiSrc: SRC, accessToken: "viewer-token", body: new Uint8Array([1]) }))).kind).toBe("unauthorized");
    expect((await kindOf(viewer.putRelativeFile({ wopiSrc: SRC, accessToken: "viewer-token", body: new Uint8Array([1]), suggestedTarget: "c.docx" }))).kind).toBe("unauthorized");
    await client.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "lock-a" });
    expect((await kindOf(viewer.renameFile({ wopiSrc: SRC, accessToken: "viewer-token", requestedName: "Nope", lock: "lock-a" }))).kind).toBe("unauthorized");
  });

  test("per token expiresAt expires only that token on the host clock", async () => {
    const second = setup();
    second.host.tokens.set("short-token", { userId: "user2", canWrite: true, expiresAt: 1700000000000 + 10 * 60 * 1000 });
    const shortLived = createHostClient({ signer: second.signer, maxFileBytes: 1024, fetch: second.fetchFn, now: second.nowMs });
    const before = await shortLived.checkFileInfo({ wopiSrc: SRC, accessToken: "short-token" });
    expect(before.UserId).toBe("user2");
    second.advance(11 * 60 * 1000);
    expect((await kindOf(shortLived.checkFileInfo({ wopiSrc: SRC, accessToken: "short-token" }))).kind).toBe("unauthorized");
    const stillValid = await second.client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN });
    expect(stillValid.UserId).toBe("user1");
  });

  test("raw requests enforce PUT override, MaxExpectedSize and lock headers", async () => {
    const { host, signer, nowMs } = setup();
    const noOverride = await host.fetch(rawRequest(signer, `${SRC}/contents`, TOKEN, nowMs(), { method: "POST", body: new Uint8Array([1]) }));
    expect(noOverride.status).toBe(400);
    const wrongOverride = await host.fetch(rawRequest(signer, `${SRC}/contents`, TOKEN, nowMs(), { method: "POST", headers: { "X-WOPI-Override": "LOCK" }, body: new Uint8Array([1]) }));
    expect(wrongOverride.status).toBe(400);
    const tooBig = await host.fetch(rawRequest(signer, `${SRC}/contents`, TOKEN, nowMs(), { method: "GET", headers: { "X-WOPI-MaxExpectedSize": "3" } }));
    expect(tooBig.status).toBe(412);
    const missingLock = await host.fetch(rawRequest(signer, SRC, TOKEN, nowMs(), { method: "POST", headers: { "X-WOPI-Override": "LOCK" } }));
    expect(missingLock.status).toBe(400);
    const missingRequested = await host.fetch(rawRequest(signer, SRC, TOKEN, nowMs(), { method: "POST", headers: { "X-WOPI-Override": "RENAME_FILE" } }));
    expect(missingRequested.status).toBe(400);
    const noTarget = await host.fetch(rawRequest(signer, SRC, TOKEN, nowMs(), { method: "POST", headers: { "X-WOPI-Override": "PUT_RELATIVE" }, body: new Uint8Array([1]) }));
    expect(noTarget.status).toBe(400);
    const bothTargets = await host.fetch(rawRequest(signer, SRC, TOKEN, nowMs(), { method: "POST", headers: { "X-WOPI-Override": "PUT_RELATIVE", "X-WOPI-SuggestedTarget": "a.docx", "X-WOPI-RelativeTarget": "b.docx" }, body: new Uint8Array([1]) }));
    expect(bothTargets.status).toBe(400);
  });

  test("putRelativeFile overwrite respects the target lock and keeps one file entry", async () => {
    const { host, client } = setup();
    host.setFile("doc1", new TextEncoder().encode("v1"), { name: "doc.docx" });
    const docSrc = "https://cloud.example.com/wopi/files/doc1";
    await client.lock({ wopiSrc: docSrc, accessToken: TOKEN, lock: "holder" });
    const err = await kindOf(client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([9]), relativeTarget: "doc.docx", overwrite: true }));
    expect(err.kind).toBe("lockMismatch");
    expect(err.currentLock).toBe("holder");
    const ok = await client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([9]), relativeTarget: "doc.docx", overwrite: true, lock: "holder" });
    expect(ok.name).toBe("doc.docx");
    const named = [...host.files.values()].filter((entry) => entry.name === "doc.docx");
    expect(named.length).toBe(1);
    expect(named[0]?.id).toBe("doc1");
  });
});

describe("reference host file identity and body length", () => {
  test("create, rename, create again keeps both files under distinct ids", async () => {
    const { host, client } = setup();
    const first = await client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new TextEncoder().encode("first"), suggestedTarget: "copy.docx" });
    expect(first.name).toBe("copy.docx");
    const renamed = await client.renameFile({ wopiSrc: first.url, accessToken: TOKEN, requestedName: "renamed" });
    expect(renamed.name).toBe("renamed");
    const second = await client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new TextEncoder().encode("second"), suggestedTarget: "copy.docx" });
    expect(second.name).toBe("copy.docx");
    expect(second.url).not.toBe(first.url);
    const byName = (name: string) => [...host.files.values()].filter((entry) => entry.name === name);
    expect(byName("renamed.docx").length).toBe(1);
    expect(byName("copy.docx").length).toBe(1);
    expect(new TextDecoder().decode(byName("renamed.docx")[0]?.bytes)).toBe("first");
    expect(new TextDecoder().decode(byName("copy.docx")[0]?.bytes)).toBe("second");
    expect(byName("renamed.docx")[0]?.id).not.toBe(byName("copy.docx")[0]?.id);
  });

  test("concurrent copies to one name get distinct names or a conflict", async () => {
    const { host, client } = setup();
    const body = () => new TextEncoder().encode("copy");
    const suggested = await Promise.all([
      client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: body(), suggestedTarget: "race.docx" }),
      client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: body(), suggestedTarget: "race.docx" }),
    ]);
    expect(new Set(suggested.map((copy) => copy.name)).size).toBe(2);
    const relative = await Promise.allSettled([
      client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: body(), relativeTarget: "exact.docx" }),
      client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: body(), relativeTarget: "exact.docx" }),
    ]);
    expect(relative.filter((result) => result.status === "fulfilled").length).toBe(1);
    expect([...host.files.values()].filter((entry) => entry.name === "exact.docx").length).toBe(1);
  });

  test("a body longer than its declared Content-Length is rejected and not stored", async () => {
    const { host, signer, nowMs } = setup();
    let step = 0;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        step += 1;
        if (step > 2) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array(6).fill(65));
      },
    });
    const url = new URL(`${SRC}/contents`);
    url.searchParams.set("access_token", TOKEN);
    const signed = signer.signProof({ accessToken: TOKEN, url: url.toString(), ticks: dotnetTicks(new Date(nowMs())) });
    const request = new Request(url.toString(), {
      method: "POST",
      headers: {
        "X-WOPI-Proof": signed.proof,
        "X-WOPI-ProofOld": signed.proofOld,
        "X-WOPI-TimeStamp": signed.timestamp,
        "X-WOPI-Override": "PUT",
        "Content-Length": "10",
        "Content-Type": "application/octet-stream",
      },
      body: stream,
      duplex: "half",
    } as RequestInit);
    host.setFile("file1", new Uint8Array(0), { name: "hello.docx" });
    expect(host.files.get("file1")?.bytes.length).toBe(0);
    const response = await host.fetch(request);
    expect([400, 413]).toContain(response.status);
    expect(host.files.get("file1")?.bytes.length).toBe(0);
  });
});

describe("streaming through the reference host", () => {
  test("a 64 MiB generated putFile stream is stored without a test side copy", async () => {
    const { host, signer } = setup();
    const total = 64 * 1024 * 1024;
    const chunkSize = 64 * 1024;
    let offset = 0;
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (offset >= total) {
          controller.close();
          return;
        }
        const chunk = new Uint8Array(Math.min(chunkSize, total - offset));
        for (let i = 0; i < chunk.length; i += 1) chunk[i] = (offset + i) % 251;
        offset += chunk.length;
        controller.enqueue(chunk);
      },
    });
    host.setFile("big1", new Uint8Array(0), { name: "big.docx" });
    const bigClient = createHostClient({
      signer,
      maxFileBytes: total,
      fetch: async (input, init) => host.fetch(new Request(input as string, init as RequestInit)),
      now: () => 1700000000000,
    });
    const result = await bigClient.putFile({
      wopiSrc: "https://cloud.example.com/wopi/files/big1",
      accessToken: TOKEN,
      body: source,
      size: total,
    });
    expect(result.itemVersion).toBeTruthy();
    const stored = host.files.get("big1");
    expect(stored?.bytes.length).toBe(total);
    expect(stored?.bytes[123456]).toBe(123456 % 251);
    expect(stored?.bytes[total - 1]).toBe((total - 1) % 251);
  });
});

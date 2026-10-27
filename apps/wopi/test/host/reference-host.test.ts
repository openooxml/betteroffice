import { describe, expect, test } from "bun:test";
import { createProofSigner, generateKeyPair, proofKeyAttributes } from "../../src/proof";
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
    recorded,
    advance: (ms: number) => {
      nowMs += ms;
    },
  };
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

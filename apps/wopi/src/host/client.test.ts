import { describe, expect, test } from "bun:test";
import { createProofSigner, generateKeyPair } from "../proof";
import { createHostClient, type HostFetch } from "./client";
import { WopiHostError, type WopiHostErrorKind } from "./errors";

const SRC = "https://cloud.example.com/wopi/files/file1";
const TOKEN = "DISTINCTIVESECRETTOKEN12345";
const signer = createProofSigner(generateKeyPair());

type Recorded = { url: string; method: string; headers: Headers };

function jsonResponse(body: unknown, init: ResponseInit = {}): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    ...init,
    headers: { "Content-Type": "application/json", ...(init.headers as Record<string, string> | undefined) },
  });
}

function validInfo(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return { BaseFileName: "a.docx", Size: 3, UserId: "user1", UserCanWrite: true, ...overrides };
}

function makeClient(fetchFn: HostFetch, maxFileBytes = 1024 * 1024, extra: { timeoutMs?: number; bodyIdleTimeoutMs?: number } = {}) {
  return createHostClient({ signer, maxFileBytes, fetch: fetchFn, now: () => 1700000000000, ...extra });
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

describe("host client requests", () => {
  test("checkFileInfo signs the exact final URL and validates the result", async () => {
    const recorded: Recorded[] = [];
    const client = makeClient(async (input, init) => {
      const req = new Request(input as string, init as RequestInit);
      recorded.push({ url: req.url, method: req.method, headers: req.headers });
      return jsonResponse(validInfo());
    });
    const info = await client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN });
    expect(info.BaseFileName).toBe("a.docx");
    expect(info.Size).toBe(3);
    expect(recorded[0]?.url).toBe(`${SRC}?access_token=${TOKEN}`);
    expect(recorded[0]?.headers.get("X-WOPI-Proof")).toBeTruthy();
    expect(recorded[0]?.headers.get("X-WOPI-ProofOld")).toBeTruthy();
    expect(recorded[0]?.headers.get("X-WOPI-TimeStamp")).toBeTruthy();
  });

  test("checkFileInfo rejects a missing field and an oversize file", async () => {
    const missing = makeClient(async () => jsonResponse({ Size: 1, UserId: "u" }));
    expect((await kindOf(missing.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("invalidResponse");
    const big = makeClient(async () => jsonResponse(validInfo({ Size: 999999999 })), 100);
    expect((await kindOf(big.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("tooLarge");
  });

  test("maps statuses to typed errors without leaking the token", async () => {
    const cases: Array<[number, WopiHostErrorKind]> = [
      [401, "unauthorized"],
      [404, "notFound"],
      [412, "preconditionFailed"],
      [413, "tooLarge"],
      [501, "notImplemented"],
      [500, "server"],
      [302, "redirect"],
    ];
    for (const [status, kind] of cases) {
      const client = makeClient(async () => new Response(`failed access_token=${TOKEN}`, { status, headers: status === 302 ? { Location: "https://x.example/" } : {} }));
      const err = await kindOf(client.getLock({ wopiSrc: SRC, accessToken: TOKEN }));
      expect(err.kind).toBe(kind);
      expect(err.message).not.toContain(TOKEN);
      expect(String(err)).not.toContain(TOKEN);
    }
  });

  test("maps a lock mismatch with currentLock and reason, and a timestamp body to conflict", async () => {
    const lockClient = makeClient(async () =>
      new Response("locked", { status: 409, headers: { "X-WOPI-Lock": "other-lock", "X-WOPI-LockFailureReason": "locked by another" } }),
    );
    const mismatch = await kindOf(lockClient.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "my-lock" }));
    expect(mismatch.kind).toBe("lockMismatch");
    expect(mismatch.currentLock).toBe("other-lock");
    expect(mismatch.reason).toBe("locked by another");

    const tsClient = makeClient(async () =>
      new Response(JSON.stringify({ COOLStatusCode: 1010 }), { status: 409, headers: { "Content-Type": "application/json" } }),
    );
    const conflict = await kindOf(tsClient.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), lock: "l" }));
    expect(conflict.kind).toBe("conflict");
  });

  test("sends override, lock and contents URL for putFile and returns the item version", async () => {
    const recorded: Recorded[] = [];
    const client = makeClient(async (input, init) => {
      const req = new Request(input as string, init as RequestInit);
      recorded.push({ url: req.url, method: req.method, headers: req.headers });
      return new Response(null, { status: 200, headers: { "X-WOPI-ItemVersion": "7" } });
    });
    const result = await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1, 2]), lock: "lock1", extraHeaders: { "X-COOL-WOPI-Timestamp": "t0" } });
    expect(result.itemVersion).toBe("7");
    expect(recorded[0]?.url).toBe(`${SRC}/contents?access_token=${TOKEN}`);
    expect(recorded[0]?.headers.get("X-WOPI-Override")).toBe("PUT");
    expect(recorded[0]?.headers.get("X-WOPI-Lock")).toBe("lock1");
    expect(recorded[0]?.headers.get("Content-Length")).toBe("2");
    expect(recorded[0]?.headers.get("X-COOL-WOPI-Timestamp")).toBe("t0");
  });
});

describe("retries", () => {
  test("retries without waiting for response cancellation and aborts discarded requests", async () => {
    for (const rejects of [false, true]) {
      let calls = 0;
      let cancellations = 0;
      const signals: Array<AbortSignal | null | undefined> = [];
      const client = makeClient(async (_input, init) => {
        calls += 1;
        if (calls === 3) return jsonResponse(validInfo());
        signals.push(init?.signal);
        return new Response(new ReadableStream<Uint8Array>({
          cancel() {
            cancellations += 1;
            return rejects ? Promise.reject(new Error("cancel failed")) : new Promise<void>(() => {});
          },
        }), { status: 503 });
      });
      expect((await client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN })).BaseFileName).toBe("a.docx");
      expect(calls).toBe(3);
      expect(cancellations).toBe(2);
      expect(signals.every((signal) => signal?.aborted)).toBe(true);
    }
  }, 2000);

  test("retries an idempotent op on 503 and then succeeds", async () => {
    let calls = 0;
    const client = makeClient(async () => {
      calls += 1;
      return calls < 3 ? new Response("busy", { status: 503 }) : jsonResponse(validInfo());
    });
    const info = await client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN });
    expect(info.BaseFileName).toBe("a.docx");
    expect(calls).toBe(3);
  });

  test("retries a network error for getLock but never for putFile or lock", async () => {
    let getLockCalls = 0;
    const getLockClient = makeClient(async () => {
      getLockCalls += 1;
      if (getLockCalls < 3) throw new TypeError("fetch failed");
      return new Response(null, { status: 200, headers: { "X-WOPI-Lock": "abc" } });
    });
    expect((await getLockClient.getLock({ wopiSrc: SRC, accessToken: TOKEN })).lock).toBe("abc");
    expect(getLockCalls).toBe(3);

    let putCalls = 0;
    const putClient = makeClient(async () => {
      putCalls += 1;
      return new Response("busy", { status: 503 });
    });
    expect((await kindOf(putClient.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]) }))).kind).toBe("server");
    expect(putCalls).toBe(1);

    let lockCalls = 0;
    const lockClient = makeClient(async () => {
      lockCalls += 1;
      throw new TypeError("fetch failed");
    });
    expect((await kindOf(lockClient.lock({ wopiSrc: SRC, accessToken: TOKEN, lock: "l" }))).kind).toBe("network");
    expect(lockCalls).toBe(1);
  });

  test("does not retry a 500 for an idempotent op", async () => {
    let calls = 0;
    const client = makeClient(async () => {
      calls += 1;
      return new Response("err", { status: 500 });
    });
    expect((await kindOf(client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("server");
    expect(calls).toBe(1);
  });
});

describe("body cancellation", () => {
  test("successful saves and lock operations do not wait for response cancellation", async () => {
    for (const rejects of [false, true]) {
      let cancellations = 0;
      const signals: Array<AbortSignal | null | undefined> = [];
      const client = makeClient(async (_input, init) => {
        signals.push(init?.signal);
        return new Response(new ReadableStream<Uint8Array>({
          cancel() {
            cancellations += 1;
            return rejects ? Promise.reject(new Error("cancel failed")) : new Promise<void>(() => {});
          },
        }), { headers: { "X-WOPI-ItemVersion": "7", "X-WOPI-Lock": "holder" } });
      });
      const input = { wopiSrc: SRC, accessToken: TOKEN, lock: "holder" };
      expect((await client.putFile({ ...input, body: new Uint8Array([1]) })).itemVersion).toBe("7");
      await client.lock(input);
      await client.refreshLock(input);
      await client.unlock(input);
      await client.unlockAndRelock({ ...input, oldLock: "old" });
      expect((await client.getLock(input)).lock).toBe("holder");
      expect(cancellations).toBe(6);
      expect(signals.every((signal) => signal?.aborted)).toBe(true);
    }
  }, 1000);

  test("cancelling either stream wrapper does not wait for its source", async () => {
    for (const direction of ["download", "upload"]) {
      for (const rejects of [false, true]) {
        let cancelledWith: unknown;
        let signal: AbortSignal | null | undefined;
        const source = new ReadableStream<Uint8Array>({
          cancel(reason) {
            cancelledWith = reason;
            return rejects ? Promise.reject(new Error("cancel failed")) : new Promise<void>(() => {});
          },
        });
        const client = makeClient(async (_input, init) => {
          signal = init?.signal;
          if (direction === "download") return new Response(source);
          await (init?.body as ReadableStream<Uint8Array>).cancel("stop");
          return new Response(null);
        });
        if (direction === "download") {
          const result = await client.getFile({ wopiSrc: SRC, accessToken: TOKEN });
          await result.stream.cancel("stop");
        } else {
          await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source });
        }
        expect(cancelledWith).toBe("stop");
        expect(signal?.aborted).toBe(true);
      }
    }
  }, 1000);
});

describe("streaming", () => {
  test("readFile streams a counted chunked body into one buffer", async () => {
    const payload = new Uint8Array(1024);
    for (let i = 0; i < payload.length; i += 1) payload[i] = i % 251;
    let pulls = 0;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        pulls += 1;
        const start = (pulls - 1) * 128;
        if (start >= payload.length) {
          controller.close();
          return;
        }
        controller.enqueue(payload.subarray(start, start + 128));
      },
    });
    const client = makeClient(async () => new Response(stream, { status: 200, headers: { "Content-Length": String(payload.length), "X-WOPI-ItemVersion": "3" } }));
    const data = await client.readFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(data.length).toBe(1024);
    expect(data[500]).toBe(500 % 251);
    expect(pulls).toBeGreaterThan(1);
  });

  test("getFile returns a stream, size and version without buffering", async () => {
    const client = makeClient(async () =>
      new Response(new Uint8Array([9, 8, 7]), { status: 200, headers: { "Content-Length": "3", "X-WOPI-ItemVersion": "2" } }),
    );
    const result = await client.getFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(result.size).toBe(3);
    expect(result.itemVersion).toBe("2");
    const reader = result.stream.getReader();
    const first = await reader.read();
    expect([...(first.value ?? [])]).toEqual([9, 8, 7]);
  });

  test("readFile enforces the cap mid stream when no length is announced", async () => {
    let cancelled = false;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array(64).fill(1));
      },
      cancel() {
        cancelled = true;
      },
    });
    const client = makeClient(async () => new Response(stream, { status: 200 }), 100);
    expect((await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("tooLarge");
    expect(cancelled).toBe(true);
  });

  test("putFile enforces the cap mid stream", async () => {
    let sourceCancelled = false;
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array(64).fill(2));
      },
      cancel() {
        sourceCancelled = true;
      },
    });
    const client = makeClient(async (_input, init) => {
      const body = (init as RequestInit).body as ReadableStream<Uint8Array>;
      const reader = body.getReader();
      try {
        for (;;) {
          const chunk = await reader.read();
          if (chunk.done) break;
        }
      } catch {
        return new Response(null, { status: 200 });
      }
      return new Response(null, { status: 200 });
    }, 100);
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source }))).kind).toBe("tooLarge");
    expect(sourceCancelled).toBe(true);
  });

  test("download cap failures abort without waiting for source cancellation", async () => {
    for (const announced of [false, true]) {
      let cancelled = false;
      let signal: AbortSignal | null | undefined;
      const stream = new ReadableStream<Uint8Array>({
        pull(controller) {
          controller.enqueue(new Uint8Array(64));
        },
        cancel() {
          cancelled = true;
          return new Promise<void>(() => {});
        },
      });
      const client = makeClient(async (_input, init) => {
        signal = init?.signal;
        return new Response(stream, { headers: announced ? { "Content-Length": "128" } : {} });
      }, 100, { timeoutMs: 5000, bodyIdleTimeoutMs: 5000 });
      const err = await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }));
      expect(err.kind).toBe("tooLarge");
      expect(cancelled).toBe(true);
      expect(signal?.aborted).toBe(true);
    }
  }, 1000);

  test("upload cap failures abort without waiting for source cancellation", async () => {
    for (const op of ["putFile", "putRelativeFile"] as const) {
      let cancelled = false;
      let signal: AbortSignal | null | undefined;
      const source = new ReadableStream<Uint8Array>({
        pull(controller) {
          controller.enqueue(new Uint8Array(64));
        },
        cancel() {
          cancelled = true;
          return new Promise<void>(() => {});
        },
      });
      const client = makeClient(async (_input, init) => {
        signal = init?.signal;
        const reader = (init?.body as ReadableStream<Uint8Array>).getReader();
        try {
          for (;;) {
            if ((await reader.read()).done) break;
          }
        } finally {
          reader.releaseLock();
        }
        return new Response(null);
      }, 100, { timeoutMs: 5000, bodyIdleTimeoutMs: 5000 });
      const err = await kindOf(client[op]({ wopiSrc: SRC, accessToken: TOKEN, body: source, suggestedTarget: "copy.docx" }));
      expect(err.kind).toBe("tooLarge");
      expect(cancelled).toBe(true);
      expect(signal?.aborted).toBe(true);
    }
  }, 1000);

  test("a 64 MiB generated body streams through getFile and readFile by checksum", async () => {
    const total = 64 * 1024 * 1024;
    const chunkSize = 64 * 1024;
    function generatingStream(): ReadableStream<Uint8Array> {
      let offset = 0;
      return new ReadableStream<Uint8Array>({
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
    }
    const getClient = makeClient(async () => new Response(generatingStream(), { status: 200, headers: { "Content-Length": String(total) } }), total);
    const got = await getClient.getFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(got.size).toBe(total);
    let count = 0;
    let sum = 0;
    const reader = got.stream.getReader();
    for (;;) {
      const chunk = await reader.read();
      if (chunk.done) break;
      count += chunk.value.length;
      for (const byte of chunk.value) sum = (sum + byte) % 1000000007;
    }
    expect(count).toBe(total);
    let expected = 0;
    for (let i = 0; i < total; i += 1) expected = (expected + (i % 251)) % 1000000007;
    expect(sum).toBe(expected);

    const readClient = makeClient(async () => new Response(generatingStream(), { status: 200, headers: { "Content-Length": String(total) } }), total);
    const data = await readClient.readFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(data.length).toBe(total);
    expect(data[123456]).toBe(123456 % 251);
    expect(data[total - 1]).toBe((total - 1) % 251);
  });
});

describe("timeouts and body reads", () => {
  test("a stalled file body fails with an idle timeout, not a hang", async () => {
    const stream = new ReadableStream<Uint8Array>({
      pull() {
        return new Promise<void>(() => {});
      },
    });
    const client = makeClient(async () => new Response(stream, { status: 200 }), 1024, { timeoutMs: 5000, bodyIdleTimeoutMs: 20 });
    expect((await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("timeout");
  });

  test("metadata body timeouts cancel the reader, abort and release the body lock", async () => {
    for (const op of ["CheckFileInfo", "PutRelativeFile", "RenameFile", "error"] as const) {
      let cancelled = false;
      let signal: AbortSignal | null | undefined;
      const stream = new ReadableStream<Uint8Array>({
        pull() {
          return new Promise<void>(() => {});
        },
        cancel() {
          cancelled = true;
          return new Promise<void>(() => {});
        },
      });
      const client = makeClient(async (_input, init) => {
        signal = init?.signal;
        return new Response(stream, { status: op === "error" ? 500 : 200 });
      }, 1024, { timeoutMs: 20, bodyIdleTimeoutMs: 60000 });
      const input = { wopiSrc: SRC, accessToken: TOKEN };
      const pending = op === "CheckFileInfo" ? client.checkFileInfo(input)
        : op === "PutRelativeFile" ? client.putRelativeFile({ ...input, body: new Uint8Array([1]), suggestedTarget: "copy.docx" })
        : op === "RenameFile" ? client.renameFile({ ...input, requestedName: "renamed" })
        : client.getLock(input);
      expect((await kindOf(pending)).kind).toBe("timeout");
      expect(cancelled).toBe(true);
      expect(signal?.aborted).toBe(true);
      expect(stream.locked).toBe(false);
    }
  }, 1000);

  test("a steady stream longer than the header timeout still completes", async () => {
    let sent = 0;
    const stream = new ReadableStream<Uint8Array>({
      async pull(controller) {
        await new Promise((resolve) => setTimeout(resolve, 10));
        sent += 1;
        if (sent > 10) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array([sent]));
      },
    });
    const client = makeClient(async () => new Response(stream, { status: 200 }), 1024, { timeoutMs: 30, bodyIdleTimeoutMs: 5000 });
    const data = await client.readFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(data.length).toBe(10);
  });

  test("a mid stream error maps to a redacted WopiHostError", async () => {
    let step = 0;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        step += 1;
        if (step === 1) {
          controller.enqueue(new Uint8Array([1]));
          return;
        }
        controller.error(new Error(`reset https://h/f?access_token=${TOKEN}`));
      },
    });
    const client = makeClient(async () => new Response(stream, { status: 200 }));
    const err = await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }));
    expect(err.kind).toBe("network");
    expect(err.message).not.toContain(TOKEN);
    expect(String(err)).not.toContain(TOKEN);
  });
});

describe("upload header timeout", () => {
  function consumingFetch(onAbort: boolean): HostFetch {
    return async (_input, init) => {
      const body = (init as RequestInit).body as ReadableStream<Uint8Array>;
      const signal = (init as RequestInit).signal;
      const consume = (async () => {
        const reader = body.getReader();
        for (;;) {
          const chunk = await reader.read();
          if (chunk.done) break;
        }
      })();
      if (onAbort && signal) {
        const aborted = new Promise<never>((_resolve, reject) => {
          signal.addEventListener("abort", () => reject(new Error("aborted by signal")));
        });
        try {
          await Promise.race([consume, aborted]);
        } catch (err) {
          if (signal.aborted) throw err;
          return new Response(null, { status: 200, headers: { "X-WOPI-ItemVersion": "4" } });
        }
      } else {
        try {
          await consume;
        } catch {
          return new Response(null, { status: 200, headers: { "X-WOPI-ItemVersion": "4" } });
        }
      }
      return new Response(null, { status: 200, headers: { "X-WOPI-ItemVersion": "4" } });
    };
  }

  test("a slow upload that keeps sending past timeoutMs succeeds", async () => {
    let sent = 0;
    const source = new ReadableStream<Uint8Array>({
      async pull(controller) {
        await new Promise((resolve) => setTimeout(resolve, 15));
        sent += 1;
        if (sent > 10) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array([sent]));
      },
    });
    const client = makeClient(consumingFetch(true), 1024 * 1024, { timeoutMs: 30, bodyIdleTimeoutMs: 5000 });
    const result = await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source, size: 10 });
    expect(result.itemVersion).toBe("4");
    expect(sent).toBeGreaterThan(10);
  });

  test("a stalled upload still times out via the idle timeout", async () => {
    const source = new ReadableStream<Uint8Array>({
      pull() {
        return new Promise<void>(() => {});
      },
    });
    const client = makeClient(consumingFetch(false), 1024 * 1024, { timeoutMs: 5000, bodyIdleTimeoutMs: 20 });
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source }))).kind).toBe("timeout");
  });

  test("a stalled large buffered upload times out instead of hanging", async () => {
    let sawStream = false;
    const stalledFetch: HostFetch = (_input, init) => {
      sawStream = (init as RequestInit).body instanceof ReadableStream;
      const signal = (init as RequestInit).signal;
      return new Promise<Response>((_resolve, reject) => {
        signal?.addEventListener("abort", () => reject(new Error("aborted by signal")));
      });
    };
    const client = makeClient(stalledFetch, 4 * 1024 * 1024, { timeoutMs: 50, bodyIdleTimeoutMs: 5000 });
    const err = await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array(2 * 1024 * 1024) }));
    expect(err.kind).toBe("timeout");
    expect(sawStream).toBe(true);
  });

  test("a source whose first chunk arrives after timeoutMs but within the idle limit succeeds", async () => {
    let pulls = 0;
    const source = new ReadableStream<Uint8Array>({
      async pull(controller) {
        pulls += 1;
        if (pulls === 1) {
          await new Promise((resolve) => setTimeout(resolve, 150));
          controller.enqueue(new Uint8Array([1, 2, 3]));
          return;
        }
        controller.close();
      },
    });
    const client = makeClient(consumingFetch(true), 1024 * 1024, { timeoutMs: 50, bodyIdleTimeoutMs: 5000 });
    const result = await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source, size: 3 });
    expect(result.itemVersion).toBe("4");
  });

  test("a fetch that never pulls the upload stream times out", async () => {
    const neverPullFetch: HostFetch = (_input, init) => {
      const signal = (init as RequestInit).signal;
      return new Promise<Response>((_resolve, reject) => {
        signal?.addEventListener("abort", () => reject(new Error("aborted by signal")));
      });
    };
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array([1]));
      },
    });
    const client = makeClient(neverPullFetch, 1024 * 1024, { timeoutMs: 50, bodyIdleTimeoutMs: 5000 });
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source }))).kind).toBe("timeout");
  });

  test("a fetch that takes one chunk and then never pulls again times out on putFile", async () => {
    let taken: Uint8Array | undefined;
    const stallFetch: HostFetch = async (_input, init) => {
      const body = (init as RequestInit).body as ReadableStream<Uint8Array>;
      const signal = (init as RequestInit).signal;
      const reader = body.getReader();
      const first = await reader.read();
      taken = first.value;
      return new Promise<Response>((_resolve, reject) => {
        signal?.addEventListener("abort", () => reject(new Error("aborted by signal")));
      });
    };
    let step = 0;
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        step += 1;
        if (step > 2) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array([step]));
      },
    });
    const client = makeClient(stallFetch, 1024 * 1024, { timeoutMs: 5000, bodyIdleTimeoutMs: 30 });
    const err = await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source }));
    expect(err.kind).toBe("timeout");
    expect(taken).toEqual(new Uint8Array([1]));
  });

  test("a fetch that takes one chunk and then never pulls again times out on putRelativeFile", async () => {
    const stallFetch: HostFetch = async (_input, init) => {
      const body = (init as RequestInit).body as ReadableStream<Uint8Array>;
      const signal = (init as RequestInit).signal;
      const reader = body.getReader();
      await reader.read();
      return new Promise<Response>((_resolve, reject) => {
        signal?.addEventListener("abort", () => reject(new Error("aborted by signal")));
      });
    };
    let step = 0;
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        step += 1;
        if (step > 2) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array([step]));
      },
    });
    const client = makeClient(stallFetch, 1024 * 1024, { timeoutMs: 5000, bodyIdleTimeoutMs: 30 });
    const err = await kindOf(client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: source, suggestedTarget: "b.docx" }));
    expect(err.kind).toBe("timeout");
  });

  test("a slow source with a prompt consumer keeps its full idle budget per chunk", async () => {
    let sent = 0;
    const source = new ReadableStream<Uint8Array>({
      async pull(controller) {
        await new Promise((resolve) => setTimeout(resolve, 40));
        sent += 1;
        if (sent > 5) {
          controller.close();
          return;
        }
        controller.enqueue(new Uint8Array([sent]));
      },
    });
    const client = makeClient(consumingFetch(true), 1024 * 1024, { timeoutMs: 5000, bodyIdleTimeoutMs: 100 });
    const result = await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source, size: 5 });
    expect(result.itemVersion).toBe("4");
    expect(sent).toBeGreaterThan(5);
  });

  test("a slow but progressing large buffered upload succeeds", async () => {
    let received = 0;
    const slowFetch: HostFetch = async (_input, init) => {
      const body = (init as RequestInit).body as ReadableStream<Uint8Array>;
      const reader = body.getReader();
      for (;;) {
        await new Promise((resolve) => setTimeout(resolve, 5));
        const chunk = await reader.read();
        if (chunk.done) break;
        received += chunk.value.byteLength;
      }
      return new Response(null, { status: 200, headers: { "X-WOPI-ItemVersion": "9" } });
    };
    const client = makeClient(slowFetch, 4 * 1024 * 1024, { timeoutMs: 50, bodyIdleTimeoutMs: 5000 });
    const body = new Uint8Array(2 * 1024 * 1024);
    const result = await client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body });
    expect(result.itemVersion).toBe("9");
    expect(received).toBe(body.byteLength);
  });
});

describe("body caps", () => {
  test("an over cap CheckFileInfo body maps to invalidResponse", async () => {
    const big = { ...validInfo(), padding: "x".repeat(1024 * 1024) };
    const client = makeClient(async () => jsonResponse(big));
    expect((await kindOf(client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("invalidResponse");
  });

  test("an over cap PutRelativeFile body maps to invalidResponse", async () => {
    const big = { Name: "a.docx", Url: "https://cloud.example.com/wopi/files/a", padding: "x".repeat(1024 * 1024) };
    const client = makeClient(async () => jsonResponse(big));
    expect((await kindOf(client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), suggestedTarget: "a.docx" }))).kind).toBe("invalidResponse");
  });

  test("an over cap error body is truncated, cancelled and still maps by status", async () => {
    let cancelled = false;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array(128 * 1024).fill(120));
      },
      cancel() {
        cancelled = true;
      },
    });
    const client = makeClient(async () => new Response(stream, { status: 500 }));
    const err = await kindOf(client.getLock({ wopiSrc: SRC, accessToken: TOKEN }));
    expect(err.kind).toBe("server");
    expect(err.message.length).toBeLessThan(1000);
    expect(cancelled).toBe(true);
  });
});

describe("content length and truncation", () => {
  test("excess download bytes fail without waiting for cancellation and abort the request", async () => {
    for (const rejects of [false, true]) {
      let cancelled = false;
      let signal: AbortSignal | null | undefined;
      const source = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(new Uint8Array([1, 2, 3]));
        },
        cancel() {
          cancelled = true;
          return rejects ? Promise.reject(new Error("cancel failed")) : new Promise<void>(() => {});
        },
      });
      const client = makeClient(async (_input, init) => {
        signal = init?.signal;
        return new Response(source, { headers: { "Content-Length": "1" } });
      });
      expect((await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("invalidResponse");
      expect(cancelled).toBe(true);
      expect(signal?.aborted).toBe(true);
    }
  }, 1000);

  test("writes reject success statuses other than 200", async () => {
    for (const status of [202, 204]) {
      let signal: AbortSignal | null | undefined;
      const client = makeClient(async (_input, init) => {
        signal = init?.signal;
        return new Response(status === 204 ? null : "queued", { status });
      });
      expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]) }))).kind).toBe("invalidResponse");
      expect(signal?.aborted).toBe(true);
      expect((await kindOf(client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), suggestedTarget: ".docx" }))).kind).toBe("invalidResponse");
    }
  });

  test("GetFile rejects partial and empty success statuses", async () => {
    for (const [status, body, headers] of [
      [206, new Uint8Array([1, 2, 3]), { "Content-Length": "3", "Content-Range": "bytes 0-2/10" }],
      [204, null, {}],
    ] as const) {
      let signal: AbortSignal | null | undefined;
      const client = makeClient(async (_input, init) => {
        signal = init?.signal;
        return new Response(body, { status, headers });
      });
      expect((await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("invalidResponse");
      expect((await kindOf(client.getFile({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("invalidResponse");
      expect(signal?.aborted).toBe(true);
    }
  });

  test("readFile fails on premature EOF instead of returning a truncated file", async () => {
    const client = makeClient(async () => new Response(new Uint8Array([1, 2, 3]), { status: 200, headers: { "Content-Length": "10" } }));
    expect((await kindOf(client.readFile({ wopiSrc: SRC, accessToken: TOKEN }))).kind).toBe("invalidResponse");
  });

  test("invalid Content-Length values are treated as unknown", async () => {
    for (const value of ["", "-5", "1.5", "999999999999999999999999"]) {
      const client = makeClient(async () => new Response(new Uint8Array([7]), { status: 200, headers: { "Content-Length": value } }));
      const result = await client.getFile({ wopiSrc: SRC, accessToken: TOKEN });
      expect(result.size).toBeNull();
      const reader = result.stream.getReader();
      await reader.read();
      await reader.cancel();
    }
  });

  test("getFile sends X-WOPI-MaxExpectedSize", async () => {
    const recorded: Recorded[] = [];
    const client = makeClient(async (input, init) => {
      const req = new Request(input as string, init as RequestInit);
      recorded.push({ url: req.url, method: req.method, headers: req.headers });
      return new Response(new Uint8Array([1]), { status: 200, headers: { "Content-Length": "1" } });
    }, 4242);
    await client.getFile({ wopiSrc: SRC, accessToken: TOKEN });
    expect(recorded[0]?.headers.get("X-WOPI-MaxExpectedSize")).toBe("4242");
  });

  test("a WOPISrc fragment is rejected before any fetch or signing", async () => {
    let calls = 0;
    const client = makeClient(async () => {
      calls += 1;
      return jsonResponse(validInfo());
    });
    expect((await kindOf(client.checkFileInfo({ wopiSrc: `${SRC}#frag`, accessToken: TOKEN }))).kind).toBe("notAllowed");
    expect(calls).toBe(0);
  });
});

describe("409 details", () => {
  test("an empty X-WOPI-Lock maps currentLock to null", async () => {
    const client = makeClient(async () =>
      new Response("x", { status: 409, headers: { "X-WOPI-Lock": "", "X-WOPI-LockFailureReason": "not locked" } }),
    );
    const err = await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]) }));
    expect(err.kind).toBe("lockMismatch");
    expect(err.currentLock).toBeNull();
  });

  test("COOLStatusCode past 200 body characters still maps to conflict", async () => {
    const body = JSON.stringify({ padding: "x".repeat(300), COOLStatusCode: 1010 });
    const client = makeClient(async () =>
      new Response(body, { status: 409, headers: { "X-WOPI-Lock": "someone", "Content-Type": "application/json" } }),
    );
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), lock: "l" }))).kind).toBe("conflict");
  });
});

describe("putRelativeFile and renameFile validation", () => {
  test("rejects both targets, sends the lock, and requires a valid JSON reply", async () => {
    const recorded: Recorded[] = [];
    const client = makeClient(async (input, init) => {
      const req = new Request(input as string, init as RequestInit);
      recorded.push({ url: req.url, method: req.method, headers: req.headers });
      return jsonResponse({ Name: "copy.docx", Url: "https://cloud.example.com/wopi/files/copy" });
    });
    const both = await kindOf(client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), suggestedTarget: "a.docx", relativeTarget: "b.docx" }));
    expect(both.kind).toBe("invalidResponse");
    expect(recorded.length).toBe(0);
    const ok = await client.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), relativeTarget: "b.docx", lock: "lock-9" });
    expect(ok.name).toBe("copy.docx");
    expect(ok.url).toContain("copy");
    expect(recorded[0]?.headers.get("X-WOPI-Lock")).toBe("lock-9");

    const htmlClient = makeClient(async () => new Response("<html>nope</html>", { status: 200 }));
    expect((await kindOf(htmlClient.putRelativeFile({ wopiSrc: SRC, accessToken: TOKEN, body: new Uint8Array([1]), suggestedTarget: "a.docx" }))).kind).toBe("invalidResponse");
    const emptyClient = makeClient(async () => new Response(null, { status: 200 }));
    expect((await kindOf(emptyClient.renameFile({ wopiSrc: SRC, accessToken: TOKEN, requestedName: "New" }))).kind).toBe("invalidResponse");
  });

  test("an oversize stream body is cancelled before any fetch", async () => {
    let cancelled = false;
    let calls = 0;
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(new Uint8Array([1]));
      },
      cancel() {
        cancelled = true;
        return new Promise<void>(() => {});
      },
    });
    const client = makeClient(async () => {
      calls += 1;
      return new Response(null, { status: 200 });
    }, 10);
    expect((await kindOf(client.putFile({ wopiSrc: SRC, accessToken: TOKEN, body: source, size: 11 }))).kind).toBe("tooLarge");
    expect(cancelled).toBe(true);
    expect(calls).toBe(0);
  }, 1000);
});

describe("redaction through the client", () => {
  test("a failing op never prints the token to console or the error", async () => {
    const logged: string[] = [];
    const originalLog = console.log;
    const originalError = console.error;
    const originalWarn = console.warn;
    console.log = (...args: unknown[]) => logged.push(args.map(String).join(" "));
    console.error = (...args: unknown[]) => logged.push(args.map(String).join(" "));
    console.warn = (...args: unknown[]) => logged.push(args.map(String).join(" "));
    try {
      const client = makeClient(async () => new Response(`denied access_token=${TOKEN}`, { status: 401 }));
      const err = await kindOf(client.checkFileInfo({ wopiSrc: SRC, accessToken: TOKEN }));
      expect(err.message).not.toContain(TOKEN);
      expect(String(err)).not.toContain(TOKEN);
    } finally {
      console.log = originalLog;
      console.error = originalError;
      console.warn = originalWarn;
    }
    expect(logged.join("\n")).not.toContain(TOKEN);
  });
});
